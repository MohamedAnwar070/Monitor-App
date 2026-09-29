import json
import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import Mock, patch

import pandas as pd
from openpyxl import Workbook
from requests import Response
from selenium.common.exceptions import TimeoutException

from website_monitor import (
    BrowserCaptureResult,
    Config,
    ExcelReporter,
    ScreenshotCapture,
    TemporaryScreenshotUploader,
    URLChecker,
    URLMonitor,
    URLProcessor,
    URLResult,
)


def make_response(status_code: int, payload=None, headers=None) -> Response:
    response = Response()
    response.status_code = status_code
    response.url = "https://example.test/"
    response.headers.update(headers or {})
    if payload is not None:
        response._content = json.dumps(payload).encode("utf-8")
        response.headers["Content-Type"] = "application/json"
    else:
        response._content = b""
    return response


class URLCheckerClassificationTests(unittest.TestCase):
    def setUp(self):
        self.checker = URLChecker(Config(max_retries=2, retry_delay=0))

    def test_normal_http_success(self):
        self.assertEqual(
            self.checker._classify_response(
                "https://example.test/", "frontend", make_response(200)
            ),
            ("SUCCESS", ""),
        )

    def test_authentication_and_protection_prove_reachability(self):
        status_401, note_401 = self.checker._classify_response(
            "https://example.test/", "frontend", make_response(401)
        )
        status_403, note_403 = self.checker._classify_response(
            "https://example.test/", "frontend", make_response(403)
        )

        self.assertEqual(status_401, "SUCCESS")
        self.assertEqual(note_401, "")
        self.assertEqual(status_403, "SUCCESS")
        self.assertEqual(note_403, "")

    def test_backend_root_404_is_reachable(self):
        status, note = self.checker._classify_response(
            "https://api.example.test/",
            "Backend",
            make_response(404, {"message": "Route GET / not found"}),
        )

        self.assertEqual(status, "SUCCESS")
        self.assertEqual(note, "")

    def test_frontend_and_non_root_404_remain_failures(self):
        frontend = self.checker._classify_response(
            "https://example.test/",
            "frontend",
            make_response(404, {"message": "Page not found"}),
        )
        api_route = self.checker._classify_response(
            "https://api.example.test/health",
            "Backend",
            make_response(404, {"message": "Route not found"}),
        )

        self.assertEqual(frontend, ("FAILED", "HTTP 404"))
        self.assertEqual(api_route, ("FAILED", "HTTP 404"))

    def test_expected_codes_enforce_an_exact_contract(self):
        session = Mock()
        session.get.side_effect = [make_response(401), make_response(200)]

        with patch.object(self.checker, "_session", return_value=session):
            accepted = self.checker.check_url(
                "https://example.test/", expected_status_codes=(401,)
            )
            rejected = self.checker.check_url(
                "https://example.test/", expected_status_codes=(401,)
            )

        self.assertEqual(accepted[0:2], ("SUCCESS", 401))
        self.assertEqual(rejected[0:2], ("FAILED", 200))
        self.assertEqual(rejected[3], "HTTP 200 (expected 401)")

    def test_transient_server_errors_are_retried(self):
        session = Mock()
        session.get.side_effect = [make_response(503), make_response(200)]

        with patch.object(self.checker, "_session", return_value=session):
            result = self.checker.check_url("https://example.test/")

        self.assertEqual(result[0:2], ("SUCCESS", 200))
        self.assertEqual(session.get.call_count, 2)


class ExpectedStatusParserTests(unittest.TestCase):
    def test_parses_excel_friendly_status_list(self):
        self.assertEqual(
            URLProcessor._expected_status_codes("204, 200 / 401"),
            (200, 204, 401),
        )

    def test_ignores_empty_and_invalid_values(self):
        self.assertEqual(URLProcessor._expected_status_codes(None), ())
        self.assertEqual(URLProcessor._expected_status_codes("anything"), ())


class ScreenshotCaptureTests(unittest.TestCase):
    def test_final_document_status_uses_post_challenge_response(self):
        driver = Mock()
        driver.get_log.return_value = [
            {
                "message": json.dumps({
                    "message": {
                        "method": "Network.responseReceived",
                        "params": {
                            "type": "Document",
                            "response": {"status": 403},
                        },
                    }
                })
            },
            {
                "message": json.dumps({
                    "message": {
                        "method": "Network.responseReceived",
                        "params": {
                            "type": "Document",
                            "response": {"status": 200},
                        },
                    }
                })
            },
        ]

        self.assertEqual(ScreenshotCapture._final_document_status(driver), 200)

    def test_page_load_timeout_still_captures_rendered_content(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            capture = ScreenshotCapture(temp_dir, pool_size=1)
            driver = Mock()
            driver.get.side_effect = [None, TimeoutException("renderer timeout")]
            driver.current_url = "https://example.test/"
            driver.save_screenshot.side_effect = lambda path: Path(path).write_bytes(b"png")

            def create_thumbnail(_source, destination):
                destination.write_bytes(b"jpeg")

            with (
                patch.object(capture, "_get_driver", return_value=driver),
                patch.object(capture, "_return_driver") as return_driver,
                patch.object(capture, "_create_thumbnail_and_convert", side_effect=create_thumbnail),
                patch("website_monitor.time.sleep"),
            ):
                result = capture.capture("https://example.test/")

            self.assertIsNotNone(result)
            self.assertTrue(Path(result).exists())
            driver.execute_script.assert_any_call("window.stop();")
            return_driver.assert_called_once_with(driver)

    def test_uncommitted_timeout_never_captures_the_previous_site(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            capture = ScreenshotCapture(temp_dir, pool_size=1)
            stale_driver = Mock()
            stale_driver.get.side_effect = [None, TimeoutException("connection timeout")]
            stale_driver.current_url = "about:blank"
            stale_driver.title = "IESA Vision Summit"
            stale_driver.page_source = "<html>previous site</html>"

            fresh_driver = Mock()
            fresh_driver.current_url = "https://example.test/"
            fresh_driver.save_screenshot.side_effect = lambda path: Path(path).write_bytes(b"png")

            def create_thumbnail(_source, destination):
                destination.write_bytes(b"jpeg")

            with (
                patch.object(capture, "_get_driver", side_effect=[stale_driver, fresh_driver]),
                patch.object(capture, "_discard_driver") as discard_driver,
                patch.object(capture, "_return_driver") as return_driver,
                patch.object(capture, "_create_thumbnail_and_convert", side_effect=create_thumbnail),
                patch("website_monitor.time.sleep"),
            ):
                result = capture.capture("https://example.test/")

            self.assertIsNotNone(result)
            stale_driver.save_screenshot.assert_not_called()
            discard_driver.assert_called_once_with(stale_driver)
            return_driver.assert_called_once_with(fresh_driver)

    def test_committed_network_error_captures_the_target_error_page(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            capture = ScreenshotCapture(temp_dir, pool_size=1)
            driver = Mock()
            driver.get.side_effect = [
                None,
                RuntimeError("net::ERR_CONNECTION_TIMED_OUT"),
            ]
            driver.current_url = "https://example.test/"
            driver.title = "example.test"
            driver.page_source = "<html>ERR_CONNECTION_TIMED_OUT</html>"
            driver.save_screenshot.side_effect = lambda path: Path(path).write_bytes(b"png")

            def create_thumbnail(_source, destination):
                destination.write_bytes(b"jpeg")

            with (
                patch.object(capture, "_get_driver", return_value=driver),
                patch.object(capture, "_return_driver") as return_driver,
                patch.object(capture, "_create_thumbnail_and_convert", side_effect=create_thumbnail),
                patch("website_monitor.time.sleep") as sleep_mock,
            ):
                result = capture.capture_result("https://example.test/")

            self.assertEqual(result.state, "browser_error")
            self.assertIsNotNone(result.path)
            sleep_mock.assert_any_call(15)
            driver.save_screenshot.assert_called_once()
            return_driver.assert_called_once_with(driver)

    def test_stuck_renderer_retries_with_a_fresh_driver(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            capture = ScreenshotCapture(temp_dir, pool_size=1)
            stuck_driver = Mock()
            stuck_driver.get.side_effect = [None, TimeoutException("renderer timeout")]
            stuck_driver.current_url = "about:blank"
            stuck_driver.execute_script.side_effect = TimeoutException("renderer stuck")
            fresh_driver = Mock()
            fresh_driver.current_url = "https://example.test/"
            fresh_driver.save_screenshot.side_effect = lambda path: Path(path).write_bytes(b"png")

            def create_thumbnail(_source, destination):
                destination.write_bytes(b"jpeg")

            with (
                patch.object(capture, "_get_driver", side_effect=[stuck_driver, fresh_driver]),
                patch.object(capture, "_discard_driver") as discard_driver,
                patch.object(capture, "_return_driver") as return_driver,
                patch.object(capture, "_create_thumbnail_and_convert", side_effect=create_thumbnail),
                patch("website_monitor.time.sleep"),
            ):
                result = capture.capture("https://example.test/")

            self.assertIsNotNone(result)
            discard_driver.assert_called_once_with(stuck_driver)
            return_driver.assert_called_once_with(fresh_driver)

    def test_browser_failure_does_not_generate_a_fake_screenshot(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            capture = ScreenshotCapture(temp_dir, pool_size=1)

            with patch.object(
                capture,
                "_get_driver",
                side_effect=[RuntimeError("Chrome failed"), RuntimeError("Chrome failed")],
            ):
                result = capture.capture("https://example.test/")

            self.assertIsNone(result)
            self.assertFalse((Path(temp_dir) / "example.test.jpg").exists())

    def test_visually_blank_page_saves_the_real_browser_output(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            capture = ScreenshotCapture(temp_dir, pool_size=1)
            driver = Mock()
            driver.current_url = "https://example.test/"
            driver.save_screenshot.side_effect = lambda path: Path(path).write_bytes(b"png")

            def create_thumbnail(_source, destination):
                destination.write_bytes(b"jpeg")

            with (
                patch.object(capture, "_get_driver", return_value=driver),
                patch.object(capture, "_return_driver") as return_driver,
                patch.object(capture, "_wait_for_browser_challenge", return_value="loaded"),
                patch.object(capture, "_wait_for_rendered_content", return_value=False),
                patch.object(capture, "_has_rendered_content", return_value=False),
                patch.object(capture, "_final_document_status", return_value=200),
                patch.object(capture, "_create_thumbnail_and_convert", side_effect=create_thumbnail),
                patch("website_monitor.time.sleep"),
            ):
                result = capture.capture_result("https://example.test/")

            self.assertEqual(result.state, "blank")
            self.assertEqual(result.status_code, 200)
            self.assertIsNotNone(result.path)
            self.assertTrue(Path(result.path).is_file())
            driver.save_screenshot.assert_called_once()
            return_driver.assert_called_once_with(driver)

    def test_browser_ui_mode_captures_real_desktop_pixels(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            capture = ScreenshotCapture(
                temp_dir,
                pool_size=1,
                capture_browser_ui=True,
            )
            driver = Mock()
            driver.current_url = "https://example.test/"
            driver.get_window_rect.return_value = {
                "x": 0,
                "y": 0,
                "width": 1366,
                "height": 768,
            }

            def save_desktop(destination, _region, window_handle=None):
                destination.write_bytes(b"desktop png")

            def create_thumbnail(_source, destination):
                destination.write_bytes(b"desktop jpeg")

            with (
                patch.object(capture, "_get_driver", return_value=driver),
                patch.object(capture, "_discard_driver") as discard_driver,
                patch.object(capture, "_focus_browser_window", return_value=1234) as focus_browser,
                patch.object(capture, "_save_desktop_screenshot", side_effect=save_desktop) as save_screen,
                patch.object(capture, "_create_thumbnail_and_convert", side_effect=create_thumbnail),
                patch("website_monitor.time.sleep"),
            ):
                result = capture.capture("https://example.test/")

            self.assertIsNotNone(result)
            self.assertTrue(Path(result).is_file())
            focus_browser.assert_called_once_with(driver)
            save_screen.assert_called_once()
            self.assertEqual(
                save_screen.call_args.args[1],
                driver.get_window_rect.return_value,
            )
            self.assertEqual(save_screen.call_args.kwargs["window_handle"], 1234)
            driver.save_screenshot.assert_not_called()
            discard_driver.assert_called_once_with(driver)

    def test_browser_ui_auth_dialog_can_capture_before_url_commit(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            capture = ScreenshotCapture(
                temp_dir,
                pool_size=1,
                capture_browser_ui=True,
            )
            driver = Mock()
            driver.current_url = "about:blank"
            driver.get.side_effect = [
                None,
                RuntimeError("net::ERR_INVALID_AUTH_CREDENTIALS"),
            ]
            driver.get_window_rect.return_value = {
                "x": 0,
                "y": 0,
                "width": 1366,
                "height": 768,
            }

            def save_desktop(destination, _region, window_handle=None):
                destination.write_bytes(b"auth dialog png")

            def create_thumbnail(_source, destination):
                destination.write_bytes(b"auth dialog jpeg")

            with (
                patch.object(capture, "_get_driver", return_value=driver),
                patch.object(capture, "_discard_driver") as discard_driver,
                patch.object(capture, "_focus_browser_window") as focus_browser,
                patch.object(capture, "_save_desktop_screenshot", side_effect=save_desktop),
                patch.object(capture, "_create_thumbnail_and_convert", side_effect=create_thumbnail),
                patch("website_monitor.time.sleep"),
            ):
                result = capture.capture_result("https://auth.example.test/")

            self.assertIsNotNone(result.path)
            self.assertTrue(Path(result.path).is_file())
            focus_browser.assert_called_once_with(driver)
            discard_driver.assert_called_once_with(driver)


class ConfigDefaultsTests(unittest.TestCase):
    def test_auth_browser_ui_screenshots_are_enabled_by_default(self):
        self.assertTrue(Config().auth_browser_ui_screenshot)


class URLProcessorScreenshotTests(unittest.TestCase):
    def test_backend_401_uses_real_browser_ui_capture(self):
        screenshot_capture = Mock()
        auth_screenshot_capture = Mock()
        auth_screenshot_capture.capture.return_value = "screenshots/api.example.test.jpg"
        processor = URLProcessor(
            Config(),
            screenshot_capture=screenshot_capture,
            auth_screenshot_capture=auth_screenshot_capture,
        )
        row = pd.Series({
            "Group ID": 1,
            "Host": "api.example.test",
            "Type": "Backend",
            "Target URL": "https://api.example.test/",
        })

        with (
            patch.object(
                processor.url_checker,
                "check_url",
                return_value=("SUCCESS", 401, 10.0, ""),
            ),
            patch.object(
                processor.ssl_checker,
                "check_ssl",
                return_value=("Valid", "2027-01-01", 100),
            ),
        ):
            result = processor.process_url(0, row)

        screenshot_capture.capture.assert_not_called()
        auth_screenshot_capture.capture.assert_called_once_with("https://api.example.test/")
        self.assertEqual(result.screenshot_path, "screenshots/api.example.test.jpg")

    def test_docker_registry_root_uses_v2_auth_screenshot(self):
        screenshot_capture = Mock()
        auth_screenshot_capture = Mock()
        auth_screenshot_capture.capture.return_value = "screenshots/repo.example.test.jpg"
        processor = URLProcessor(
            Config(),
            screenshot_capture=screenshot_capture,
            auth_screenshot_capture=auth_screenshot_capture,
        )
        row = pd.Series({
            "Group ID": 1,
            "Host": "repo.example.test",
            "Type": "frontend",
            "Target URL": "https://repo.example.test/",
        })

        with (
            patch.object(
                processor.url_checker,
                "check_url",
                return_value=("SUCCESS", 200, 10.0, ""),
            ),
            patch.object(
                processor.ssl_checker,
                "check_ssl",
                return_value=("Valid", "2027-01-01", 100),
            ),
            patch.object(
                processor,
                "_registry_screenshot_target",
                return_value=("https://repo.example.test/v2/", 401),
            ),
        ):
            result = processor.process_url(0, row)

        screenshot_capture.capture.assert_not_called()
        auth_screenshot_capture.capture.assert_called_once_with(
            "https://repo.example.test/v2/"
        )
        self.assertEqual(result.target_url, "https://repo.example.test/")
        self.assertEqual(result.status_code, 200)
        self.assertEqual(result.screenshot_path, "screenshots/repo.example.test.jpg")


class ScreenshotPublishingTests(unittest.TestCase):
    def test_real_capture_gets_stable_github_link(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as temp_dir:
            monitor = URLMonitor.__new__(URLMonitor)
            monitor.config = Config(screenshot_upload_provider="github")
            monitor.run_dt = datetime(2026, 7, 28, 10, 0, 0)
            monitor.session_label = "Morning"
            monitor.sharepoint_client = None
            monitor.screenshot_capture = ScreenshotCapture(temp_dir, pool_size=1)
            screenshot_path = Path(temp_dir) / "api.example.test.jpg"
            screenshot_path.write_bytes(b"real screenshot")
            result = URLResult(
                index=0,
                group_id="1",
                host="api.example.test",
                type_="Backend",
                target_url="https://api.example.test/",
                status="SUCCESS",
                status_code=401,
                response_time=10.0,
                failure_reason="",
                ssl_status="Valid",
                ssl_expiry="2027-01-01",
                ssl_days_left=100,
                screenshot_path=str(screenshot_path),
            )

            with patch.dict(
                os.environ,
                {"GITHUB_REPOSITORY": "owner/repo", "GITHUB_REF_NAME": "main"},
            ):
                monitor._publish_screenshot_links([result])
                monitor._validate_screenshot_results([result])

            self.assertTrue(Path(result.screenshot_path).is_file())
            self.assertEqual(
                result.screenshot_url,
                "https://raw.githubusercontent.com/owner/repo/main/"
                + Path(result.screenshot_path).resolve().relative_to(Path.cwd().resolve()).as_posix(),
            )

    def test_failed_url_without_capture_is_excluded_from_coverage(self):
        monitor = URLMonitor.__new__(URLMonitor)
        monitor.config = Config(with_screenshot=True)
        result = URLResult(
            index=0,
            group_id="1",
            host="example.test",
            type_="frontend",
            target_url="https://example.test/",
            status="FAILED",
            status_code=None,
            response_time=None,
            failure_reason="Connection Failed",
            ssl_status="Invalid",
            ssl_expiry="",
            ssl_days_left=None,
        )

        monitor._validate_screenshot_results([result])

    def test_missing_capture_warns_by_default(self):
        monitor = URLMonitor.__new__(URLMonitor)
        monitor.config = Config(with_screenshot=True)
        result = URLResult(
            index=0,
            group_id="1",
            host="example.test",
            type_="frontend",
            target_url="https://example.test/",
            status="SUCCESS",
            status_code=200,
            response_time=10.0,
            failure_reason="",
            ssl_status="Valid",
            ssl_expiry="2027-01-01",
            ssl_days_left=100,
        )

        with self.assertLogs("website_monitor", level="WARNING") as logs:
            monitor._validate_screenshot_results([result])

        self.assertIn("Screenshot coverage incomplete", "\n".join(logs.output))

    def test_missing_capture_is_rejected_in_strict_mode(self):
        monitor = URLMonitor.__new__(URLMonitor)
        monitor.config = Config(with_screenshot=True, strict_screenshot_coverage=True)
        result = URLResult(
            index=0,
            group_id="1",
            host="example.test",
            type_="frontend",
            target_url="https://example.test/",
            status="SUCCESS",
            status_code=200,
            response_time=10.0,
            failure_reason="",
            ssl_status="Valid",
            ssl_expiry="2027-01-01",
            ssl_days_left=100,
        )

        with self.assertRaisesRegex(RuntimeError, "Screenshot coverage incomplete"):
            monitor._validate_screenshot_results([result])

    def test_tmpfiles_failure_does_not_fall_back_to_private_github_url(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            screenshot_path = Path(temp_dir) / "example.test.jpg"
            screenshot_path.write_bytes(b"real screenshot")
            monitor = URLMonitor.__new__(URLMonitor)
            monitor.config = Config(
                screenshot_upload_provider="tmpfiles",
                strict_screenshot_coverage=True,
            )
            monitor.run_dt = datetime(2026, 7, 28, 10, 0, 0)
            monitor.session_label = "Morning"
            monitor.sharepoint_client = None
            result = URLResult(
                index=0,
                group_id="1",
                host="example.test",
                type_="frontend",
                target_url="https://example.test/",
                status="SUCCESS",
                status_code=200,
                response_time=10.0,
                failure_reason="",
                ssl_status="Valid",
                ssl_expiry="2027-01-01",
                ssl_days_left=100,
                screenshot_path=str(screenshot_path),
            )

            with (
                patch.object(TemporaryScreenshotUploader, "upload", return_value=None),
                patch.dict(
                    os.environ,
                    {"GITHUB_REPOSITORY": "private/repo", "GITHUB_REF_NAME": "main"},
                ),
            ):
                monitor._publish_screenshot_links([result])

            self.assertIsNone(result.screenshot_url)
            with self.assertRaisesRegex(RuntimeError, "Screenshot coverage incomplete"):
                monitor._validate_screenshot_results([result])


class BrowserVerificationTests(unittest.TestCase):
    def setUp(self):
        self.processor = URLProcessor(Config())

    def test_resolved_frontend_challenge_uses_final_browser_200(self):
        result = self.processor._apply_browser_verification(
            endpoint_type="frontend",
            expected_status_codes=(),
            status="SUCCESS",
            status_code=403,
            failure_reason="",
            browser_result=BrowserCaptureResult(
                path="shot.jpg",
                status_code=200,
                state="loaded",
            ),
        )

        self.assertEqual(result, ("SUCCESS", 200, ""))

    def test_unresolved_frontend_challenge_is_not_a_false_success(self):
        result = self.processor._apply_browser_verification(
            endpoint_type="frontend",
            expected_status_codes=(),
            status="SUCCESS",
            status_code=403,
            failure_reason="",
            browser_result=BrowserCaptureResult(
                path="challenge.jpg",
                status_code=403,
                state="challenge",
            ),
        )

        self.assertEqual(
            result,
            ("FAILED", 403, "Browser challenge did not resolve"),
        )

    def test_backend_403_keeps_endpoint_semantics(self):
        result = self.processor._apply_browser_verification(
            endpoint_type="backend",
            expected_status_codes=(),
            status="SUCCESS",
            status_code=403,
            failure_reason="",
            browser_result=BrowserCaptureResult(
                path="access-denied.jpg",
                status_code=403,
                state="access_denied",
            ),
        )

        self.assertEqual(result, ("SUCCESS", 403, ""))


class ExcelReporterFormattingTests(unittest.TestCase):
    def test_red_is_limited_to_confirmed_failure_fields(self):
        workbook = Workbook()
        sheet = workbook.active
        reporter = ExcelReporter()
        row = {column: "" for column in reporter.REPORT_COLUMNS}
        row.update({
            "Target URL": "https://example.test/",
            "Status": "FAILED",
            "Status Code": 503,
            "Failure Reason": "HTTP 503",
        })

        reporter._write_section(sheet, "Afternoon", [row])

        data_row = 3
        red_columns = {"Status", "Status Code", "Failure Reason"}
        for column_index, column_name in enumerate(reporter.REPORT_COLUMNS, start=1):
            color = sheet.cell(data_row, column_index).font.color
            rgb = color.rgb[-6:] if color and color.type == "rgb" else None
            if column_name in red_columns:
                self.assertEqual(rgb, "FF0000")
            else:
                self.assertNotEqual(rgb, "FF0000")

    def test_report_has_only_requested_columns(self):
        self.assertEqual(
            ExcelReporter.REPORT_COLUMNS,
            [
                "Group ID",
                "Host",
                "Type",
                "Target URL",
                "Status",
                "Status Code",
                "Failure Reason",
                "SSL Status",
                "SSL Expiry",
                "SSL Days Left",
                "Screenshot",
            ],
        )

    def test_old_verbose_reasons_are_cleaned_when_sheet_is_rewritten(self):
        reporter = ExcelReporter()
        self.assertEqual(
            reporter._concise_failure_reason({
                "Status": "FAILED",
                "Status Code": 503,
                "Failure Reason": "HTTP 503 after 4 attempts: service unavailable",
            }),
            "HTTP 503",
        )
        self.assertEqual(
            reporter._concise_failure_reason({
                "Status": "SUCCESS",
                "Status Code": 401,
                "Failure Reason": "Endpoint reachable; authentication is required",
            }),
            "",
        )
        self.assertEqual(
            reporter._concise_failure_reason({
                "Status": "FAILED",
                "Status Code": 403,
                "Failure Reason": "Browser challenge did not resolve",
            }),
            "HTTP 403 - Browser challenge did not resolve",
        )


if __name__ == "__main__":
    unittest.main()
