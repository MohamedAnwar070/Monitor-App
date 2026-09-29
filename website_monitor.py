import logging
import json
import os
import queue
import re
import socket
import ssl
import threading
import time
import base64
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple
from urllib.parse import urljoin, urlparse, quote

import pandas as pd
import requests
from OpenSSL import crypto
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from concurrent.futures import ThreadPoolExecutor, as_completed
from PIL import Image as PILImage
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common.exceptions import TimeoutException

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(f'url_monitor_{datetime.now().strftime("%Y%m%d")}.log')
    ]
)
logger = logging.getLogger(__name__)


# Global switches
WITH_SCREENSHOT = os.getenv("MONITOR_WITH_SCREENSHOT", "1").strip().lower() not in {"0", "false", "no"}
DEFAULT_SHAREPOINT_WORKBOOK_URL = os.getenv("SHAREPOINT_WORKBOOK_URL", "")


@dataclass
class Config:
    """Simplified configuration"""
    max_workers: int = 10
    timeout: int = 20
    max_retries: int = 4
    retry_delay: float = 2.0
    ssl_timeout: int = 5
    ssl_warning_days: int = 30
    with_screenshot: bool = True
    input_file: str = "Monitor Endpoints.xlsx"
    report_file: str = "reports/Website_Monitor_Report.xlsx"
    screenshot_online_base_url: str = ""
    screenshot_upload_provider: str = "tmpfiles"
    screenshot_link_expire_seconds: int = 172800
    sharepoint_upload_screenshots: bool = True
    auth_browser_ui_screenshot: bool = True
    strict_screenshot_coverage: bool = False
    retry_status_codes: Tuple[int, ...] = (408, 425, 429, 500, 502, 503, 504)
    protected_status_codes: Tuple[int, ...] = (401, 403)
    backend_type_names: Tuple[str, ...] = ("api", "backend")
    max_retry_wait_seconds: float = 30.0


@dataclass
class URLResult:
    """Result with clear SUCCESS/FAILED status"""
    index: int
    group_id: str
    host: str
    type_: str
    target_url: str
    status: str  # Only "SUCCESS" or "FAILED"
    status_code: Optional[int]
    response_time: Optional[float]
    failure_reason: str  # Details why it failed
    ssl_status: str
    ssl_expiry: str
    ssl_days_left: Optional[int]
    screenshot_path: Optional[str] = None
    screenshot_url: Optional[str] = None
    checked_at: str = ""
    session_label: str = ""


@dataclass
class BrowserCaptureResult:
    """Screenshot plus the final browser-visible navigation result."""
    path: Optional[str]
    status_code: Optional[int]
    state: str
    detail: str = ""


class URLChecker:
    """HTTP availability checker using endpoint semantics instead of host rules."""
    
    def __init__(self, config: Config):
        self.config = config
        self._thread_local = threading.local()

    def _session(self) -> requests.Session:
        """Return one pooled Session per worker thread."""
        session = getattr(self._thread_local, "session", None)
        if session is not None:
            return session

        session = requests.Session()
        session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"
            ),
            "Accept": "text/html,application/json;q=0.9,*/*;q=0.8",
        })
        adapter = requests.adapters.HTTPAdapter(
            pool_connections=1,
            pool_maxsize=2,
            max_retries=0,
        )
        session.mount("http://", adapter)
        session.mount("https://", adapter)
        self._thread_local.session = session
        return session

    def _is_backend(self, endpoint_type: str) -> bool:
        normalized = endpoint_type.strip().lower()
        return any(name in normalized for name in self.config.backend_type_names)

    @staticmethod
    def _is_root_url(url: str) -> bool:
        return urlparse(url).path in {"", "/"}

    def _classify_response(
        self,
        url: str,
        endpoint_type: str,
        response: requests.Response,
    ) -> Tuple[str, str]:
        status_code = response.status_code

        if 200 <= status_code < 400:
            return "SUCCESS", ""

        if status_code in self.config.protected_status_codes:
            return "SUCCESS", ""

        if (
            self._is_backend(endpoint_type)
            and self._is_root_url(url)
            and 400 <= status_code < 500
            and status_code not in self.config.retry_status_codes
        ):
            return "SUCCESS", ""

        return "FAILED", f"HTTP {status_code}"

    def _retry_wait(self, response: requests.Response, attempt: int) -> float:
        wait_seconds = self.config.retry_delay * (attempt + 1)
        retry_after = response.headers.get("Retry-After", "").strip()
        if retry_after.isdigit():
            wait_seconds = max(wait_seconds, float(retry_after))
        return min(wait_seconds, self.config.max_retry_wait_seconds)
    
    def check_url(
        self,
        url: str,
        endpoint_type: str = "",
        expected_status_codes: Tuple[int, ...] = (),
    ) -> Tuple[str, Optional[int], Optional[float], str]:
        """
        Check URL and return clear SUCCESS/FAILED status
        
        Returns:
            (status, status_code, response_time_ms, failure_reason)
            status is either "SUCCESS" or "FAILED"
        """
        for attempt in range(self.config.max_retries):
            try:
                start_time = time.time()
                response = self._session().get(
                    url,
                    timeout=self.config.timeout,
                    verify=True,
                    allow_redirects=True
                )
                response_time = round((time.time() - start_time) * 1000, 2)

                if response.status_code in expected_status_codes:
                    return (
                        "SUCCESS",
                        response.status_code,
                        response_time,
                        "",
                    )

                if response.status_code in self.config.retry_status_codes:
                    if attempt >= self.config.max_retries - 1:
                        return (
                            "FAILED",
                            response.status_code,
                            response_time,
                            f"HTTP {response.status_code}",
                        )
                    wait_seconds = self._retry_wait(response, attempt)
                    logger.info(
                        "Retrying %s after transient HTTP %s in %.1fs",
                        url,
                        response.status_code,
                        wait_seconds,
                    )
                    time.sleep(wait_seconds)
                    continue

                if expected_status_codes:
                    expected = ", ".join(str(code) for code in expected_status_codes)
                    reason = f"HTTP {response.status_code} (expected {expected})"
                    return "FAILED", response.status_code, response_time, reason

                status, reason = self._classify_response(url, endpoint_type, response)
                return status, response.status_code, response_time, reason
                
            except requests.exceptions.Timeout:
                if attempt < self.config.max_retries - 1:
                    time.sleep(self.config.retry_delay * (attempt + 1))
                    continue
                return "FAILED", None, None, "Timeout"
                
            except requests.exceptions.SSLError:
                return "FAILED", None, None, "SSL Error"
                
            except requests.exceptions.ConnectionError:
                if attempt < self.config.max_retries - 1:
                    time.sleep(self.config.retry_delay * (attempt + 1))
                    continue
                return "FAILED", None, None, "Connection Failed"
                
            except Exception:
                if attempt < self.config.max_retries - 1:
                    time.sleep(self.config.retry_delay)
                    continue
                return "FAILED", None, None, "Request Error"
                
        return "FAILED", None, None, "Max retries exceeded"


class SSLChecker:
    """Simplified SSL checker"""
    
    @staticmethod
    def check_ssl(hostname: str, config: Config) -> Tuple[str, str, Optional[int]]:
        """
        Check SSL certificate
        
        Returns:
            (status, expiry_date, days_left)
        """
        try:
            context = ssl.create_default_context()
            
            with socket.create_connection((hostname, 443), timeout=config.ssl_timeout) as sock:
                with context.wrap_socket(sock, server_hostname=hostname) as secure_sock:
                    cert_der = secure_sock.getpeercert(binary_form=True)
                    x509_cert = crypto.load_certificate(crypto.FILETYPE_ASN1, cert_der)
                    
                    expiry_str = x509_cert.get_notAfter().decode('ascii')
                    expiry_date = datetime.strptime(expiry_str, '%Y%m%d%H%M%SZ')
                    days_left = (expiry_date - datetime.utcnow()).days
                    
                    return "Valid", expiry_date.strftime("%Y-%m-%d"), days_left
                    
        except Exception as e:
            logger.debug(f"SSL check failed for {hostname}: {e}")
            return "Invalid", "", None


class ScreenshotCapture:
    """Capture webpage or full browser/desktop screenshots using Chrome.
    
    Uses a thread-safe queue pool of WebDriver instances so multiple threads can
    capture screenshots in parallel. This avoids spawning one WebDriver per thread,
    significantly reducing memory and CPU load, preventing timeouts.
    """

    VIEWPORT_WIDTH = 1280
    VIEWPORT_HEIGHT = 800
    BROWSER_UI_WIDTH = 1366
    BROWSER_UI_HEIGHT = 768
    PAGE_LOAD_TIMEOUT = 30  # seconds
    THUMB_WIDTH = 1024
    THUMB_HEIGHT = 640
    CHALLENGE_TIMEOUT = 12
    CHALLENGE_POLL_INTERVAL = 0.5
    RENDER_TIMEOUT = 12
    RENDER_POLL_INTERVAL = 0.25
    RENDER_SETTLE_TIME = 1.0
    SCREENSHOT_DELAY_SECONDS = 15
    CHALLENGE_MARKERS = (
        "checking your browser before accessing",
        "please wait for up to 5 seconds",
        "verify you are human",
        "performing security verification",
    )
    ACCESS_DENIED_MARKERS = (
        "403 forbidden",
        "access denied",
        "you don't have permission to access",
        "request blocked",
    )
    BROWSER_ERROR_MARKERS = (
        "err_connection_",
        "err_name_not_resolved",
        "err_timed_out",
        "err_address_unreachable",
        "err_internet_disconnected",
    )

    def __init__(
        self,
        output_dir: str = "screenshots",
        pool_size: int = 6,
        capture_browser_ui: bool = False,
    ):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self._pool_size = pool_size
        self.capture_browser_ui = capture_browser_ui
        self._capture_lock = threading.Lock()
        self._driver_queue = queue.Queue()
        self._drivers_lock = threading.Lock()
        self._all_drivers: list = []  # track all drivers for cleanup

    # --- WebDriver lifecycle (connection pool) ---

    def _get_driver(self) -> webdriver.Chrome:
        """Get a WebDriver from the pool queue, creating a new one if limit not reached."""
        try:
            # Try to get an idle driver immediately without blocking
            return self._driver_queue.get_nowait()
        except queue.Empty:
            # If no idle driver, check if we can spawn a new one under pool_size
            with self._drivers_lock:
                if len(self._all_drivers) < self._pool_size:
                    driver = self._create_driver()
                    self._all_drivers.append(driver)
                    return driver
            
            # Pool is full and busy; block until a driver is returned
            return self._driver_queue.get(block=True, timeout=120)

    def _return_driver(self, driver: webdriver.Chrome) -> None:
        """Return a WebDriver instance back to the pool queue."""
        self._driver_queue.put(driver)

    def _discard_driver(self, driver: webdriver.Chrome) -> None:
        """Remove an unresponsive driver so a later attempt can create a clean one."""
        with self._drivers_lock:
            if driver in self._all_drivers:
                self._all_drivers.remove(driver)
        try:
            driver.quit()
        except Exception:
            pass

    def _create_driver(self) -> webdriver.Chrome:
        opts = ChromeOptions()
        if not self.capture_browser_ui:
            opts.add_argument("--headless=new")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--disable-extensions")
        window_width = self.BROWSER_UI_WIDTH if self.capture_browser_ui else self.VIEWPORT_WIDTH
        window_height = self.BROWSER_UI_HEIGHT if self.capture_browser_ui else self.VIEWPORT_HEIGHT
        opts.add_argument(f"--window-size={window_width},{window_height}")
        opts.add_argument("--window-position=0,0")
        if not self.capture_browser_ui:
            opts.add_argument("--hide-scrollbars")
        opts.add_argument("--ignore-certificate-errors")
        opts.add_argument("--log-level=3")  # suppress Chrome console noise
        # ChromeDriver enables Chrome's automation infobar by default. Exclude
        # that switch so full-browser screenshots do not contain the
        # "controlled by automated test software" banner.
        opts.add_experimental_option(
            "excludeSwitches",
            ["enable-automation", "enable-logging"],
        )
        opts.add_experimental_option("useAutomationExtension", False)
        opts.set_capability("goog:loggingPrefs", {"performance": "ALL"})
        opts.page_load_strategy = 'eager'  # Only wait for DOMContentLoaded (interactive state)

        try:
            # Try system-installed chromedriver first (GitHub Actions)
            service = ChromeService(log_output=os.devnull)
            driver = webdriver.Chrome(service=service, options=opts)
        except Exception:
            try:
                from webdriver_manager.chrome import ChromeDriverManager
                service = ChromeService(
                    ChromeDriverManager().install(),
                    log_output=os.devnull
                )
                driver = webdriver.Chrome(service=service, options=opts)
            except Exception as e:
                logger.error(f"Failed to create Chrome WebDriver: {e}")
                raise

        driver.set_page_load_timeout(ScreenshotCapture.PAGE_LOAD_TIMEOUT)
        if self.capture_browser_ui:
            driver.set_window_position(0, 0)
            driver.set_window_size(self.BROWSER_UI_WIDTH, self.BROWSER_UI_HEIGHT)
        return driver

    @staticmethod
    def _windows_browser_handle(driver: webdriver.Chrome) -> Optional[int]:
        """Resolve the top-level Chrome window owned by this WebDriver session."""
        if os.name != "nt":
            return None

        try:
            import ctypes
            from ctypes import wintypes

            chrome_pid = int(driver.capabilities.get("goog:processID", 0) or 0)
            if not chrome_pid:
                return None

            matching_windows = []
            enum_callback = ctypes.WINFUNCTYPE(
                wintypes.BOOL,
                wintypes.HWND,
                wintypes.LPARAM,
            )

            @enum_callback
            def collect_window(hwnd, _lparam):
                process_id = wintypes.DWORD()
                ctypes.windll.user32.GetWindowThreadProcessId(
                    hwnd,
                    ctypes.byref(process_id),
                )
                if (
                    process_id.value == chrome_pid
                    and ctypes.windll.user32.IsWindowVisible(hwnd)
                ):
                    rect = wintypes.RECT()
                    ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
                    area = max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)
                    matching_windows.append((area, int(hwnd)))
                return True

            ctypes.windll.user32.EnumWindows(collect_window, 0)
            if matching_windows:
                return max(matching_windows)[1]
        except Exception as handle_error:
            logger.debug("Could not resolve the Chrome window handle: %s", handle_error)
        return None

    def _focus_browser_window(self, driver: webdriver.Chrome) -> Optional[int]:
        """Bring headed Chrome forward and return its native Windows handle."""
        try:
            driver.switch_to.window(driver.current_window_handle)
            driver.set_window_position(0, 0)
            driver.set_window_size(self.BROWSER_UI_WIDTH, self.BROWSER_UI_HEIGHT)
        except Exception:
            pass

        window_handle = self._windows_browser_handle(driver)
        if not window_handle:
            return None
        try:
            import ctypes

            ctypes.windll.user32.ShowWindow(window_handle, 9)  # SW_RESTORE
            ctypes.windll.user32.BringWindowToTop(window_handle)
            ctypes.windll.user32.SetForegroundWindow(window_handle)
        except Exception as focus_error:
            logger.debug("Could not foreground Chrome before capture: %s", focus_error)
        return window_handle

    # --- Screenshot capture ---

    @staticmethod
    def _short_error(e: Exception) -> str:
        """Extract just the meaningful first line from Selenium exceptions."""
        msg = str(e)
        # Selenium puts the useful message before 'Stacktrace:'
        if 'Stacktrace:' in msg:
            msg = msg[:msg.index('Stacktrace:')].strip()
        # Take only the first line if multi-line
        first_line = msg.split('\n')[0].strip()
        return first_line[:120]

    @staticmethod
    def _filename_for_url(url: str, filename_hint: str = "") -> str:
        """Return a safe, stable screenshot filename for a URL or row hint."""
        parsed = urlparse(str(url or ""))
        hostname = parsed.hostname
        if not hostname and url:
            hostname = urlparse(f"//{str(url).strip()}").hostname
        name = hostname or filename_hint or "unknown-site"
        name = re.sub(r"[^A-Za-z0-9._-]+", "_", str(name).strip().strip("/"))
        return f"{name or 'unknown-site'}.jpg"

    @staticmethod
    def _page_text(driver: webdriver.Chrome) -> str:
        raw_title = driver.title
        raw_source = driver.page_source
        title = raw_title if isinstance(raw_title, str) else ""
        source = raw_source if isinstance(raw_source, str) else ""
        return f"{title}\n{source}".lower()

    @classmethod
    def _has_marker(cls, page_text: str, markers: Tuple[str, ...]) -> bool:
        return any(marker in page_text for marker in markers)

    def _wait_for_browser_challenge(self, driver: webdriver.Chrome) -> str:
        """Wait for a JavaScript/CDN browser challenge and classify the final page."""
        deadline = time.monotonic() + self.CHALLENGE_TIMEOUT
        page_text = self._page_text(driver)

        while (
            self._has_marker(page_text, self.CHALLENGE_MARKERS)
            and time.monotonic() < deadline
        ):
            time.sleep(self.CHALLENGE_POLL_INTERVAL)
            page_text = self._page_text(driver)

        raw_current_url = driver.current_url
        current_url = raw_current_url if isinstance(raw_current_url, str) else ""
        if current_url.startswith("chrome-error://"):
            return "browser_error"
        if self._has_marker(page_text, self.BROWSER_ERROR_MARKERS):
            return "browser_error"
        if self._has_marker(page_text, self.CHALLENGE_MARKERS):
            return "challenge"
        if self._has_marker(page_text, self.ACCESS_DENIED_MARKERS):
            return "access_denied"
        if not page_text.strip():
            return "blank"
        return "loaded"

    @staticmethod
    def _has_rendered_content(driver: webdriver.Chrome) -> bool:
        """Return whether the document has text or a painted visual element."""
        return bool(driver.execute_script(
            """
            const body = document.body;
            const root = document.getElementById('root');
            const visual = document.querySelector('img,svg,canvas,video');
            return Boolean(
                body &&
                ['interactive', 'complete'].includes(document.readyState) &&
                (
                    body.innerText.trim().length > 0 ||
                    (root && root.childElementCount > 0) ||
                    (visual && visual.getBoundingClientRect().width > 0 &&
                        visual.getBoundingClientRect().height > 0)
                )
            );
            """
        ))

    def _wait_for_rendered_content(self, driver: webdriver.Chrome) -> bool:
        """Wait until a static page or JavaScript app has painted useful content."""
        deadline = time.monotonic() + self.RENDER_TIMEOUT

        while time.monotonic() < deadline:
            try:
                rendered = self._has_rendered_content(driver)
                if rendered:
                    # Allow fonts, styles, and the first application paint to settle.
                    time.sleep(self.RENDER_SETTLE_TIME)
                    return True
            except Exception:
                pass
            time.sleep(self.RENDER_POLL_INTERVAL)

        logger.warning(
            "Page did not render visible content within %ss: %s",
            self.RENDER_TIMEOUT,
            driver.current_url,
        )
        return False

    @staticmethod
    def _final_document_status(driver: webdriver.Chrome) -> Optional[int]:
        """Read the last top-level document response from Chrome's performance log."""
        try:
            entries = driver.get_log("performance")
        except Exception:
            return None
        if not isinstance(entries, list):
            return None

        final_status = None
        for entry in entries:
            try:
                message = json.loads(entry["message"])["message"]
                if (
                    message.get("method") == "Network.responseReceived"
                    and message.get("params", {}).get("type") == "Document"
                ):
                    status = message["params"]["response"].get("status")
                    if status is not None:
                        final_status = int(status)
            except (KeyError, TypeError, ValueError, json.JSONDecodeError):
                continue
        return final_status

    @staticmethod
    def _current_url(driver: webdriver.Chrome) -> str:
        """Return Chrome's current URL without leaking WebDriver mock/object values."""
        raw_url = driver.current_url
        return raw_url.strip() if isinstance(raw_url, str) else ""

    def _prepare_for_navigation(self, driver: webdriver.Chrome) -> None:
        """Clear the prior site before a pooled browser is sent to another URL."""
        driver.get("about:blank")
        try:
            # Discard both the previous site's events and the about:blank event.
            driver.get_log("performance")
        except Exception:
            pass

    @classmethod
    def _navigation_committed(cls, driver: webdriver.Chrome) -> bool:
        """Whether the requested navigation replaced the pool's blank reset page."""
        current_url = cls._current_url(driver).lower()
        return bool(current_url) and not current_url.startswith(("about:blank", "data:"))

    def capture_result(self, url: str) -> BrowserCaptureResult:
        """Capture *url* and return screenshot and browser navigation verification."""
        if self.capture_browser_ui:
            with self._capture_lock:
                return self._capture_result_with_retries(url)
        return self._capture_result_with_retries(url)

    def _capture_result_with_retries(self, url: str) -> BrowserCaptureResult:
        """Capture with a fresh Chrome retry and never synthesize an image."""
        last_error: Optional[Exception] = None

        parsed = urlparse(url)
        hostname = parsed.hostname or "unknown"
        filename = self._filename_for_url(url)
        full_path = self.output_dir / filename

        for attempt in range(2):
            driver = None
            driver_is_healthy = False
            try:
                # A rerun of the same session must never leave an older image behind
                # when the replacement capture fails.
                if full_path.exists():
                    full_path.unlink()

                driver = self._get_driver()
                capture_region = None
                self._prepare_for_navigation(driver)
                try:
                    driver.get(url)
                except TimeoutException:
                    # Heavy pages may keep loading analytics or large assets forever.
                    # Keep rendered target content only after the new navigation has
                    # actually replaced about:blank. Otherwise a pooled driver could
                    # save the previous website under this URL's filename.
                    logger.info("Page load timed out for %s; checking rendered content", url)
                    if not self.capture_browser_ui:
                        driver.execute_script("window.stop();")
                    if not self._navigation_committed(driver):
                        raise RuntimeError(
                            "Page load timed out before the target navigation committed"
                        )
                except Exception as navigation_error:
                    if (
                        not self.capture_browser_ui
                        and not self._navigation_committed(driver)
                    ):
                        raise
                    # Chrome network errors and native authentication dialogs can
                    # interrupt WebDriver while still rendering a page that belongs
                    # to this target. The about:blank reset makes that safe to keep.
                    logger.info(
                        "Chrome navigation reported %s for %s; capturing committed page",
                        self._short_error(navigation_error),
                        url,
                    )

                if (
                    not self.capture_browser_ui
                    and not self._navigation_committed(driver)
                ):
                    raise RuntimeError("Target navigation did not replace the blank page")

                page_state = self._wait_for_browser_challenge(driver)
                rendered = self._wait_for_rendered_content(driver)

                # Give late JavaScript, images, fonts, and redirects a fixed window
                # to finish before saving the pixels for this URL.
                logger.info(
                    "Waiting %ss before screenshot for %s",
                    self.SCREENSHOT_DELAY_SECONDS,
                    url,
                )
                time.sleep(self.SCREENSHOT_DELAY_SECONDS)
                # A challenge or redirect may have resolved during the delay.
                page_state = self._wait_for_browser_challenge(driver)
                if not rendered:
                    try:
                        rendered = self._has_rendered_content(driver)
                    except Exception:
                        rendered = False
                if not rendered and not self.capture_browser_ui:
                    # A blank viewport is still the real result of visiting the
                    # endpoint. Save those pixels instead of turning a healthy
                    # HTTP check into a missing-artifact failure. Browser-based
                    # 403 verification still treats this state as a failed load.
                    logger.warning(
                        "Page remained visually blank after the screenshot delay; "
                        "capturing the browser output for %s",
                        url,
                    )
                    if page_state == "loaded":
                        page_state = "blank"
                browser_status = self._final_document_status(driver)

                temp_png = self.output_dir / (
                    f"temp_{hostname}_{threading.get_ident()}_{attempt}.png"
                )

                if self.capture_browser_ui:
                    window_handle = self._focus_browser_window(driver)
                    capture_region = driver.get_window_rect()
                    time.sleep(0.5)
                    self._save_desktop_screenshot(
                        temp_png,
                        capture_region,
                        window_handle=window_handle,
                    )
                else:
                    driver.save_screenshot(str(temp_png))
                self._create_thumbnail_and_convert(temp_png, full_path)
                if not full_path.exists() or full_path.stat().st_size == 0:
                    raise RuntimeError("Screenshot file was not created")

                # Desktop/auth browsers are discarded after each capture so a modal
                # prompt can never leak into the next URL.
                driver_is_healthy = not self.capture_browser_ui
                logger.debug(f"Screenshot saved: {full_path}")
                return BrowserCaptureResult(
                    path=str(full_path),
                    status_code=browser_status,
                    state=page_state,
                )

            except Exception as e:
                last_error = e
                if full_path.exists():
                    try:
                        full_path.unlink()
                    except OSError:
                        pass
                if attempt == 0:
                    logger.info("Retrying screenshot for %s with a fresh browser", url)
            finally:
                if driver is not None:
                    if driver_is_healthy:
                        self._return_driver(driver)
                    else:
                        self._discard_driver(driver)

        logger.warning(
            "Screenshot failed for %s: %s",
            url,
            self._short_error(last_error or RuntimeError("Unknown capture error")),
        )
        return BrowserCaptureResult(
            path=None,
            status_code=None,
            state="capture_failed",
            detail=self._short_error(last_error or RuntimeError("Unknown capture error")),
        )

    def capture(self, url: str) -> Optional[str]:
        """Backward-compatible screenshot-only API."""
        return self.capture_result(url).path

    @staticmethod
    def _save_desktop_screenshot(
        destination: Path,
        region: Optional[dict] = None,
        window_handle: Optional[int] = None,
    ) -> None:
        """Capture real desktop pixels for the Chrome window, including native UI."""
        if os.name == "nt" and window_handle:
            try:
                import ctypes
                import win32con
                import win32gui
                import win32ui

                left, top, right, bottom = win32gui.GetWindowRect(window_handle)
                width = right - left
                height = bottom - top
                window_dc = win32gui.GetWindowDC(window_handle)
                source_dc = win32ui.CreateDCFromHandle(window_dc)
                memory_dc = source_dc.CreateCompatibleDC()
                bitmap = win32ui.CreateBitmap()
                bitmap.CreateCompatibleBitmap(source_dc, width, height)
                memory_dc.SelectObject(bitmap)
                try:
                    captured = ctypes.windll.user32.PrintWindow(
                        window_handle,
                        memory_dc.GetSafeHdc(),
                        2,  # PW_RENDERFULLCONTENT
                    )
                    if not captured:
                        raise RuntimeError("PrintWindow did not capture Chrome")
                    bitmap_info = bitmap.GetInfo()
                    bitmap_bytes = bitmap.GetBitmapBits(True)
                    image = PILImage.frombuffer(
                        "RGB",
                        (bitmap_info["bmWidth"], bitmap_info["bmHeight"]),
                        bitmap_bytes,
                        "raw",
                        "BGRX",
                        0,
                        1,
                    )
                    image.save(destination, "PNG")
                    return
                finally:
                    win32gui.DeleteObject(bitmap.GetHandle())
                    memory_dc.DeleteDC()
                    source_dc.DeleteDC()
                    win32gui.ReleaseDC(window_handle, window_dc)
            except Exception as window_capture_error:
                logger.warning(
                    "Direct Chrome window capture failed; using screen pixels: %s",
                    window_capture_error,
                )

        import mss

        with mss.MSS() as screen:
            monitor = (
                {
                    "left": int(region["x"]),
                    "top": int(region["y"]),
                    "width": int(region["width"]),
                    "height": int(region["height"]),
                }
                if region
                else screen.monitors[0]
            )
            pixels = screen.grab(monitor)
            image = PILImage.frombytes("RGB", pixels.size, pixels.rgb)
            image.save(destination, "PNG")

    def _create_thumbnail_and_convert(self, src_path: Path, dest_path: Path) -> None:
        """Resize the screenshot to high-res thumbnail and save as optimized JPEG."""
        try:
            with PILImage.open(src_path) as img:
                # Convert RGBA/P to RGB for JPEG compatibility
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                img.thumbnail((self.THUMB_WIDTH, self.THUMB_HEIGHT), PILImage.LANCZOS)
                img.save(dest_path, "JPEG", quality=85, optimize=True)
        except Exception as e:
            logger.warning(f"Thumbnail generation/JPEG conversion failed: {e}")
        finally:
            # Clean up temporary PNG file
            if src_path.exists():
                try:
                    src_path.unlink()
                except Exception:
                    pass

    def quit_all_drivers(self) -> None:
        """Quit ALL WebDriver instances. Call once after processing."""
        with self._drivers_lock:
            for driver in self._all_drivers:
                try:
                    driver.quit()
                except Exception:
                    pass
            self._all_drivers.clear()
            # Clear queue
            while not self._driver_queue.empty():
                try:
                    self._driver_queue.get_nowait()
                except queue.Empty:
                    break


class SharePointGraphClient:
    """Small Microsoft Graph client for one SharePoint/OneDrive workbook."""

    GRAPH_ROOT = "https://graph.microsoft.com/v1.0"

    def __init__(self, tenant_id: str, client_id: str, client_secret: str, workbook_url: str):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.workbook_url = workbook_url
        self._token: Optional[str] = None
        self._drive_item: Optional[dict] = None

    @classmethod
    def from_env(cls) -> Optional["SharePointGraphClient"]:
        tenant_id = os.getenv("SHAREPOINT_TENANT_ID", "").strip()
        client_id = os.getenv("SHAREPOINT_CLIENT_ID", "").strip()
        client_secret = os.getenv("SHAREPOINT_CLIENT_SECRET", "").strip()
        explicit_workbook_url = os.getenv("SHAREPOINT_WORKBOOK_URL", "").strip()
        workbook_url = explicit_workbook_url or DEFAULT_SHAREPOINT_WORKBOOK_URL

        credential_values = [tenant_id, client_id, client_secret]
        if not any(credential_values) and not workbook_url:
            return None
        if not all(credential_values) or not workbook_url:
            raise RuntimeError(
                "SharePoint sync requires SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, "
                "SHAREPOINT_CLIENT_SECRET, and SHAREPOINT_WORKBOOK_URL."
            )
        return cls(tenant_id, client_id, client_secret, workbook_url)

    def _get_token(self) -> str:
        if self._token:
            return self._token

        response = requests.post(
            f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token",
            data={
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "grant_type": "client_credentials",
                "scope": "https://graph.microsoft.com/.default",
            },
            timeout=30,
        )
        response.raise_for_status()
        self._token = response.json()["access_token"]
        return self._token

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        headers = kwargs.pop("headers", {})
        headers["Authorization"] = f"Bearer {self._get_token()}"
        response = requests.request(method, url, headers=headers, timeout=60, **kwargs)
        response.raise_for_status()
        return response

    @staticmethod
    def _share_id(url: str) -> str:
        encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii").rstrip("=")
        return f"u!{encoded}"

    def resolve_workbook(self) -> dict:
        if self._drive_item:
            return self._drive_item

        share_id = self._share_id(self.workbook_url)
        response = self._request("GET", f"{self.GRAPH_ROOT}/shares/{share_id}/driveItem")
        self._drive_item = response.json()
        return self._drive_item

    def download_workbook(self, destination: Path) -> bool:
        try:
            item = self.resolve_workbook()
            drive_id = item["parentReference"]["driveId"]
            item_id = item["id"]
            response = self._request("GET", f"{self.GRAPH_ROOT}/drives/{drive_id}/items/{item_id}/content")
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(response.content)
            logger.info("Downloaded SharePoint workbook to %s", destination)
            return True
        except Exception as e:
            logger.warning("Could not download SharePoint workbook, using local copy/new workbook: %s", e)
            return False

    def upload_workbook(self, source: Path) -> bool:
        try:
            item = self.resolve_workbook()
            drive_id = item["parentReference"]["driveId"]
            item_id = item["id"]
            self._request(
                "PUT",
                f"{self.GRAPH_ROOT}/drives/{drive_id}/items/{item_id}/content",
                data=source.read_bytes(),
                headers={"Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
            )
            logger.info("Uploaded report workbook to SharePoint")
            return True
        except Exception as e:
            logger.error("Could not upload workbook to SharePoint: %s", e)
            return False

    def upload_related_file(self, source: Path, relative_path: str) -> Optional[str]:
        """Upload a screenshot beside the workbook and return an online URL."""
        try:
            workbook = self.resolve_workbook()
            drive_id = workbook["parentReference"]["driveId"]
            folder_id = workbook["parentReference"]["id"]
            parts = [p for p in Path(relative_path).parts if p]
            if not parts:
                return None

            for folder_name in parts[:-1]:
                folder_id = self._ensure_child_folder(drive_id, folder_id, folder_name)

            filename = parts[-1]
            response = self._request(
                "PUT",
                f"{self.GRAPH_ROOT}/drives/{drive_id}/items/{folder_id}:/{quote(filename)}:/content",
                data=source.read_bytes(),
                headers={"Content-Type": "image/jpeg"},
            )
            uploaded = response.json()
            return self._create_view_link(drive_id, uploaded["id"]) or uploaded.get("webUrl")
        except Exception as e:
            logger.warning("Could not upload screenshot %s to SharePoint: %s", source, e)
            return None

    def _ensure_child_folder(self, drive_id: str, parent_id: str, folder_name: str) -> str:
        encoded_name = quote(folder_name)
        try:
            response = self._request(
                "GET",
                f"{self.GRAPH_ROOT}/drives/{drive_id}/items/{parent_id}:/{encoded_name}",
            )
            return response.json()["id"]
        except requests.HTTPError as e:
            if e.response is None or e.response.status_code != 404:
                raise

        response = self._request(
            "POST",
            f"{self.GRAPH_ROOT}/drives/{drive_id}/items/{parent_id}/children",
            json={
                "name": folder_name,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "replace",
            },
        )
        return response.json()["id"]

    def _create_view_link(self, drive_id: str, item_id: str) -> Optional[str]:
        scope = os.getenv("SHAREPOINT_LINK_SCOPE", "organization").strip() or "organization"
        try:
            response = self._request(
                "POST",
                f"{self.GRAPH_ROOT}/drives/{drive_id}/items/{item_id}/createLink",
                json={"type": "view", "scope": scope},
            )
            return response.json().get("link", {}).get("webUrl")
        except Exception as e:
            logger.debug("Could not create SharePoint view link: %s", e)
            return None


class TemporaryScreenshotUploader:
    """Upload screenshots to a public temporary file host for manual Excel sharing."""

    TMPFILES_API = "https://tmpfiles.org/api/v1/upload"

    def __init__(self, expire_seconds: int = 604800):
        self.expire_seconds = expire_seconds

    def upload(self, source: Path) -> Optional[str]:
        last_error: Optional[Exception] = None
        for attempt in range(3):
            try:
                with source.open("rb") as file_obj:
                    response = requests.post(
                        self.TMPFILES_API,
                        files={"file": (source.name, file_obj, "image/jpeg")},
                        data={"expire": str(self.expire_seconds)},
                        timeout=60,
                    )
                response.raise_for_status()
                uploaded_url = response.json().get("data", {}).get("url")
                if not uploaded_url:
                    raise RuntimeError("Upload response did not include a URL")
                return uploaded_url.replace(
                    "https://tmpfiles.org/",
                    "https://tmpfiles.org/dl/",
                )
            except Exception as e:
                last_error = e
                if attempt < 2:
                    time.sleep(2 ** attempt)

        logger.warning(
            "Could not upload screenshot %s after 3 attempts: %s",
            source,
            last_error,
        )
        return None


class URLProcessor:
    """Process URLs with clear results"""
    
    def __init__(
        self,
        config: Config,
        screenshot_capture: Optional["ScreenshotCapture"] = None,
        auth_screenshot_capture: Optional["ScreenshotCapture"] = None,
        run_dt: Optional[datetime] = None,
        session_label: str = "",
    ):
        self.config = config
        self.url_checker = URLChecker(config)
        self.ssl_checker = SSLChecker()
        self.screenshot_capture = screenshot_capture
        self.auth_screenshot_capture = auth_screenshot_capture
        self.run_dt = run_dt or datetime.now()
        self.session_label = session_label

    @staticmethod
    def _expected_status_codes(value: object) -> Tuple[int, ...]:
        """Parse an optional Excel value such as `200, 204` or `401`."""
        if value is None or pd.isna(value):
            return ()
        codes = {
            int(match)
            for match in re.findall(r"(?<!\d)([1-5]\d{2})(?!\d)", str(value))
        }
        return tuple(sorted(codes))

    def _registry_screenshot_target(
        self,
        url: str,
        endpoint_type: str,
    ) -> Tuple[str, Optional[int]]:
        """Use the Docker Registry API root when a registry website root is empty."""
        parsed = urlparse(url)
        hostname = (parsed.hostname or "").lower()
        first_label = hostname.split(".", 1)[0]
        looks_like_registry = (
            first_label in {"repo", "registry", "docker"}
            or "registry" in endpoint_type.lower()
        )
        if not looks_like_registry or parsed.path not in {"", "/"}:
            return url, None

        registry_url = urljoin(f"{parsed.scheme}://{parsed.netloc}/", "v2/")
        try:
            response = self.url_checker._session().get(
                registry_url,
                timeout=self.config.timeout,
                verify=True,
                allow_redirects=True,
            )
            api_version = response.headers.get(
                "Docker-Distribution-Api-Version",
                "",
            ).lower()
            if api_version.startswith("registry/"):
                logger.info(
                    "Docker Registry detected at %s; using %s for screenshot",
                    url,
                    registry_url,
                )
                return registry_url, response.status_code
        except requests.RequestException as registry_error:
            logger.debug("Docker Registry probe failed for %s: %s", url, registry_error)
        return url, None

    def _apply_browser_verification(
        self,
        endpoint_type: str,
        expected_status_codes: Tuple[int, ...],
        status: str,
        status_code: Optional[int],
        failure_reason: str,
        browser_result: Optional[BrowserCaptureResult],
    ) -> Tuple[str, Optional[int], str]:
        """Use Chrome as the source of truth when a frontend probe hits a CDN 403."""
        needs_verification = (
            status_code == 403
            and not expected_status_codes
            and not self.url_checker._is_backend(endpoint_type)
            and browser_result is not None
        )
        if not needs_verification:
            return status, status_code, failure_reason

        browser_status = browser_result.status_code
        if (
            browser_result.state == "loaded"
            and browser_status is not None
            and 200 <= browser_status < 400
        ):
            return "SUCCESS", browser_status, ""

        if browser_result.state == "loaded" and browser_status is None:
            # Older Chrome/driver combinations may not expose performance logs.
            # The rendered non-challenge page is still stronger evidence than the raw probe.
            return "SUCCESS", status_code, "Browser verified; HTTP probe was blocked"

        reason_by_state = {
            "challenge": "Browser challenge did not resolve",
            "access_denied": "Browser showed access denied",
            "browser_error": "Browser could not load the page",
            "blank": "Browser loaded a blank page",
            "capture_failed": "Browser verification failed",
        }
        reason = reason_by_state.get(browser_result.state, "Browser verification failed")
        return "FAILED", browser_status or status_code, reason
    
    def process_url(self, idx: int, row: pd.Series) -> URLResult:
        """Process single URL and return clear result"""
        raw_url = ""
        try:
            # Normalize URL
            raw_url = str(row.get("Target URL", "")).strip()
            if not raw_url:
                return self._create_failed_result(idx, row, raw_url, "Empty URL")
            
            url = raw_url if raw_url.startswith(('http://', 'https://')) else f"https://{raw_url}"
            hostname = urlparse(url).hostname
            
            if not hostname:
                return self._create_failed_result(idx, row, url, "Invalid URL")
            
            # Check URL
            endpoint_type = str(row.get("Type", "")).strip()
            expected_status_codes = self._expected_status_codes(
                row.get("Expected Status Codes", None)
            )
            status, status_code, response_time, failure_reason = self.url_checker.check_url(
                url,
                endpoint_type=endpoint_type,
                expected_status_codes=expected_status_codes,
            )
            
            # Check SSL for HTTPS
            ssl_status, ssl_expiry, ssl_days = "N/A", "", None
            if url.startswith("https://"):
                ssl_status, ssl_expiry, ssl_days = self.ssl_checker.check_ssl(hostname, self.config)
            
            # A dedicated headed Chrome + desktop capture is used for 401 pages so
            # the real native Basic Auth dialog appears in the saved screenshot.
            screenshot_path = None
            browser_result = None
            if self.screenshot_capture:
                screenshot_url, screenshot_status = self._registry_screenshot_target(
                    url,
                    endpoint_type,
                )
                capture = (
                    self.auth_screenshot_capture
                    if (
                        (status_code == 401 or screenshot_status == 401)
                        and self.auth_screenshot_capture
                    )
                    else self.screenshot_capture
                )
                if hasattr(capture, "capture_result"):
                    res = capture.capture_result(screenshot_url)
                    if isinstance(res, BrowserCaptureResult):
                        browser_result = res
                        screenshot_path = browser_result.path
                    else:
                        screenshot_path = capture.capture(screenshot_url)
                else:
                    screenshot_path = capture.capture(screenshot_url)

            status, status_code, failure_reason = self._apply_browser_verification(
                endpoint_type=endpoint_type,
                expected_status_codes=expected_status_codes,
                status=status,
                status_code=status_code,
                failure_reason=failure_reason,
                browser_result=browser_result,
            )
            
            return URLResult(
                index=idx,
                group_id=row.get("Group ID", ""),
                host=row.get("Host", ""),
                type_=row.get("Type", ""),
                target_url=url,
                status=status,
                status_code=status_code,
                response_time=response_time,
                failure_reason=failure_reason,
                ssl_status=ssl_status,
                ssl_expiry=ssl_expiry,
                ssl_days_left=ssl_days,
                screenshot_path=screenshot_path,
                checked_at=self.run_dt.strftime("%Y-%m-%d %H:%M:%S"),
                session_label=self.session_label,
            )
            
        except Exception as e:
            logger.error(f"Error processing URL at index {idx}: {e}")
            return self._create_failed_result(idx, row, raw_url, "Processing Error")
    
    def _create_failed_result(self, idx: int, row: pd.Series, url: str, reason: str) -> URLResult:
        """Helper to create failed result"""
        return URLResult(
            index=idx,
            group_id=row.get("Group ID", ""),
            host=row.get("Host", ""),
            type_=row.get("Type", ""),
            target_url=url,
            status="FAILED",
            status_code=None,
            response_time=None,
            failure_reason=reason,
            ssl_status="N/A",
            ssl_expiry="",
            ssl_days_left=None,
            screenshot_path=None,
            checked_at=self.run_dt.strftime("%Y-%m-%d %H:%M:%S"),
            session_label=self.session_label,
        )


class ExcelReporter:
    """Update one report workbook with date sheets and Morning/Afternoon sections."""

    SCREENSHOT_COL_HEADER = "Screenshot"
    SCREENSHOT_LINK_KEY = "_Screenshot Link"
    SCREENSHOT_COL_WIDTH = 48
    REPORT_COLUMNS = [
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
        SCREENSHOT_COL_HEADER,
    ]

    def create_report(
        self,
        sheet_results: dict,
        output_path: Path,
        with_screenshot: bool = True,
        run_dt: Optional[datetime] = None,
        session_label: str = "",
    ) -> None:
        """
        Update a single workbook.
        - One workbook is reused forever.
        - One worksheet is created per date.
        - Morning and Afternoon blocks are replaced on rerun instead of duplicated.
        """
        run_dt = run_dt or datetime.now()
        session_label = session_label or session_label_from_datetime(run_dt)
        sheet_name = run_dt.strftime("%Y-%m-%d")
        output_path.parent.mkdir(parents=True, exist_ok=True)

        wb = load_workbook(output_path) if output_path.exists() else Workbook()
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        elif self._is_single_default_blank_sheet(wb):
            ws = wb.active
            ws.title = sheet_name
        else:
            ws = wb.create_sheet(title=sheet_name)

        sections = self._read_existing_sections(ws)
        sections[session_label] = self._results_to_rows(sheet_results, with_screenshot)
        self._rewrite_date_sheet(ws, sheet_name, sections)
        saved_path = self._save_workbook(wb, output_path)
        logger.info(f"Report saved: {saved_path}")

    @staticmethod
    def _is_single_default_blank_sheet(wb: Workbook) -> bool:
        if wb.sheetnames != ["Sheet"]:
            return False
        ws = wb["Sheet"]
        return ws.max_row == 1 and ws.max_column == 1 and ws["A1"].value is None

    def _results_to_rows(self, sheet_results: dict, with_screenshot: bool) -> list:
        rows = []
        for results in sheet_results.values():
            for r in results:
                screenshot_link = r.screenshot_url if with_screenshot else ""
                screenshot_name = Path(str(r.screenshot_path or screenshot_link)).name if screenshot_link else ""
                rows.append({
                    "Group ID": r.group_id,
                    "Host": r.host,
                    "Type": r.type_,
                    "Target URL": r.target_url,
                    "Status": r.status,
                    "Status Code": r.status_code or "",
                    "Failure Reason": r.failure_reason,
                    "SSL Status": r.ssl_status,
                    "SSL Expiry": r.ssl_expiry,
                    "SSL Days Left": r.ssl_days_left if r.ssl_days_left is not None else "",
                    self.SCREENSHOT_COL_HEADER: screenshot_name,
                    self.SCREENSHOT_LINK_KEY: screenshot_link,
                })
        return rows

    def _read_existing_sections(self, ws) -> dict:
        sections = {}
        row = 1
        while row <= ws.max_row:
            value = ws.cell(row=row, column=1).value
            session = self._normalize_section_name(value)
            if session:
                header_row = row + 1
                headers = [
                    ws.cell(row=header_row, column=col).value
                    for col in range(1, ws.max_column + 1)
                ]
                data_rows = []
                current = header_row + 1
                while current <= ws.max_row:
                    first_cell = ws.cell(row=current, column=1).value
                    if self._normalize_section_name(first_cell):
                        break
                    if all(ws.cell(row=current, column=col).value in (None, "") for col in range(1, ws.max_column + 1)):
                        current += 1
                        continue
                    row_data = {}
                    for col in range(1, ws.max_column + 1):
                        header = headers[col - 1] if col - 1 < len(headers) else None
                        if header and header != "Session":
                            row_data[header] = ws.cell(row=current, column=col).value
                    screenshot_col = self._column_for_header(headers, self.SCREENSHOT_COL_HEADER)
                    screenshot_cell = ws.cell(row=current, column=screenshot_col) if screenshot_col else None
                    if screenshot_cell and screenshot_cell.hyperlink:
                        row_data[self.SCREENSHOT_LINK_KEY] = screenshot_cell.hyperlink.target
                    data_rows.append(row_data)
                    current += 1
                sections[session] = data_rows
                row = current
            else:
                row += 1
        return sections

    @staticmethod
    def _normalize_section_name(value) -> Optional[str]:
        text = str(value or "").strip().strip("-").strip()
        if text in {"Morning", "Afternoon"}:
            return text
        return None

    @staticmethod
    def _column_for_header(headers: list, header_name: str) -> Optional[int]:
        for index, header in enumerate(headers, start=1):
            if header == header_name:
                return index
        return None

    def _rewrite_date_sheet(self, ws, sheet_name: str, sections: dict) -> None:
        for merged_range in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(merged_range))
        ws._images = []
        if ws.max_row:
            ws.delete_rows(1, ws.max_row)

        for session in ("Morning", "Afternoon"):
            if session not in sections:
                continue
            if ws.max_row > 1:
                ws.append([])
            self._write_section(ws, session, sections[session])

        self._format_columns(ws)
        self._format_rows(ws)
        ws.freeze_panes = "A3"

    def _write_section(self, ws, session: str, rows: list) -> None:
        ws.append([f"--- {session} ---"])
        section_row = ws.max_row
        ws.merge_cells(
            start_row=section_row,
            start_column=1,
            end_row=section_row,
            end_column=len(self.REPORT_COLUMNS),
        )
        section_cell = ws.cell(row=section_row, column=1)
        section_cell.font = Font(bold=True, color="FFFFFF")
        section_cell.fill = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
        section_cell.alignment = Alignment(horizontal="left", vertical="center")

        header_row = ws.max_row + 1
        ws.append(self.REPORT_COLUMNS)
        for cell in ws[header_row]:
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row_data in rows:
            row_data = dict(row_data)
            row_data["Failure Reason"] = self._concise_failure_reason(row_data)
            excel_row = ws.max_row + 1
            ws.append([row_data.get(column, "") for column in self.REPORT_COLUMNS])
            status = str(row_data.get("Status", ""))
            is_success = status == "SUCCESS"
            for cell in ws[excel_row]:
                cell.font = Font(color="008000" if is_success else "000000")
                cell.alignment = Alignment(vertical="top", wrap_text=False)

            if not is_success:
                for column_name in ("Status", "Status Code", "Failure Reason"):
                    col_idx = self.REPORT_COLUMNS.index(column_name) + 1
                    ws.cell(row=excel_row, column=col_idx).font = Font(
                        color="FF0000",
                        bold=True,
                    )

            for column_name in ("Target URL",):
                col_idx = self.REPORT_COLUMNS.index(column_name) + 1
                cell = ws.cell(row=excel_row, column=col_idx)
                if isinstance(cell.value, str) and cell.value.startswith(("http://", "https://")):
                    cell.hyperlink = cell.value
                    cell.style = "Hyperlink"
                    cell.alignment = Alignment(vertical="top", wrap_text=False)

            screenshot_col_idx = self.REPORT_COLUMNS.index(self.SCREENSHOT_COL_HEADER) + 1
            screenshot_cell = ws.cell(row=excel_row, column=screenshot_col_idx)
            screenshot_link = row_data.get(self.SCREENSHOT_LINK_KEY, "")
            if screenshot_link:
                screenshot_cell.hyperlink = screenshot_link
                screenshot_cell.style = "Hyperlink"
                screenshot_cell.alignment = Alignment(vertical="top", wrap_text=False)

    @staticmethod
    def _concise_failure_reason(row_data: dict) -> str:
        original_reason = str(row_data.get("Failure Reason", "")).strip()
        if str(row_data.get("Status", "")).upper() != "FAILED":
            return ""

        status_code = row_data.get("Status Code", "")
        if status_code not in (None, ""):
            try:
                status_code = int(float(status_code))
            except (TypeError, ValueError):
                status_code = str(status_code).strip()
            if "browser" in original_reason.lower():
                return f"HTTP {status_code} - {original_reason}"
            return f"HTTP {status_code}"

        reason = original_reason.lower()
        concise_reasons = (
            ("timeout", "Timeout"),
            ("ssl", "SSL Error"),
            ("connection", "Connection Failed"),
            ("invalid url", "Invalid URL"),
            ("empty url", "Empty URL"),
            ("processing", "Processing Error"),
        )
        for marker, concise_reason in concise_reasons:
            if marker in reason:
                return concise_reason
        return "Request Error"

    def _format_columns(self, ws) -> None:
        widths = {
            "A": 10,
            "B": 28,
            "C": 14,
            "D": 44,
            "E": 12,
            "F": 12,
            "G": 12,
            "H": 12,
            "I": 14,
            "J": 14,
            "K": self.SCREENSHOT_COL_WIDTH,
        }
        for col_idx in range(1, len(self.REPORT_COLUMNS) + 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = widths.get(col_letter, 18)

    @staticmethod
    def _format_rows(ws) -> None:
        for row_idx in range(1, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 18

    @staticmethod
    def _save_workbook(wb: Workbook, output_path: Path) -> Path:
        try:
            wb.save(output_path)
            return output_path
        except PermissionError:
            fallback_path = output_path.with_name(
                f"{output_path.stem}_updated_{datetime.now().strftime('%Y%m%d_%H%M%S')}{output_path.suffix}"
            )
            wb.save(fallback_path)
            logger.warning(
                "Could not overwrite %s because it is open/locked. Saved a new copy at %s",
                output_path,
                fallback_path,
            )
            return fallback_path


class ConsoleReporter:
    """Console progress and summary"""
    
    @staticmethod
    def print_progress(result: URLResult, total: int) -> None:
        """Print progress for each URL"""
        idx = result.index + 1
        url = result.target_url[:60]
        
        status_icon = "[OK]" if result.status == "SUCCESS" else "[FAIL]"
        status_text = result.status
        
        # SSL warning if critical
        ssl_warning = ""
        if result.ssl_days_left is not None and result.ssl_days_left < 7:
            ssl_warning = f"!!! SSL: {result.ssl_days_left}d"
        
        print(f"[{idx:3d}/{total}] {status_icon} {status_text:<8} | {url:<62} {ssl_warning}", flush=True)
        
        if result.failure_reason:
            print(f"          |- {result.failure_reason}", flush=True)
    
    @staticmethod
    def print_summary(results: list, elapsed_time: float) -> None:
        """Print final summary"""
        total = len(results)
        success = sum(1 for r in results if r.status == "SUCCESS")
        failed = sum(1 for r in results if r.status == "FAILED")
        
        ssl_critical = sum(1 for r in results if r.ssl_days_left is not None and r.ssl_days_left < 7)
        ssl_warning = sum(1 for r in results if r.ssl_days_left is not None and 7 <= r.ssl_days_left < 30)
        
        response_times = [r.response_time for r in results if r.response_time]
        avg_response = round(sum(response_times) / len(response_times), 1) if response_times else 0
        
        print(f"\n{'='*90}")
        print(f"SUCCESS: {success}/{total} ({round(success/total*100, 1)}%)")
        print(f"FAILED:  {failed}/{total} ({round(failed/total*100, 1)}%)")
        print(f"Avg Response Time: {avg_response} ms")
        print(f"Completed in {elapsed_time:.2f}s ({round(total/elapsed_time, 1)} URLs/sec)")
        
        if ssl_critical > 0:
            print(f"\nCRITICAL: {ssl_critical} SSL certificate(s) expiring in <7 days")
        if ssl_warning > 0:
            print(f"WARNING: {ssl_warning} SSL certificate(s) expiring in <30 days")
        
        print(f"{'='*90}\n")
        
        # Print detailed SSL report after the separator
        if ssl_critical > 0:
            ConsoleReporter._print_ssl_details(results)
    
    @staticmethod
    def _print_ssl_details(results: list) -> None:
        """Print detailed SSL expiration report grouped by days remaining"""
        # Collect URLs with SSL expiring in <7 days
        ssl_critical_results = [r for r in results if r.ssl_days_left is not None and r.ssl_days_left < 7]
        
        if not ssl_critical_results:
            return
        
        sorted_results = sorted(ssl_critical_results, key=lambda r: (r.ssl_days_left, r.target_url))

        print("The SSL certificates expiring in <7 days are:\n", flush=True)

        for result in sorted_results:
            print(f"{result.target_url} - {result.ssl_days_left} days", flush=True)


def session_label_from_datetime(run_dt: datetime) -> str:
    override = os.getenv("MONITOR_RUN_SESSION", "").strip().lower()
    if override in {"morning", "am"}:
        return "Morning"
    if override in {"afternoon", "pm", "evening"}:
        return "Afternoon"
    return "Morning" if run_dt.hour < 12 else "Afternoon"


class URLMonitor:
    """Main orchestrator"""
    
    def __init__(self, config: Config = None):
        self.config = config or Config()
        self.run_dt = datetime.now()
        self.session_label = session_label_from_datetime(self.run_dt)
        screenshot_dir = Path("screenshots") / self.run_dt.strftime("%Y-%m-%d") / self.session_label
        self.screenshot_capture = ScreenshotCapture(output_dir=str(screenshot_dir)) if self.config.with_screenshot else None
        self.auth_screenshot_capture = (
            ScreenshotCapture(
                output_dir=str(screenshot_dir),
                pool_size=1,
                capture_browser_ui=True,
            )
            if self.config.with_screenshot and self.config.auth_browser_ui_screenshot
            else None
        )
        self.processor = URLProcessor(
            self.config,
            self.screenshot_capture,
            self.auth_screenshot_capture,
            self.run_dt,
            self.session_label,
        )
        self.excel_reporter = ExcelReporter()
        self.console_reporter = ConsoleReporter()
        self.sharepoint_client = SharePointGraphClient.from_env()
        if self.sharepoint_client:
            logger.info("SharePoint sync enabled: %s", self.sharepoint_client.workbook_url)
        else:
            logger.warning(
                "SharePoint sync is OFF. Local report will update at %s. "
                "Set SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, and SHAREPOINT_CLIENT_SECRET "
                "to update the SharePoint workbook.",
                self.config.report_file,
            )
    
    def monitor_urls(self, input_file: str, output_file: str = None) -> None:
        """Monitor URLs from all sheets in the Excel file"""
        try:
            input_path = Path(input_file)
            if not input_path.exists():
                raise FileNotFoundError(f"Input file not found: {input_file}")
            
            if output_file is None:
                output_file = self.config.report_file

            output_path = Path(output_file)
            if self.sharepoint_client:
                self.sharepoint_client.download_workbook(output_path)
            
            # Load the Excel file to get all sheets
            logger.info(f"Reading Excel file: {input_path}")
            xl = pd.ExcelFile(input_path)
            sheet_names = xl.sheet_names
            
            all_sheet_results = {}
            total_start_time = time.time()
            all_results_flat = []
            
            print(f"\n{'='*90}")
            print(f"STARTING URL MONITOR - Found {len(sheet_names)} sheets")
            print(f"{'='*90}\n")
            
            for sheet_name in sheet_names:
                logger.info(f"Processing sheet: {sheet_name}")
                df = pd.read_excel(xl, sheet_name=sheet_name)
                df.columns = df.columns.str.strip()
                
                if df.empty:
                    logger.warning(f"Sheet '{sheet_name}' is empty")
                    continue
                
                print(f"\n--- Sheet: {sheet_name} ({len(df)} URLs) ---")
                
                results = self._process_parallel(df)
                results.sort(key=lambda x: x.index)
                
                all_sheet_results[sheet_name] = results
                all_results_flat.extend(results)
            
            total_elapsed_time = time.time() - total_start_time
            
            if not all_results_flat:
                logger.warning("No URLs found in any sheet")
                return
            
            # Generate reports
            self._publish_screenshot_links(all_results_flat)
            self._validate_screenshot_results(all_results_flat)
            self.excel_reporter.create_report(
                all_sheet_results,
                output_path,
                with_screenshot=self.config.with_screenshot,
                run_dt=self.run_dt,
                session_label=self.session_label,
            )
            if self.sharepoint_client and not self.sharepoint_client.upload_workbook(output_path):
                raise RuntimeError("SharePoint workbook upload failed")
            self.console_reporter.print_summary(all_results_flat, total_elapsed_time)
            
            # Save SSL summary for notifications
            self._save_ssl_summary(all_results_flat)
            
            logger.info(f"Monitoring completed. Report: {output_file}")
            
        except Exception as e:
            logger.error(f"Error: {e}")
            raise
        finally:
            # Clean up ALL WebDriver instances across all threads
            if self.screenshot_capture:
                self.screenshot_capture.quit_all_drivers()
            if self.auth_screenshot_capture:
                self.auth_screenshot_capture.quit_all_drivers()

    def _save_ssl_summary(self, results: list) -> None:
        """Save detailed SSL expiration report to a text file for notifications"""
        try:
            ssl_critical_results = [r for r in results if r.ssl_days_left is not None and r.ssl_days_left < 7]
            summary_file = Path("ssl_summary.txt")
            if ssl_critical_results:
                sorted_results = sorted(ssl_critical_results, key=lambda r: (r.ssl_days_left, r.target_url))
                lines = ["The SSL certificates expiring in <7 days are:\n"]
                for result in sorted_results:
                    lines.append(f"{result.target_url} - {result.ssl_days_left} days")
                summary_text = "\n".join(lines)
                summary_file.write_text(summary_text, encoding="utf-8")
                logger.info(f"SSL summary saved to {summary_file}")
            else:
                if summary_file.exists():
                    summary_file.unlink()
        except Exception as e:
            logger.error(f"Failed to save SSL summary: {e}")

    def _publish_screenshot_links(self, results: list) -> None:
        if not self.config.with_screenshot:
            return

        tmp_uploader = None
        if (
            self.sharepoint_client is None
            and self.config.screenshot_upload_provider.strip().lower() == "tmpfiles"
        ):
            tmp_uploader = TemporaryScreenshotUploader(self.config.screenshot_link_expire_seconds)

        for result in results:
            if not result.screenshot_path:
                continue

            screenshot_path = Path(result.screenshot_path)
            if not screenshot_path.exists():
                continue

            if self.sharepoint_client and self.config.sharepoint_upload_screenshots:
                remote_path = (
                    Path("screenshots")
                    / self.run_dt.strftime("%Y-%m-%d")
                    / self.session_label
                    / screenshot_path.name
                )
                result.screenshot_url = self.sharepoint_client.upload_related_file(
                    screenshot_path,
                    remote_path.as_posix(),
                )

            if not result.screenshot_url and tmp_uploader:
                result.screenshot_url = tmp_uploader.upload(screenshot_path)

            # Do not fall back to raw GitHub URLs when tmpfiles was requested:
            # private repositories return 404 for Excel/Teams report viewers.
            if not result.screenshot_url and tmp_uploader is None:
                result.screenshot_url = self._build_online_url(screenshot_path)

    def _validate_screenshot_results(self, results: list) -> None:
        """Report screenshot gaps, optionally failing when strict mode is enabled."""
        if not self.config.with_screenshot:
            return

        # URLs that failed to load (timeout, connection refused, etc.) cannot
        # produce a screenshot — that is expected, not an error.
        failed_urls = [
            result.target_url or str(result.host)
            for result in results
            if result.status == "FAILED"
        ]
        if failed_urls:
            logger.warning(
                "Skipping screenshot validation for %d FAILED URL(s): %s",
                len(failed_urls),
                ", ".join(failed_urls[:5])
                + (f", and {len(failed_urls) - 5} more" if len(failed_urls) > 5 else ""),
            )

        # Only enforce screenshot coverage on URLs that loaded successfully.
        missing_results = [
            result
            for result in results
            if result.status != "FAILED"
            and (
                not result.screenshot_path
                or not Path(result.screenshot_path).is_file()
                or not result.screenshot_url
            )
        ]
        missing_urls = [
            result.target_url or str(result.host)
            for result in missing_results
        ]
        if missing_urls:
            preview = ", ".join(missing_urls[:5])
            if len(missing_urls) > 5:
                preview += f", and {len(missing_urls) - 5} more"
            message = f"Screenshot coverage incomplete for: {preview}"
            if self.config.strict_screenshot_coverage:
                raise RuntimeError(message)
            logger.warning(
                "%s. The endpoint results and Excel report will still be published. "
                "Set MONITOR_STRICT_SCREENSHOT_COVERAGE=1 to fail on this condition.",
                message,
            )

        successful_count = sum(1 for r in results if r.status != "FAILED")
        captured_count = successful_count - len(missing_results)
        logger.info("Screenshot coverage: %d/%d (excluding %d failed URLs)", captured_count, successful_count, len(failed_urls))

    def _build_online_url(self, path: Path) -> str:
        base_url = self.config.screenshot_online_base_url.strip()
        if not base_url:
            repo = os.getenv("GITHUB_REPOSITORY", "").strip()
            branch = os.getenv("GITHUB_REF_NAME", "main").strip() or "main"
            if repo:
                base_url = f"https://raw.githubusercontent.com/{repo}/{branch}/"

        if not base_url:
            return str(path)

        try:
            relative_path = path.resolve().relative_to(Path.cwd().resolve()).as_posix()
        except ValueError:
            relative_path = path.as_posix()
        return f"{base_url.rstrip('/')}/{quote(relative_path)}"
    
    def _process_parallel(self, df: pd.DataFrame) -> list:
        """Process URLs in parallel without ever dropping an input row."""
        results = []
        total = len(df)
        
        with ThreadPoolExecutor(max_workers=self.config.max_workers) as executor:
            future_to_row = {
                executor.submit(self.processor.process_url, idx, row): (idx, row)
                for idx, row in df.iterrows()
            }
            
            for future in as_completed(future_to_row):
                idx, row = future_to_row[future]
                try:
                    result = future.result()
                    results.append(result)
                    self.console_reporter.print_progress(result, total)
                except Exception as e:
                    logger.error("Worker failed at index %s: %s", idx, e)
                    fallback = self.processor._create_failed_result(
                        idx,
                        row,
                        str(row.get("Target URL", "")).strip(),
                        "Processing Error",
                    )
                    results.append(fallback)
                    self.console_reporter.print_progress(fallback, total)
        return results


def main():
    """Entry point"""
    config = Config(
        max_workers=int(os.getenv("MONITOR_MAX_WORKERS", "10")),
        timeout=int(os.getenv("MONITOR_TIMEOUT", "20")),
        max_retries=int(os.getenv("MONITOR_MAX_RETRIES", "4")),
        retry_delay=float(os.getenv("MONITOR_RETRY_DELAY", "2")),
        with_screenshot=WITH_SCREENSHOT,
        input_file=os.getenv("MONITOR_INPUT_FILE", "Monitor Endpoints.xlsx"),
        report_file=os.getenv("MONITOR_REPORT_FILE", "reports/Website_Monitor_Report.xlsx"),
        screenshot_online_base_url=os.getenv("SCREENSHOT_ONLINE_BASE_URL", ""),
        screenshot_upload_provider=os.getenv("SCREENSHOT_UPLOAD_PROVIDER", "tmpfiles"),
        screenshot_link_expire_seconds=int(os.getenv("SCREENSHOT_LINK_EXPIRE_SECONDS", "172800")),
        sharepoint_upload_screenshots=os.getenv("SHAREPOINT_UPLOAD_SCREENSHOTS", "1").strip().lower() not in {"0", "false", "no"},
        auth_browser_ui_screenshot=os.getenv("MONITOR_AUTH_BROWSER_UI_SCREENSHOT", "1").strip().lower() not in {"0", "false", "no"},
        strict_screenshot_coverage=os.getenv("MONITOR_STRICT_SCREENSHOT_COVERAGE", "0").strip().lower() in {"1", "true", "yes"},
    )
    
    monitor = URLMonitor(config)
    monitor.monitor_urls(config.input_file)


if __name__ == "__main__":
    main()
