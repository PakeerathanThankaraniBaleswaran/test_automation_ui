from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright
from pathlib import Path
import argparse
import base64
import csv
import sys
import time
from datetime import datetime
from openpyxl import Workbook


DEFAULT_URL = "https://www.pixelssuite.com/convert-to-png"
DEFAULT_TIMEOUT_MS = 60000
DEFAULT_SLOW_MO_MS = 0
DEFAULT_CSV = "execution_results.csv"
DEFAULT_XLSX = "execution_results.xlsx"
DEFAULT_OUT_DIR = "results"
DEFAULT_INPUT_FILE = "sample.png"

PNG_1X1_BASE64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMB/6X9wYQAAAAASUVORK5CYII="
)

CSV_FIELDS = [
    "timestamp",
    "url",
    "file_type",
    "file_path",
    "file_name",
    "preview_detected",
    "status",
    "message",
    "before_upload_screenshot",
    "after_upload_screenshot",
]


def configure_stdout():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    except Exception:
        pass


def parse_args():
    parser = argparse.ArgumentParser(
        description="Upload a PNG file to PixelsSuite and verify that a preview is displayed."
    )
    parser.add_argument("--url", default=DEFAULT_URL)
    parser.add_argument("--png", default=DEFAULT_INPUT_FILE)
    parser.add_argument("--out-dir", default=DEFAULT_OUT_DIR)
    parser.add_argument("--csv", default=DEFAULT_CSV)
    parser.add_argument("--xlsx", default=DEFAULT_XLSX)
    parser.add_argument("--headless", action="store_true", default=False)
    parser.add_argument("--timeout-ms", type=int, default=DEFAULT_TIMEOUT_MS)
    parser.add_argument("--slow-mo-ms", type=int, default=DEFAULT_SLOW_MO_MS)
    parser.add_argument(
        "--append-csv",
        action="store_true",
        default=False,
        help="Append to the CSV file instead of overwriting it for a fresh run.",
    )
    return parser.parse_args()


def ensure_default_png(file_path: Path):
    file_path.parent.mkdir(parents=True, exist_ok=True)
    if not file_path.exists():
        file_path.write_bytes(base64.b64decode(PNG_1X1_BASE64))


def reset_csv(csv_path: Path):
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with csv_path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_FIELDS)
        writer.writeheader()


def append_result(csv_path: Path, result: dict):
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    file_exists = csv_path.exists()
    with csv_path.open("a", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_FIELDS)
        if not file_exists:
            writer.writeheader()
        writer.writerow({field: result.get(field, "") for field in CSV_FIELDS})


def write_result_to_excel(xlsx_path: Path, result: dict):
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Execution Results"
    sheet.append(CSV_FIELDS)
    sheet.append([result.get(field, "") for field in CSV_FIELDS])
    workbook.save(xlsx_path)


def safe_wait_for_page_idle(page):
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except PlaywrightTimeoutError:
        pass


def dismiss_common_popups(page):
    popup_labels = [
        "Accept",
        "Accept All",
        "I Agree",
        "Got it",
        "Close",
    ]
    for label in popup_labels:
        try:
            button = page.get_by_role("button", name=label, exact=False).first
            if button.count() > 0 and button.is_visible(timeout=500):
                button.click(timeout=1000)
                page.wait_for_timeout(300)
        except Exception:
            continue


def find_upload_input(page, timeout_ms: int):
    selectors = [
        'input[type="file"]',
        'input[accept*="image"]',
        'input[accept*="png"]',
    ]

    deadline = time.time() + (timeout_ms / 1000)
    while time.time() < deadline:
        dismiss_common_popups(page)
        for selector in selectors:
            try:
                locator = page.locator(selector).first
                if locator.count() > 0:
                    return locator
            except Exception:
                continue
        page.wait_for_timeout(300)

    raise RuntimeError("File upload input was not found on the page.")


def preview_state(page, expected_file_name: str):
    script = """
    ({ expectedFileName }) => {
        const isVisible = (element) => !!(
            element &&
            element.isConnected &&
            element.getClientRects &&
            element.getClientRects().length
        );

        const lowerExpectedName = (expectedFileName || "").toLowerCase();
        const labels = Array.from(document.querySelectorAll("body *"))
            .filter((element) => element.childElementCount === 0)
            .filter((element) => isVisible(element))
            .map((element) => (element.textContent || "").trim())
            .filter(Boolean);

        const previewLabel = labels.find((text) => text.toLowerCase().includes("preview"));
        const filenameVisible = labels.some((text) => text.toLowerCase().includes(lowerExpectedName));
        const visibleMediaCount = Array.from(document.querySelectorAll("img, canvas, svg, video"))
            .filter((element) => isVisible(element))
            .length;
        const fileInputHasSelection = Array.from(document.querySelectorAll('input[type="file"]'))
            .some((input) => input.files && input.files.length > 0);

        return {
            previewLabelFound: Boolean(previewLabel),
            filenameVisible,
            visibleMediaCount,
            fileInputHasSelection,
        };
    }
    """
    return page.evaluate(script, {"expectedFileName": expected_file_name})


def wait_for_preview(page, expected_file_name: str, timeout_ms: int):
    deadline = time.time() + (timeout_ms / 1000)
    last_state = {}

    while time.time() < deadline:
        last_state = preview_state(page, expected_file_name)
        if (
            last_state.get("previewLabelFound")
            and last_state.get("visibleMediaCount", 0) > 0
        ) or (
            last_state.get("filenameVisible")
            and last_state.get("fileInputHasSelection")
        ):
            return True, last_state

        page.wait_for_timeout(500)

    return False, last_state


def capture_screenshot(page, output_path: Path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(output_path), full_page=True)
    return str(output_path)


def print_summary(result: dict):
    print("========== ASSIGNMENT EXECUTION RESULT ==========")
    print(f"Timestamp             : {result['timestamp']}")
    print(f"URL                   : {result['url']}")
    print(f"File type             : {result['file_type']}")
    print(f"File path             : {result['file_path']}")
    print(f"Preview detected      : {result['preview_detected']}")
    print(f"Status                : {result['status']}")
    print(f"Message               : {result['message']}")
    print(f"Before upload shot    : {result['before_upload_screenshot']}")
    print(f"After upload shot     : {result['after_upload_screenshot']}")


def build_result(args, png_path: Path, status: str, message: str):
    return {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "url": args.url,
        "file_type": png_path.suffix.replace(".", "").upper() or "UNKNOWN",
        "file_path": str(png_path),
        "file_name": png_path.name,
        "preview_detected": False,
        "status": status,
        "message": message,
        "before_upload_screenshot": "",
        "after_upload_screenshot": "",
    }


def run_assignment():
    configure_stdout()
    args = parse_args()

    png_path = Path(args.png).resolve()
    output_dir = Path(args.out_dir).resolve()
    csv_path = Path(args.csv).resolve()
    xlsx_path = Path(args.xlsx).resolve()

    output_dir.mkdir(parents=True, exist_ok=True)
    ensure_default_png(png_path)
    if not args.append_csv:
        reset_csv(csv_path)

    result = build_result(args, png_path, "FAIL", "Execution did not start.")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=args.headless,
            slow_mo=args.slow_mo_ms,
        )
        page = browser.new_page()
        page.set_default_timeout(args.timeout_ms)

        try:
            page.goto(args.url, wait_until="domcontentloaded")
            safe_wait_for_page_idle(page)
            dismiss_common_popups(page)

            before_upload = output_dir / "01_before_upload.png"
            result["before_upload_screenshot"] = capture_screenshot(page, before_upload)

            file_input = find_upload_input(page, args.timeout_ms)
            file_input.set_input_files(str(png_path))
            safe_wait_for_page_idle(page)

            preview_detected, state = wait_for_preview(page, png_path.name, args.timeout_ms)

            after_upload_name = "02_after_upload_pass.png" if preview_detected else "02_after_upload_fail.png"
            after_upload = output_dir / after_upload_name
            result["after_upload_screenshot"] = capture_screenshot(page, after_upload)

            result["preview_detected"] = preview_detected
            result["status"] = "PASS" if preview_detected else "FAIL"
            result["message"] = (
                "PNG uploaded successfully and preview was detected."
                if preview_detected
                else f"Upload completed, but preview evidence was not detected. Last observed state: {state}"
            )
        except Exception as error:
            error_shot = output_dir / "02_after_upload_error.png"
            try:
                result["after_upload_screenshot"] = capture_screenshot(page, error_shot)
            except Exception:
                pass

            result["status"] = "FAIL"
            result["preview_detected"] = False
            result["message"] = f"Execution failed: {error}"
        finally:
            browser.close()

    append_result(csv_path, result)
    write_result_to_excel(xlsx_path, result)
    print_summary(result)
    print(f"CSV                   : {csv_path}")
    print(f"Excel                 : {xlsx_path}")
    return 0 if result["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(run_assignment())
