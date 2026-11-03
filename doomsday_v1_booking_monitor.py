
import requests
from bs4 import BeautifulSoup
from datetime import datetime
from io import BytesIO
import boto3
from openpyxl import Workbook

# BookMyShow page to monitor
URL = (
    "https://in.bookmyshow.com/movies/hyderabad/fauzi/ET00467913"
)

# AWS configuration
BUCKET_NAME = "project-doomsday"
AWS_REGION = "us-east-1"

# S3 destinations: use your existing folder names
S3_FOLDERS = {
    "BOOKINGS_NOT_OPEN": "bookings-not-open",
    "BOOKING_TEXT_FOUND_NEEDS_VERIFICATION": "bookings-open",
    "UNKNOWN": "error",
}

# Same filename each time, so the previous object is overwritten
FILE_NAME = "status.xlsx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/130.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-IN,en;q=0.9",
}

s3 = boto3.client("s3", region_name=AWS_REGION)


def check_booking_status():
    checked_at = datetime.now().astimezone().isoformat(
        timespec="seconds"
    )
    http_status = None
    details = ""

    try:
        response = requests.get(
            URL,
            headers=HEADERS,
            timeout=30,
        )

        http_status = response.status_code
        print("HTTP status:", http_status)

        if http_status != 200:
            return (
                "UNKNOWN",
                checked_at,
                http_status,
                f"HTTP {http_status}",
            )

        soup = BeautifulSoup(response.text, "html.parser")
        page_text = soup.get_text(" ", strip=True).lower()

        waiting_hint = (
            "mark interested to know when bookings open"
        )
        booking_hint = "book tickets"

        if waiting_hint in page_text:
            status = "BOOKINGS_NOT_OPEN"
            details = "Waiting message found"

        elif booking_hint in page_text:
            status = "BOOKING_TEXT_FOUND_NEEDS_VERIFICATION"
            details = "Booking text found; verify manually"

        else:
            status = "UNKNOWN"
            details = "Expected booking text not found"

    except requests.RequestException as exc:
        status = "UNKNOWN"
        details = f"Request error: {type(exc).__name__}"

    return status, checked_at, http_status, details


def create_excel(status, checked_at, http_status, details):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Booking Status"

    sheet.append(["Field", "Value"])
    sheet.append(["Movie URL", URL])
    sheet.append(["Checked At", checked_at])
    sheet.append(["HTTP Status", http_status])
    sheet.append(["Booking Status", status])
    sheet.append(["Details", details])

    # Make the columns readable in Excel
    sheet.column_dimensions["A"].width = 22
    sheet.column_dimensions["B"].width = 75

    excel_buffer = BytesIO()
    workbook.save(excel_buffer)
    excel_buffer.seek(0)

    return excel_buffer


def upload_result(status, checked_at, http_status, details):
    folder = S3_FOLDERS[status]
    s3_key = f"{folder}/{FILE_NAME}"

    excel_file = create_excel(
        status,
        checked_at,
        http_status,
        details,
    )

    s3.put_object(
        Bucket=BUCKET_NAME,
        Key=s3_key,
        Body=excel_file.getvalue(),
        ContentType=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    print("STATUS:", status)
    print("S3 location:", f"s3://{BUCKET_NAME}/{s3_key}")
    print("Upload successful.")


if __name__ == "__main__":
    status, checked_at, http_status, details = (
        check_booking_status()
    )

    try:
        upload_result(
            status,
            checked_at,
            http_status,
            details,
        )
    except Exception as exc:
        print("S3 upload failed:", str(exc))
        raise
