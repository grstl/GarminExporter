#!/usr/bin/env python3
"""Export Garmin Connect activities to an Excel file."""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime
from getpass import getpass
from pathlib import Path
from typing import Any

from garth.exc import GarthException, GarthHTTPError
from garminconnect import (
    Garmin,
    GarminConnectAuthenticationError,
    GarminConnectConnectionError,
    GarminConnectTooManyRequestsError,
)
from openpyxl import Workbook

LOGGER = logging.getLogger("garmin_exporter")

EXCEL_COLUMNS = [
    "activity_id",
    "activity_name",
    "activity_type",
    "event_type",
    "start_time_local",
    "start_time_gmt",
    "distance_m",
    "duration_s",
    "moving_duration_s",
    "elapsed_duration_s",
    "average_speed_mps",
    "max_speed_mps",
    "calories",
    "average_hr_bpm",
    "max_hr_bpm",
    "elevation_gain_m",
    "elevation_loss_m",
    "steps",
    "location_name",
    "privacy",
]

REDACTION_PATTERNS = (
    re.compile(r"(?i)(password\s*[=:]\s*)(\S+)"),
    re.compile(r"(?i)(mfa[_\s-]*code\s*[=:]\s*)(\S+)"),
)


@dataclass
class Credentials:
    email: str
    password: str


class SecretRedactionFilter(logging.Filter):
    """Redacts credential-like values from log output."""

    def __init__(self, secrets: list[str]) -> None:
        super().__init__()
        self._secrets = [secret for secret in secrets if secret]

    def filter(self, record: logging.LogRecord) -> bool:
        message = record.getMessage()
        redacted = message

        for secret in self._secrets:
            redacted = redacted.replace(secret, "***REDACTED***")

        for pattern in REDACTION_PATTERNS:
            redacted = pattern.sub(r"\1***REDACTED***", redacted)

        if redacted != message:
            record.msg = redacted
            record.args = ()

        return True


def setup_logging(level_name: str, secrets: list[str]) -> None:
    level = getattr(logging, level_name.upper(), logging.INFO)
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    handler.addFilter(SecretRedactionFilter(secrets))

    root_logger = logging.getLogger()
    root_logger.handlers.clear()
    root_logger.setLevel(level)
    root_logger.addHandler(handler)

    # Avoid noisy internals from dependency libraries.
    logging.getLogger("garminconnect").setLevel(logging.WARNING)
    logging.getLogger("urllib3").setLevel(logging.WARNING)


def parse_cli_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export Garmin Connect activities to an Excel workbook."
    )
    parser.add_argument(
        "-o",
        "--output",
        default="garmin_activities.xlsx",
        help="Output Excel file path (default: garmin_activities.xlsx).",
    )
    parser.add_argument(
        "--from-date",
        type=parse_date_arg,
        help="Optional earliest activity date (YYYY-MM-DD).",
    )
    parser.add_argument(
        "--to-date",
        type=parse_date_arg,
        help="Optional latest activity date (YYYY-MM-DD).",
    )
    parser.add_argument(
        "--max-activities",
        type=int,
        default=None,
        help="Optional max number of newest activities to fetch.",
    )
    parser.add_argument(
        "--tokens-dir",
        default="~/.garminconnect",
        help="Token directory for Garmin session reuse (default: ~/.garminconnect).",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        help="Logging level (DEBUG, INFO, WARNING, ERROR).",
    )
    return parser.parse_args()


def parse_date_arg(raw_date: str) -> date:
    try:
        return date.fromisoformat(raw_date)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"Invalid date '{raw_date}'. Use YYYY-MM-DD format."
        ) from exc


def get_nested(source: dict[str, Any], *keys: str, default: Any = None) -> Any:
    current: Any = source
    for key in keys:
        if not isinstance(current, dict):
            return default
        current = current.get(key)
        if current is None:
            return default
    return current


def parse_activity_date(activity: dict[str, Any]) -> date | None:
    raw = activity.get("startTimeLocal") or activity.get("startTimeGMT")
    if not raw:
        return None

    # Handles values such as 2024-10-20 06:02:00 or ISO strings with timezone.
    normalized = str(raw).replace("Z", "+00:00").replace(" ", "T")
    try:
        return datetime.fromisoformat(normalized).date()
    except ValueError:
        return None


def activity_is_in_range(
    activity: dict[str, Any], from_date: date | None, to_date: date | None
) -> bool:
    if from_date is None and to_date is None:
        return True

    activity_date = parse_activity_date(activity)
    if activity_date is None:
        return False
    if from_date and activity_date < from_date:
        return False
    if to_date and activity_date > to_date:
        return False
    return True


def get_credentials() -> Credentials:
    email = os.getenv("GARMIN_EMAIL")
    password = os.getenv("GARMIN_PASSWORD")

    if not email:
        email = input("Garmin email: ").strip()
    if not password:
        password = getpass("Garmin password: ").strip()

    if not email or not password:
        raise ValueError("Garmin email and password are required.")

    return Credentials(email=email, password=password)


def authenticate(tokens_dir: Path, credentials: Credentials) -> Garmin:
    tokens_dir = tokens_dir.expanduser()

    # Prefer token login for security and better UX.
    try:
        api = Garmin()
        api.login(str(tokens_dir))
        LOGGER.info("Authenticated with existing Garmin tokens.")
        return api
    except (
        FileNotFoundError,
        GarminConnectAuthenticationError,
        GarminConnectConnectionError,
        GarthHTTPError,
    ):
        LOGGER.info("No valid token session found. Using credential login.")

    api = Garmin(
        email=credentials.email,
        password=credentials.password,
        is_cn=False,
        return_on_mfa=True,
    )
    login_result = api.login()
    if isinstance(login_result, tuple) and login_result and login_result[0] == "needs_mfa":
        mfa_code = os.getenv("GARMIN_MFA_CODE") or getpass("MFA code: ").strip()
        if not mfa_code:
            raise GarminConnectAuthenticationError("MFA code is required.")
        api.resume_login(login_result[1], mfa_code)

    tokens_dir.mkdir(parents=True, exist_ok=True)
    api.garth.dump(str(tokens_dir))
    LOGGER.info("Authenticated with credentials and refreshed Garmin tokens.")
    return api


def fetch_activities(api: Garmin, max_activities: int | None = None) -> list[dict[str, Any]]:
    page_size = 100
    offset = 0
    activities: list[dict[str, Any]] = []

    while True:
        requested = page_size
        if max_activities is not None:
            remaining = max_activities - len(activities)
            if remaining <= 0:
                break
            requested = min(requested, remaining)

        batch = api.get_activities(offset, requested)
        if not batch:
            break

        activities.extend(batch)
        offset += len(batch)
        LOGGER.info("Fetched %s total activities.", len(activities))

        if len(batch) < requested:
            break

    return activities


def activity_to_row(activity: dict[str, Any]) -> dict[str, Any]:
    return {
        "activity_id": activity.get("activityId"),
        "activity_name": activity.get("activityName"),
        "activity_type": get_nested(activity, "activityType", "typeKey"),
        "event_type": activity.get("eventType"),
        "start_time_local": activity.get("startTimeLocal"),
        "start_time_gmt": activity.get("startTimeGMT"),
        "distance_m": activity.get("distance"),
        "duration_s": activity.get("duration"),
        "moving_duration_s": activity.get("movingDuration"),
        "elapsed_duration_s": activity.get("elapsedDuration"),
        "average_speed_mps": activity.get("averageSpeed"),
        "max_speed_mps": activity.get("maxSpeed"),
        "calories": activity.get("calories"),
        "average_hr_bpm": activity.get("averageHR"),
        "max_hr_bpm": activity.get("maxHR"),
        "elevation_gain_m": activity.get("elevationGain"),
        "elevation_loss_m": activity.get("elevationLoss"),
        "steps": activity.get("steps"),
        "location_name": activity.get("locationName"),
        "privacy": activity.get("privacy"),
    }


def write_excel(rows: list[dict[str, Any]], output_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Activities"

    worksheet.append(EXCEL_COLUMNS)
    for row in rows:
        worksheet.append([row.get(column) for column in EXCEL_COLUMNS])

    output_path = output_path.expanduser()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def run() -> int:
    args = parse_cli_args()
    credentials = get_credentials()
    setup_logging(
        args.log_level,
        [
            credentials.email,
            credentials.password,
            os.getenv("GARMIN_MFA_CODE", ""),
        ],
    )

    if args.max_activities is not None and args.max_activities <= 0:
        LOGGER.error("--max-activities must be a positive integer.")
        return 2
    if args.from_date and args.to_date and args.from_date > args.to_date:
        LOGGER.error("--from-date cannot be later than --to-date.")
        return 2

    try:
        api = authenticate(Path(args.tokens_dir), credentials)
        activities = fetch_activities(api, max_activities=args.max_activities)
        filtered = [
            activity_to_row(activity)
            for activity in activities
            if activity_is_in_range(activity, args.from_date, args.to_date)
        ]
        write_excel(filtered, Path(args.output))
        LOGGER.info("Export complete: %s activities written to %s", len(filtered), args.output)
        return 0
    except (
        GarminConnectAuthenticationError,
        GarminConnectConnectionError,
        GarminConnectTooManyRequestsError,
        GarthHTTPError,
        GarthException,
        OSError,
        ValueError,
    ) as exc:
        LOGGER.error("Export failed: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(run())
