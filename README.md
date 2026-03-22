# Garmin Activity Excel Exporter

CLI application that logs in to Garmin Connect securely and exports activities to
an Excel workbook.

## Features

- Exports Garmin activities into `.xlsx` format.
- Exactly one Excel row per activity.
- Supports date-range filtering.
- Supports activity limit pagination.
- Uses token reuse when available to avoid frequent credential prompts.
- Avoids leaking credentials in logs (password/MFA values are redacted).

## Requirements

- Python 3.10+
- A Garmin Connect account

Install dependencies:

```bash
python -m pip install -r requirements.txt
```

## Secure authentication model

The app **does not** accept password in CLI arguments (to avoid shell history leaks).
Use environment variables or secure prompts:

- `GARMIN_EMAIL` (optional, otherwise prompted)
- `GARMIN_PASSWORD` (optional, otherwise prompted with hidden input)
- `GARMIN_MFA_CODE` (optional, otherwise prompted if MFA is required)

Session tokens are reused from `~/.garminconnect` by default and refreshed after
successful credential login.

## Usage

```bash
python garmin_exporter.py --output garmin_activities.xlsx
```

Useful options:

```bash
python garmin_exporter.py \
  --output exports/activities.xlsx \
  --from-date 2025-01-01 \
  --to-date 2025-12-31 \
  --max-activities 500 \
  --tokens-dir ~/.garminconnect \
  --log-level INFO
```

## Output columns

- activity_id
- activity_name
- activity_type
- event_type
- start_time_local
- start_time_gmt
- distance_m
- duration_s
- moving_duration_s
- elapsed_duration_s
- average_speed_mps
- max_speed_mps
- calories
- average_hr_bpm
- max_hr_bpm
- elevation_gain_m
- elevation_loss_m
- steps
- location_name
- privacy