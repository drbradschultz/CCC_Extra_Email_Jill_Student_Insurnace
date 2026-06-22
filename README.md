# CCC Student Roster Email

Google Apps Script that emails Dr. Schultz a weekly summary of students **starting at** or **leaving** Collective Care Clinic.

Source: [src/Student_Departures_Email.gs](src/Student_Departures_Email.gs)

## What it does

- Runs Monday morning (7 AM) via a time-based trigger.
- Reads the roster tab (`Staff and Students 25-26`): student names in column **C**, dates in column **D**.
- Reports the **upcoming week only** — anyone starting or leaving in the next 7 days (today → +7). Past dates are not listed.
- Sends a styled HTML email (with plain-text fallback) split into **Starting This Week** and **Leaving This Week**. Nothing is sent if no one falls in the window.

## Marking dates in column D

- **End date (departure):** just the date — e.g. `6/27/26`. This is the default.
- **Start date (arrival):** add the word `START` to the cell — e.g. `START 6/25/26` or `6/25/26 (START)`. Any cell containing "START" is treated as an arrival.
- **Unknown date:** a placeholder (`?`, `X`, or `TBD`) omits the student until a real date is entered.

Dates accept `m/d`, `m/d/yy`, `mm/dd/yyyy`, etc. A missing year is inferred to the nearest occurrence.

## Functions

- `sd_runWeekly()` — entry point; builds and sends the email.
- `sd_setupTrigger()` — run once from the editor to schedule the Monday send.
- `sd_previewSample()` — emails a sample with fabricated entries to preview the layout (does not read the roster).

## Config

Edit `SD_CONFIG` at the top of the script: spreadsheet ID, roster tab name, column letters, recipient/CC, and `WINDOW_FWD_DAYS` (default 7).
