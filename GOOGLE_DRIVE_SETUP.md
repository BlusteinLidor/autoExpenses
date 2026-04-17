# Google Drive Integration Setup

This project can upload generated monthly and yearly Excel files to the connected Google Drive account.

## 1) Create OAuth credentials

1. Open [Google Cloud Console](https://console.cloud.google.com/).
2. Create/select a project.
3. Enable **Google Drive API**.
4. Configure OAuth consent screen.
5. Create an **OAuth Client ID** (Web application).
6. Add an authorized redirect URI that points to your API callback route:
   - Example: `http://localhost:8000/drive/connect/callback`

## 2) Set environment variables

Add these variables to your local `.env`:

- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`
- `GOOGLE_OAUTH_REDIRECT_URI` (must match the Google OAuth redirect URI)
- `DASHBOARD_BASE_URL` (optional, defaults to `http://localhost:3000`)

## 3) Install dependencies

Install Python dependencies from `requirements.txt` (includes Google API libs).

## 4) Connect from dashboard

1. Open the dashboard.
2. In **Google Drive** section click **Connect Google Drive**.
3. Approve access and return to dashboard.

## Upload behavior

- Monthly upload target path: `AutoExpenses/{year}/{month}` (month is zero-padded).
- Yearly workbook target path: `AutoExpenses/{year}`.
- Two files are uploaded after each successful monthly finalize:
  - `expenses_output_{year}_{MM}.xlsx` (monthly workbook, in month folder)
  - `expenses_output_{year}_total.xlsx` (yearly workbook, in year folder)
- Local save always stays the source of truth.
- If Drive upload fails, monthly finalize still succeeds locally and the UI shows a warning.

## Validation checklist

- Connect succeeds and dashboard shows connected account.
- Run monthly flow and verify both files appear in Google Drive under expected folder.
- Re-run same month and verify files are updated (not duplicated by name).
- Disconnect and verify uploads are skipped.
