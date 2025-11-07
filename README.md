# Email Import Script for Google Sheets

A Google Apps Script that imports emails from Gmail to Google Sheets with real-time progress tracking, email count verification, and flexible sheet support.

## Features

✅ **Batch Processing** - Imports up to 300 emails per batch (no timeouts)
✅ **Email Verification** - RFC 822 header verification to ensure only your sent emails are imported
✅ **Live Progress Dialog** - See real-time progress with threads, emails collected, and time elapsed
✅ **Email Count Verification** - Verifies total emails imported matches expected count
✅ **Flexible Sheets** - Works with any sheet name (e.g., 09/25, 08/25, October 2025)
✅ **Resume Capability** - If interrupted, resume from last batch

## Installation

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Create two files:
   - **`Code.gs`** - Paste content from Code.gs
   - **`ProgressMonitor.html`** - Paste content from ProgressMonitor.html
4. Save and close
5. Refresh your sheet (F5)

## Usage

1. Create a sheet for each month (e.g., "09/25" for September 2025)
2. Click on the sheet
3. Click **📧 Monthly Import → 📅 Import Month**
4. Enter month number (1-12) and year
5. Watch the progress dialog
6. Get verification when complete

## Configuration

Edit the top of `Code.gs`:
```javascript
const CONFIG_MONTHLY = {
  WORK_EMAIL_ADDRESS: 'your-email@domain.com',  // Change this!
  SPREADSHEET_NAME: 'TR_Master_Directory',
  METADATA_SHEET_NAME: 'Metadata',
  BATCH_SIZE: 300,  // Increase/decrease based on your needs
  BODY_LIMIT: 25000,  // Plaintext body character limit
  // ...
};
```

## Menu Options

- 📅 **Import Month** - Start importing emails
- 📊 **View Progress** - See which months have been imported
- 🗑️ **Clear & Reset** - Delete all data and reset progress

## How It Works

1. **Searches Gmail** for emails sent FROM your work address in the specified month
2. **Processes threads** and verifies sender using RFC 822 headers
3. **Collects email data** (ID, From, To, Subject, Body, Date)
4. **Batches writes** to avoid timeouts
5. **Tracks counts** and verifies all emails were imported
6. **Resumes** if interrupted by processing next batch

## Features Breakdown

### Email Verification
- Uses RFC 822 raw message headers for 100% certainty of sender
- Filters out received emails in multi-party threads
- Handles Gmail aliases correctly

### Performance
- 300 emails per batch (configurable)
- Single batch write (1 API call, not 300)
- No re-verification of existing emails

### Verification
- Counts expected emails during import
- Verifies actual count in sheet at end
- Shows "Mismatch" if counts differ

## Troubleshooting

**"This month is already imported"**
- Delete the entry in the "Metadata" sheet and try again
- Or use "Clear & Reset" to start fresh

**"Timeout error"**
- Decrease `BATCH_SIZE` in config (e.g., 200 instead of 300)

**No emails found**
- Verify your `WORK_EMAIL_ADDRESS` is correct
- Check that you actually sent emails that month

## License

MIT - Feel free to use and modify!

## Author

Created with ❤️ for email management and organization.
