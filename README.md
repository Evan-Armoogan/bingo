# Bingo

Run the script by calling python and running `main.py`. You have to pass arguments for
`--spreadsheet_ids`. Give a list of spreadsheet IDs (separated by spaces) for each team's
game spreadsheet. Note, the spreadsheet ID can be found in the URL of the Google Sheet
after the `d/` until the next `/`. Note: each spreadsheet must have a sheet named `Game`.
Other sheets may be present, the script won't modify them. Note that this script will
automatically create the bingo room on `bingosync.com`.

In order to connect with Google Sheets, you need to paste the Google Service Account
credentials JSON into the file `google_sheets/GServAcc`. All game spreadsheets must
be shared with this Service Account.

Details on rules, template spreadsheets to come.
