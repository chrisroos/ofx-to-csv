## Intro

My bank, HSBC, offers an OFX (and QIF) download.  I want to play around with storing my transactions in a Google Spreadsheet.  This utility will convert the OFX to CSV and then I'm hoping to create a Google Spreadsheet Script that removes duplicates.

## TODO

* Add a note columns to the spreadsheet
* Don't delete if the transaction contains a note
* Add tests around the remove-duplicates.js