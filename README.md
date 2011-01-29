## Intro

My bank, HSBC, offers an OFX (and QIF) download.  I want to play around with storing my transactions in a Google Spreadsheet.  This utility will convert the OFX to CSV and then I'm hoping to create a Google Spreadsheet Script that removes duplicates.

## TODO

* Make it clearer (e.g. highlight rows) that duplicate rows both contain a note - at the moment it just logs which isn't very obvious
* Add tests around the remove-duplicates.js