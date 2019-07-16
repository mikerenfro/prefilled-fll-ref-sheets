% FLL Robot Game Score Sheet Filler

# Requirements

- Python 3
- Pandas Python module
- Reportlab Python module
- pdfrw Python module

Developed on Python 3.7, Pandas 0.24.1, 3.5.13, pdfrw 0.4

# Instructions

## CSV of teams' robot game schedule

Have a CSV of teams' robot game schedule saved in the current directory. Edit the `data =` line in the Python script to match.

The CSV should be literally comma-separated, with 7 columns labeled:

- Team
- Time1
- Table1
- Time2
- Table2
- Time3
- Table3

The reference CSV used was copied from an Excel spreadsheet used to make a printed team schedule.
The `TimeX` columns were originally in Excel time format, and ended up expressed as floating point numbers between 0 and 1 indicating the time of day.
That makes it easier to sort times to figure out which table has a given team's round 1, 2, or 3.

## Score sheet template

Have a template score sheet saved in the current directory. Edit the `template_path` variable in the Python script to match.

## Other Python script changes

Adjust table colors and table numbers in the Python script as necessary.
