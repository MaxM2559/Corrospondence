# Corrospondence
Functions created to help create corrospondence sheets between industry
codes such as corresponding NAICS, HS, ISIC, CPC, and more.

These functions where created to mainly help clean and harmonize data
across different corrospondence sheets between industry codes, such as the
NAICS code, HS code, and more. A copy of the masterlist I created with these
functions will be posted on my GitHub.

# General Notes

Any changes made to Excel files from Python can NOT be reversed on Excel, so
always create a copy of the Excel file or sheet before making any changes.

Always save and close the Excel file before marking any alterations with Python.

Working with Microsoft Excel through Python can be very slow at times. I
usually implement several print statements in my function to check if my
program is running correctly.

When working with new Excel files, sometimes the file must first be opened
and saved before Python can alter it. Also, be sure to enable editing for
said Excel files.

Python will only alter the most recently selected and saved Excel sheet unless
otherwise specified.
