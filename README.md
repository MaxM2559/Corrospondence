# Correspondence
The first-ever tools to create correspondence sheets for international industry codes such as NAICS, HS, ISIC, CPC, and their variations. To my knowledge, there did not exist a single correspondence sheet capable of translating codes across all variations of industry codes. Only sheets corresponding to the same code across years or just two different codes existed.

Correspondence_Functions.py contains the functions created to help clean and harmonize data across various correspondence sheets, which I searched the web for.

A copy of the harmonized list is included here. After some data exploration, I found that there is no single best list possible. The data becomes more and more abstracted as it moves further from the starting list. At a certain point, codes near the end of the list are hundreds of characters long.

My solution is to create different lists based on different needs. A master list is not feasible, so creating shorter lists with different starting points is the best solution I've found. The functions in the Python file are designed to facilitate this process.

![NAICS_2017_codes](https://github.com/user-attachments/assets/a6c04a95-e06c-4c2c-b6b8-d3e7496dc1f1)
![HS_2017_codes](https://github.com/user-attachments/assets/5b49ca18-f93a-4999-a107-9f86d74be033)

# Notes for Future Work

Any changes made to Excel files from Python cannot be reversed in Excel, so always create a copy of the Excel file or sheet before making any changes.

Always save and close the Excel file before making any alterations with Python.

Working with Microsoft Excel through Python can be very slow. I usually implement several print statements in my functions to check if my program is running correctly.

When working with new Excel files, sometimes the file must first be opened and saved before Python can alter it. Also, be sure to enable editing for these Excel files.

Python will only alter the most recently selected and saved Excel sheet unless otherwise specified.
