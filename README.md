# todays-yield-curve

This exercise pulls in the most recent treasury rates available for free online instead of paying for an API and then uploads to an excel file and craetes a yield curve and then describes its slope

The website used to scrape this information is "https://ustreasuries.online"

Steps to setting up the Excel file:
1. Macros are not enabled for this workbook so you can just download the spread sheet and save down to a folder where you won't move it.
2. Open the excel file and go to settings (Alt F T)
3. Go to the "Trust Center Settings" tab and open the "Trust Center Settings"
4. Go to trusted locations
5. Click "Add New Location"
6. Click "Browse"
7. Locate the folder the excel file is in and hit "OK"
8. Check the box that says "Sub folders of this location are also trusted"
9. Hit "OK"

Steps to setting up the code:
1. Save the file down with the Execl file
2. Copy the location of the excel file in the folder (in the file explorer highlight the Excel file then hit (Ctrl Shift C) to copy the location)
3. Open the coding enviornment
4. Go to ln 122
5. Replace the r"" with the location of the Excel file
6. Ensure only one set of quotes are around the location
7. Save
8. Run
