# todays-yield-curve

This exercise pulls in the most recent treasury rates available for free online instead of paying for an API and then uploads to an Excel file creates a yield curve and then describes its slope. The slope of the yield curve is calculated by comparing the 10 Year Treasury to the 1 Year Treasury and comparing the delta to different measurements to describe the shape of the slope. 

The website used to scrape this information is "https://ustreasuries.online"

Steps to setting up the Excel file:
1. Macros are not enabled for this workbook so you can just download the spreadsheet and save it down to a folder where you won't move it.
2. Open the Excel file and go to settings (Alt F T)
3. Go to the "Trust Center Settings" tab and open the "Trust Center Settings"
4. Go to trusted locations
5. Click "Add New Location"
6. Click "Browse"
7. Locate the folder the Excel file is in and hit "OK"
8. Check the box that says "Subfolders of this location are also trusted"
9. Hit "OK"
10. Don't change the name of the file or the worksheets

Steps to setting up the code:
1. Save the file down in the same location as the Excel file
2. Copy the location of the Excel file in the folder (in the file explorer highlight the Excel file then hit (Ctrl Shift C) to copy the location)
3. Open the coding environment
4. Go to ln 123
5. Replace the r"" with the location of the Excel file
6. Ensure only one set of quotes is around the location
7. Save
8. Make sure the Excel file is not open when you run the script otherwise it won't work
9. Run the script every time you need to update the yield curve
