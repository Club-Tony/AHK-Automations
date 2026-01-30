================================================================================
                    DSRF to CSV Export Tool - README
================================================================================

OVERVIEW
--------
This tool fetches shipping data from an Intra Shipping Request Form (DSRF) and
exports it to a CSV file that can be imported into UPS WorldShip.

REQUIREMENTS
------------
- Windows 10 or later (uses built-in curl.exe)
- Active Intra session in your browser
- Valid session cookies

FILES
-----
- DSRF-Export.bat      : Main script - double-click to run
- cookies.txt          : Your session cookies (YOU MUST CREATE THIS)
- cookies.txt.template : Example template for cookies.txt
- README.txt           : This file
- dsrf_export.csv      : Output file (created after running)

SETUP - Getting Your Cookies
-----------------------------
1. Open Firefox and log into Intra
2. Navigate to any Shipping Request Form page
3. Press F12 to open Developer Tools
4. Go to the "Network" tab
5. Refresh the page (F5)
6. Click on any request in the list
7. In the right panel, find "Request Headers"
8. Find the "Cookie:" line
9. Copy the ENTIRE cookie string (it's very long)
10. Create a file named "cookies.txt" in this folder
11. Paste the cookie string into cookies.txt (single line, no line breaks)
12. Save the file

The cookie string looks something like this:
CurrentProfile=ProfileId=1; AWSALB=xxx...; .AspNet.ApplicationCookie=xxx...

USAGE
-----
1. Make sure cookies.txt exists with valid cookies
2. Double-click DSRF-Export.bat
3. Enter the PK# when prompted (e.g., PK438893)
4. The script will fetch data and create dsrf_export.csv
5. Import the CSV into UPS WorldShip

TROUBLESHOOTING
---------------
"cookies.txt not found"
  - Create the cookies.txt file per the setup instructions above

"API call failed" or "No data returned"
  - Your cookies may have expired - refresh them from browser DevTools
  - Make sure you're logged into Intra
  - Check that the PK# is correct

"No data returned for this PK#"
  - The PK# doesn't exist or has no shipping data
  - Double-check the asset ID

CSV OUTPUT FORMAT
-----------------
The CSV contains these columns:
  Name, Company, Address1, Address2, City, State, Postal, ServiceType

NOTES
-----
- Cookies expire periodically - you'll need to refresh them
- This tool requires an active Intra session in your browser
- Only fetches shipping-related fields from the form

================================================================================
