# Tender_Crawler
## How To Use
1. Create a google spreadsheet
2. Copy function.gs to your appscript which connected to the spreadsheet
3. If you **do not** want to recieve email, please comment **that part of code** and the call line in **Line 220**
4. Set timer to make the script automatically run once or more eachy period
5. Build an official Line account and add your channelAccessToken to properties named "TOKEN"
6. Add your sheet id to properties named "SHEETID", the content would be in `"https://docs.google.com/spreadsheets/d/XXXXXXXXXXXX...XXX/edit?gid=GGGG..."` (The XXXXX...XXXX is your sheet id)
8. Add a worksheet named "keyword"
9. Add a worksheet named "Email" if you need
10. Type your keywords in order in column A
11. Type mail addresses in order in column A to recieve emails after each runtime
12. Run the script for the first time to find if any problems and sync the newest infos
13. Receive your report every day and take the infos to bid

   
