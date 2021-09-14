# Customer-Info-App-Fastly - Generate a Google Spreadsheet with a report of your account usage.

Install Instructions

How to install/run? - You can make copies of this Google spreadsheet (https://docs.google.com/spreadsheets/d/1xAPnUSq8WnW4_936kjz6pJCsJQbtRRvsAJZTtakb2t0/edit?usp=sharing) and run the app by going to Add on -> Customer Info App Fastly -> Get started. This will throw up an app sidebar in the sheet that you can use to generate the report.

Authorization - You will need to authorize the app to run - ![image](https://user-images.githubusercontent.com/4117801/132954239-5faa3264-2334-4fdc-a9f8-0836b88b873f.png)

Warning - Since this is a Google Doc Script you will see the following warning - <img width="594" alt="Warning - Google docs" src="https://user-images.githubusercontent.com/4117801/132954370-a25ce9f4-636a-40dc-b6f9-3e6cdf9468d2.png">

Advanced - Hit advanced and then continue with the option -- Go to Customer Info App Fastly (unsafe) - <img width="588" alt="Advanced" src="https://user-images.githubusercontent.com/4117801/132954472-f58ae06f-a84c-4bda-8692-f98afb2b086b.png">

Permissions - The app needs the following permissions to generate the report that you need to allow - ![image](https://user-images.githubusercontent.com/4117801/132954510-422f8093-2c90-47d5-945c-59da96e69cdc.png)

Returning to app - After you hit Allow go back to your app and click Add-ons -> Customer Info App Fastly -> Get started - <img width="1127" alt="Get started" src="https://user-images.githubusercontent.com/4117801/132954555-4575618e-4024-4d84-b8ee-a0068bcdb1b0.png">

Usage - This should give you a side bar that into which you need to enter your customer id and Fastly API key and after that scroll to the bottom of the sidebar and hit Submit - <img width="285" alt="Sidebar" src="https://user-images.githubusercontent.com/4117801/132954604-b71aedae-047a-4183-91d4-91084a97e675.png">

Output - This should generate a usage report for all Fastly services in your account for the past 90 days

Various app features - a.) You can change the column title and any of the regular expressions starting row 23 in the Maintenance sheet and the app will search for that expression in your VCL b.) You can fetch all certificates your account uses in the Certificate tab c.) You can get details for specific services in the Data sheet d.) You can add additional columns to the data sheet by adding new ones to the bottom in the maintenance sheet or replacing existing ones d.) You can get region specvific traffic information e.) You can also add API end points from https://developer.fastly.com/reference/api/metrics-stats/historical-stats/ which will appear as new columns but you will need to add them before row 23 and update the VCL row info (where regular expression in VCL begin in our report) since that row shifts down by the whatever API endpoints you call

Caveats - a.) The tool is limited to run a maximum of 29 minutes so if you have too many Fastly services or fetch information over a longer period of time than default 90 days you may have to split your reports and run it multiple times b.) It won't run VCL regex info for C@E services c.) You will need an API key to Fastly and an account on Fastly to be able to generate reports
