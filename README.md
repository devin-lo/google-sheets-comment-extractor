# google-sheets-comment-extractor
Extracts all comments in a Spreadsheet file from Google's in-built comment feature, and compiles feedback files to Google Drive, to folders organized for uploading to a Moodle course website.

This will be updated with additional instructions, as Google Apps Script cannot 100% do all the tasks needed.
The anchor IDs can only be matched by hacking and slashing in Developer Mode of your browser and an external text editor.

As well, this guide assumes a particular format for the source values of your grading.
For a README on how to set up a grading workbook from the Moodle grading worksheet, please see here.

Feel free to modify the code and these steps to a workflow that works best for you. But for me, this workflow was the best (due to the limitations from Google Apps Script, and me having more familiarity with Excel than with coding)

## Step 1 - Extract all comments to a sheet
This step assumes that you have created a spreadsheet called Comments.
Use the function listComments() to extract all the comments to that sheet. There will be 3 columns: Anchor ID, value of the cell that had the comment, and the comment text itself. Unfortunately, having only the anchor ID is not helpful, as you don't have the corresponding row and column of the comment, nor do you know which student the comment belongs to.

## Step 2 - Extract anchor UID mapping to sheet, row and col from Developer Mode
Unfortunately, this step cannot be scripted.
First, go to your Comments sheet, select any of the anchor IDs, and copy that value. You'll need this to find the info we're going to extract.
Then, bring up Developer Mode in your browser. In the Elements screen, do a search for your anchor ID. It should pull up something like:


I believe this is what Google is using to lay-out the screen, and this is the only place in the whole page where the anchor IDs get used.

Copy this to your text editor of choice. I prefer VSCode for this step.
You need to cut away at the front and the end of this block of text, to get a JSON object.
These parts are extraneous:


Then, this still needs to be modified more. In VSCode, I use this extension to sort the JSON (https://marketplace.visualstudio.com/items?itemName=richie5um2.vscode-sort-json), that way it gets arranged in a way that is easy to read.

All the comment anchor entries look something like:

[
    34070425,
    "[\"1000930816\",[\"1437253575\",61,62,68,69],2,[{}]]"
]

The first element 34070425 is common to all of the entries. 1000930816 is the anchor ID. 1437253575 is the sheet ID (same as the gid that is visible in your browser's address bar when you are on a certain sheet). 62 and 69 are the row and column numbers respectively.

When the JSON has been sorted, then these entries related to the comments appear between an entry that has a timezone, and  a different entry that is very long. It will be obvious, but you can refer to the pictures below.

(two images for showing the start and end)

Select all of these JSON array items (I just used my cursor), and paste to a new file.
Now, you'll need to strip away all the extra characters from the file (you can just use find and replace), so that the example entry from before would now look like this:

1000930816,1437253575,61,62,68,69

Optional: you can replace the gid with the sheet ID, or a different identifier. I did this to faciliate mapping the comments to the person they belong to.

Once it is all stripped, you basically have created a CSV file, and you can copy-paste this to your Google Sheets file in a new sheet (you could call it Mapping).
Use text-to-columns to separate them into columns.

Remember, only columns A, B, D, and F will be relevant. They are Anchor ID, Sheet, Row and Column respectively. You can delete the other two columns as I found they were junk values for me.

## Step 3 - Map the anchor IDs in Comments to the row and col in Mapping, and Map a person's name to each comment
This step is certainly debatable, but I have done this in order to centralize all the information I need to match up every comment to the person it belongs to.
This step is achieved with only Excel formulas.

First, I edit Mapping, to match each anchor ID to a student's name. I use INDEX and INDIRECT formulas in conjunction to achieve this.

Example formula:
=INDEX(INDIRECT(B2&"!$1:$1"),1,D2)

What this does:

INDIRECT(B2&"!$1:$1") is getting the first row of the spreadsheet, the name of which is the value in cell B2.

Then, INDEX is getting from that particular range, the value in row 1, column (value stored in D2).

Then, I named the whole table in Mapping as the named range "CommentMapping", to facilitate VLOOKUP formulas later.

Then, in Comments, I have three additional columns for name, row, and column.
These will all use VLOOKUP on the CommentMapping range, to map from the Anchor ID to those three values, respectively.

Example:

Name has =VLOOKUP(B2,CommentMapping,5,FALSE)

Row has =VLOOKUP(B2,CommentMapping,3,FALSE)

Col has =VLOOKUP(B2,CommentMapping,4,FALSE)

B2 is the value of the anchorID. 3, 4, and 5 correspond to the columns with those values in Mapping. The last parameter is always FALSE, otherwise there will be strange results from the VLOOKUP.

Now, your Comment sheet is all set up, and ready to be used by the code to 

## Step 3 - prepare folder structure in Google Drive
First, download all submissions from Moodle, which creates a ZIP file. Inside, there should already be a folder structure that Moodle can also use to upload feedback in one paylod.

Note: this step is dependent on the ZIP downloading properly. More often than not, I've experienced issues downloading all submissions at once. If this occurs, you will have to break up the download into multiple smaller selections of submissions.

The expected structure is:
ZIP file > each student's folder > any files as feedback

The student folder name is formatted as:
First Last_participantId_assignsubmission_file_

In order to reduce the file size of your upload in the final step, you should use your computer's Command Prompt or Terminal to recursively delete all the student submission files.

In Windows, the command would be something like:

del /s /q *.pdf

(Source https://stackoverflow.com/questions/12748786/delete-files-or-folder-recursively-on-windows-cmd)

Run this from the folder level that is containing all the student folders.
And, make sure to run this command several times to cover all these common file formats: pdf, docx, doc, jpg, png, txt, zip, etc.

Once all cleared, you can re-zip the folder structure, and upload this to Google Drive.
Then, use the ZIP Extractor Add-in (link: https://workspace.google.com/u/1/marketplace/app/zip_extractor/824911851129) to do the extraction into Google Drive.

This tool is very good, as it can also show you any remaining student submission files that you may have missed.

Finally, before the next step, go to the root folder where the student folders are contained, and copy the folder ID from your browser's address bar, as you'll need this for the next step.

## Step 4 - extract folder IDs recursively (using another person's script)
I used this script from nk-gears (who based it off of another script by mesgarpour), without any modification other than to put my own folder ID in line 14. Therefore, I did not include this code in my repository.
https://gist.github.com/nk-gears/c03f3a10120f45f278cfb53b13932e7f

Run listFolers while open to a new spreadsheet. I named it "folders" in my case.

Then, you will need to do some clean-up on this sheet.
The columns Full Path, Date, Last Updated, Description and Size were of no use to me.
Additionally, I butchered the Name column, by using text-to-columns to extract the name. Set the delimiter to be underscores, and only keep the first column of the result.
Then, you also need to butcher the URL - just do a find and replace on this sheet to remove "https://drive.google.com/drive/folders/"
Optional: You can select the whole sheet, and select "Remove link" to get rid of the hyperlinks (the option might be hidden under "View more cell actions")

## Step 5 - Generate feedback files
Now that we finally got all of the set-up tasks out of the way, we can generate our individual feedback files.

Assumption: please create a named range for each section. For my case, I have SecMnames and SecNnames. The named ranges I used are 1 row tall, so the code reflects this for extracting the values. If you use a named range that is 1 column wide instead, then you'll have to rewrite the code.

The code does the following:
- duplicates a Feedback Template sheet
- copypastes student's numerical feedback from the master marking sheet
- inserts the comments using the values from the Comments sheet (that's why we did all that preparatory work!)
- exports the feedback sheet as a PDF, automatically saving it to the correct folder for that student

The code is written such that:
1. It should be able to change between different sections of a course (which will have their own folder structures). For the code in the repository, it's hardcoded to deal with only sections M and N.
2. If the script times out due to Google's limitations, you should be able to resume where you left off. This is due to suggestions that I read on StackOverflow to store some values in the script properties as a way to work around the Google Apps Script timeout of 6 minutes for free users.
3. Because a GET request has to be made to do the export to PDF, we may accidentally hit a rate limit from Google (100 requests per 100 seconds, as stated here https://developers.google.com/analytics/devguides/reporting/mcf/v3/limits-quotas). I have implemented a retry mechanism that incrementally increases the time to wait between retries. It will make 3 re-attempts to export the PDF, before logging an error.

## Step 6 - upload to Moodle
Go to Google Drive, and download the folder that directly contains the student folders.
Google Drive will automatically zip up the folder.
Then, just upload to Moodle with the option "Upload multiple feedback files in a zip"

## Step 7 - redo individual feedback files
The tool has the possibility of failing for some students if the export API fails all retry requests, or the folder ID was somehow not found in the earlier step, or if you have to re-export a regraded assignment.
You can do this using the createSingleFeedbackPage function. In my file, I had a sheet called "Feedback Input" where I had cell B1 for the section letter, and cell B2 for the student name (which was a drop-down validated from a named range)

** NOTE: comments do not auto-update to the Comments sheet, so if you edited the comments, you may want to edit the individualized feedback spreadsheet.

## Acknowledgement
In addition to the documentation provided by Google for Apps Script and Google API, I used the following sites to help figure out the correct syntax and correct features of the Google Apps Script API and library to make use of.

Extracting comments - https://stackoverflow.com/questions/66103464/export-google-docs-comments-into-google-sheets-along-with-highlighted-text


Turn an object into a JSON - https://stackoverflow.com/questions/30895943/log-javascript-object-as-string-google-app-scripts


Idea that I followed to work around script timeout - https://pulse.appsscript.info/p/2021/08/an-easy-way-to-deal-with-google-apps-scripts-6-minute-limit/


How to use the script properties - https://stackoverflow.com/questions/14450819/script-runtime-execution-time-limit


TextFinder class for dealing with searching a sheet in Google Apps Script - https://stackoverflow.com/a/55769313


Example code for generating PDF from a Google document - https://developers.google.com/apps-script/samples/automations/generate-pdfs


Clarifies some of the query parameters for the GET export request - https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70


How to safely send a file to the Trash bin in Google Drive - https://stackoverflow.com/a/41319000

I also used my own auto-attendance-maker project from 2019-20 for guidance, as I haven't coded in Google Apps Script since finishing that earlier project.