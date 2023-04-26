function listComments() {
  // this was adapted from the link below
  // https://stackoverflow.com/questions/66103464/export-google-docs-comments-into-google-sheets-along-with-highlighted-text
  var docId = getDocId(); // from separate variables file that I am not publishing
  var optionalArgs = {
    fields: "*",
    pageSize: 100 // this didn't work for some reason, it defaulted to 20
  }; 
  var comments = Drive.Comments.list(docId, optionalArgs);
  var commentList = [];
  var aList = [], bList = [], cList = [];
  var bListStart = bList.length + 1;
  var aListStart = aList.length + 1;
  var cListStart = cList.length + 1;
  // from reading the documentation on Drive.Comments.list,
  // I realized that nextPageToken is what I need to be able
  // to keep pulling the comments with the limitation of only 20 comments returned by the list API
  var nextPageToken = comments.nextPageToken;
  console.log("Current page token = " + nextPageToken);

  // get all of the comments from the API into an array
  while(nextPageToken != undefined) {
    // Get list of comments
    if (comments.items && comments.items.length > 0) {
      for (var i = 0; i < comments.items.length; i++) {
        var comment = comments.items[i];
        commentList.unshift(comment);
      }
    }
    // this older version of the code below would work if you only needed to call Drive.Comment.list once
    // but, it doesn't pull the anchor values
    // if (comments.items && comments.items.length > 0) {
    //   for (var i = 0; i < comments.items.length; i++) {
    //     var comment = comments.items[i]; 
    //     // add comment and highlight to array's first element 
    //     var contextValue = comment.context === null ? "??" : comment.context.value;
    //     hList.unshift([contextValue]);
    //     cList.unshift([comment.content]);
    //   }
    //   // Set values to A and B
    //   var sheet = SpreadsheetApp.openById(docId).getSheetByName('Comments');
    //   sheet.getRange("A" + hListStart + ":A" + (hListStart+hList.length-1)).setValues(hList);
    //   sheet.getRange("B" + cListStart + ":B" + (cListStart+cList.length-1)).setValues(cList);
    // }
    // hListStart += hList.length;
    // cListStart += cList.length;
    // hList = [], cList = [];
    optionalArgs = {
      pageSize: 100, // this didn't work for some reason, it defaulted to 20
      pageToken: nextPageToken
    };
    comments = Drive.Comments.list(docId, optionalArgs);
    nextPageToken = comments.nextPageToken;
    console.log("Current page token = " + nextPageToken);
  }
  commentList.sort(compareAnchorUid);

  console.log("see comment list now"); // this line could be used to set a breakpoint for debugging.

  // Build the Excel layout of the values that we'll put in the spreadsheets
  // from the array that we constructed
  for (var i = 0; i < commentList.length; i++) {
    var comment = commentList[i];
    var contextValue = comment.context === null ? "??" : comment.context.value;
    if (!comment.deleted) {
      aList.unshift([JSON.parse(comment.anchor).range]); // pull the anchor UID values for mapping the cell location later
      bList.unshift([contextValue]);
      cList.unshift([comment.content]);
    }
  }
  // NOTE: assumes that you have a spreadsheet already created called Comments
  var sheet = SpreadsheetApp.openById(docId).getSheetByName('Comments');
  sheet.getRange("A" + aListStart + ":A" + (aListStart+aList.length-1)).setValues(aList);
  sheet.getRange("B" + bListStart + ":B" + (bListStart+bList.length-1)).setValues(bList);
  sheet.getRange("C" + cListStart + ":C" + (cListStart+cList.length-1)).setValues(cList);
  
}

/**
 * Helper function to sort Comment objects by the anchor
 * anchor is an undocumented object, which has a "range" property that appears to be a uid
 * Although, this ended up being an unnecessary step since uid doesn't directly map to a comment's location
 *
 */
function compareAnchorUid(comment1, comment2) {
  // https://stackoverflow.com/questions/30895943/log-javascript-object-as-string-google-app-scripts
  var anchor1 = JSON.parse(comment1.anchor);
  var anchor2 = JSON.parse(comment2.anchor);
  var uid1 = anchor1.range;
  var uid2 = anchor2.range;
  return uid2 - uid1;
}

/**
 * Function to reset the script properties that are needed for creating feedback pages
 *
 */
function resetProperties() {
  scriptProp.setProperty("currentEntry", 0);
  scriptProp.setProperty("currentSection", "M");
}

/**
 * Call this function if you need to create a single feedback page
 *
 */
function createSingleFeedbackPage() {
  var onSheet = SpreadsheetApp.getActive();
  var inputSheet = onSheet.getSheetByName("Feedback Input");
  var studentName = inputSheet.getRange(2, 2).getValue();
  var sectionLetter = inputSheet.getRange(1, 2).getValue();
  createFeedbackPage(studentName, sectionLetter, false);
}

/**
 * Main function to call for creating all feedback
 *
 */
function createAllFeedback() {
  var onSheet = SpreadsheetApp.getActive();

  // https://pulse.appsscript.info/p/2021/08/an-easy-way-to-deal-with-google-apps-scripts-6-minute-limit/
  // https://stackoverflow.com/questions/14450819/script-runtime-execution-time-limit
  var scriptProp = PropertiesService.getScriptProperties();
  var currSection = scriptProp.getProperty("currentSection");
  while (currSection !== "done") {
    var myRange = onSheet.getRangeByName("Sec" + currSection + "names");
    var myNames = myRange.getValues()[0];
    console.log("Starting " + currSection + " names");
    var i = Math.trunc(scriptProp.getProperty("currentEntry"));
    for (; i < myNames.length; i++) {
      console.log("Entry #" + i + ": " + myNames[i]);
      createFeedbackPage(myNames[i], currSection, true);
      scriptProp.setProperty("currentEntry", i+1);
    }
    if (currSection === "M") {
      scriptProp.setProperty("currentSection", "N");
      scriptProp.setProperty("currentEntry", 0);
      currSection = scriptProp.getProperty("currentSection");
    } else {
      scriptProp.setProperty("currentSection", "done");
      scriptProp.setProperty("currentEntry", 0);
      currSection = scriptProp.getProperty("currentSection");
    }
  }

  // var mRange = onSheet.getRangeByName("SecMnames");
  // var mNames = mRange.getValues()[0];
  // console.log("Starting M names");
  // mNames.forEach(function(value){
  //   createFeedbackPage(value, "M", true);
  // });

  // var nRange = onSheet.getRangeByName("SecNnames");
  // var nNames = nRange.getValues()[0];
  // console.log("Starting N names");
  // nNames.forEach(function(value){
  //   createFeedbackPage(value, "N", true);
  // });
}

/**
 * Drives the creation of feedback page
 * @param studentName
 * @param sectionLetter
 * @param runFromAllFeedback if this is part of the batch, or just a single request
 *
 */
function createFeedbackPage(studentName, sectionLetter, runFromAllFeedback) {
  var onSheet = SpreadsheetApp.getActive();

  // create the feedback sheet
  var outSheet = onSheet.getSheetByName("Feedback");
  if (outSheet) {
    onSheet.deleteSheet(outSheet);
  }
	var tempSheet = onSheet.getSheetByName("Feedback Template");
  if (!tempSheet) {
    throw new Error("Template sheet is missing!"); // eventually come back and hard-code a template sheet.
  }
  onSheet.setActiveSheet(tempSheet, true);
  onSheet.duplicateActiveSheet();
  onSheet.renameActiveSheet("Feedback");
  onSheet.moveActiveSheet(tempSheet.getIndex());
  outSheet = onSheet.getSheetByName("Feedback");

  var markSheet = onSheet.getSheetByName("Sec " + sectionLetter + " INPUT");
  var lastRow = markSheet.getMaxRows();
  var studentMarksFinder = markSheet.createTextFinder(studentName).matchEntireCell(true); // https://stackoverflow.com/a/55769313
  var search_column = studentMarksFinder.findNext().getColumn();
  var marks = markSheet.getRange(1, search_column, lastRow, 1);

  marks.copyTo(outSheet.getRange('C1'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  // copy all comments
  findAllComments(studentName, outSheet);

  var folderId = getFolderId(studentName, outSheet);
  if (!folderId) {
    console.warn("Couldn't find folder ID for student " + studentName + ", remake this file later.");
    return;
  }
  // export PDF to Google Drive
  var pdf = createPDF(studentName, outSheet, folderId, runFromAllFeedback);
  if (!pdf) {
    console.error("PDF export failed for student " + studentName + ", remake this file later.");
  }
}

/**
 * Helper method to find all comments related to a particular student
 * Assumes that student names are unique
 */
function findAllComments(studentName, outSheet) {
  var onSheet = SpreadsheetApp.getActive();
  var commentSheet = onSheet.getSheetByName("Comments");
  var commentFinder = commentSheet.createTextFinder(studentName).matchEntireCell(true);
  var matchCells = commentFinder.findAll();
  for (var i = 0; i < matchCells.length; i++) {
    var matchRow = matchCells[i].getRow();
    var sourceRow = commentSheet.getRange(matchRow, 3).getValue();
    var comment = commentSheet.getRange(matchRow, 5);

    comment.copyTo(outSheet.getRange(sourceRow, 4), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  }

  outSheet.getRange('D:D').activate();
  outSheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

/**
 * Helper method to get the folder ID from the folder ID sheet
 * Assumes that student names are unique
 */
function getFolderId(studentName, outSheet) {
  var onSheet = SpreadsheetApp.getActive();
  var folderSheet = onSheet.getSheetByName("folders");
  var folderFinder = folderSheet.createTextFinder(studentName).matchEntireCell(true);
  var folderFind = folderFinder.findNext();
  if (!folderFind) {
    return null;
  }
  var folderRow = folderFind.getRow();
  var folderId = folderSheet.getRange(folderRow, 2).getValue();
  return folderId;
}

/**
 * Creates a PDF for the customer given sheet.
 * @param {string} studentName - student name
 * @param {object} outSheet - Sheet to be converted as PDF
 * @param {string} folderId - folder ID to save the PDF to
 * @param {boolean} runFromAllFeedback - whether or not this was run as part of the batch
 * @return {file object} PDF file as a blob
 * From https://developers.google.com/apps-script/samples/automations/generate-pdfs
 */

// https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
// PDF OPTIONS****************
  //format=pdf     
  //size=0,1,2..10             paper size. 0=letter, 1=tabloid, 2=Legal, 3=statement, 4=executive, 5=folio, 6=A3, 7=A4, 8=A5, 9=B4, 10=B5  
  //fzr=true/false             repeat row headers
  //portrait=true/false        false =  landscape
  //fitw=true/false            fit window or actual size
  //gridlines=true/false
  //printtitle=true/false
  //pagenum=CENTER/UNDEFINED      CENTER = show page numbers / UNDEFINED = do not show
  //attachment = true/false      dunno? Leave this as true
  //gid=sheetId                 Sheet Id if you want a specific sheet. The first sheet will be 0. others will have a uniqe ID. 
                               // Leave this off for all sheets. 
  // EXPORT RANGE OPTIONS FOR PDF
  //need all the below to export a range
  //gid=sheetId                must be included. The first sheet will be 0. others will have a uniqe ID
  //ir=false                   seems to be always false
  //ic=false                   same as ir
  //r1=Start Row number - 1        row 1 would be 0 , row 15 wold be 14
  //c1=Start Column number - 1     column 1 would be 0, column 8 would be 7   
  //r2=End Row number
  //c2=End Column number
function createPDF(studentName, outSheet, folderId, runFromAllFeedback) {
  const url = "https://docs.google.com/spreadsheets/d/" + SpreadsheetApp.getActiveSpreadsheet().getId() + "/export" +
    "?format=pdf&" +
    "size=0&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=true&" +
    "printtitle=true&" +
    "top_margin=0.5&" +
    "bottom_margin=0.25&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=CENTER&" +
    "attachment=true&" +
    "gid=" + outSheet.getSheetId();

  console.log(url);

  const fileName = studentName + ' A5 feedback.pdf';

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  var blob = null;
  var retryExhausted = 0;
  // introduce a retry mechanism to the GET request
  while (blob === null && retryExhausted < 3) {
    try {
      blob = UrlFetchApp.fetch(url, params).getBlob().setName(fileName);
    } catch(err) {
      console.log("Retry #" + (retryExhausted + 1) + " out of 3");
      Utilities.sleep((retryExhausted + 1) * 3500);
      retryExhausted++;
    }
  }
  if (!blob) {
    console.warn("blob fetch retries exhausted");
    return null;
  }

  // while (blob === null && retryExhausted < 3) {
  // var response = UrlFetchApp.fetch(url, params);
  //  if (response.getResponseCode() != 200) {
  //    console.log(response.getContentText());
  //  } else {
  //    blob = response.getBlob().setName(fileName);
  //  }
  // }
  // if (!blob) {
  //   console.warn("blob fetch retries exhausted");
  //   return null;
  // }

  // var blob = null;
  // var response = UrlFetchApp.fetch(url, params);
  // if (response.getResponseCode() != 200) {
  //   console.log(response.getContentText());
  //   return null;
  // }
  // blob = response.getBlob().setName(fileName);

  // Gets the folder in Drive where the PDFs are stored.
  const folder = DriveApp.getFolderById(folderId);

  // unfortunately, there is currently NO way to just update an existing PDF unlike uploading via Google Drive
  var pdfFiles = folder.getFilesByName(fileName);
  // delete the existing PDF in the folder if being run from the batch, in case the existing file is corrupt/incomplete
  if (runFromAllFeedback) {
    while (pdfFiles.hasNext()) {
      var delFile = pdfFiles.next();
      delFile.setTrashed(true); // https://stackoverflow.com/a/41319000
    }
  }
  var pdfFile = folder.createFile(blob);
  return pdfFile;
}