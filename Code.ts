// Compiled using ts2gas 3.6.4 (TypeScript 4.2.4)
var exports = exports || {};
var module = module || { exports: exports };
// Compiled using ts2gas 3.6.4 (TypeScript 4.2.4)
var exports = exports || {};
var module = module || { exports: exports };
var exports = exports || {};
var module = module || { exports: exports };
// @ts-ignore
// @ts-ignore
// Logger = BetterLog.useSpreadsheet('1eaOMLHtPLIT6EAV6RUjMkGPBP0e7Lhw2CZbn3aO3gEY');
var ss = SpreadsheetApp.getActiveSpreadsheet();
function onFileOpen() {
  SpreadsheetApp.getUi().createMenu('Functions')
    .addItem('import excel file', 'importXLS')
    .addToUi();
}

function getMAPdata() {
  var extSs = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1hBFCtjdhqQAANB9mNmredwpcQDCFORRDhrygcv9VSv8/edit#gid=1223497952');
  var sheet = extSs.getSheetByName('Sheet1');
  extSs.sort(4);
  var dest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('map1908');
  dest.sort(1);
  var srcVals = extSs.getRange('A1:O12405').getValues();
  var destVals = dest.getRange('A1:AG630').getValues();
  destVals.shift();
  loop1: for (var i = 0; i < destVals.length; i++) {
    var stuRow = destVals[i];
    var idDest = stuRow[0];
    var f = 0; // number of scores found
    loop2: for (var j = 0; j < srcVals.length; j++) {
      var mapRow = srcVals[j];
      var mapDate = mapRow[10];
      var idSrc = mapRow[3];
      if (mapRow[7] == 'MAP' && mapDate == '1908' && idSrc == idDest) {
        var s = mapRow[12];
        if (mapRow[8] == 1) {
          stuRow.splice(32, 1, s);
          f++;
        }
        else if (mapRow[8] == 20) {
          stuRow.splice(33, 1, s);
          f++;
        }
        if (f > 1) {
          break loop2;
        }
      }
    }
  }
  var destRange = dest.getRange(2, 1, destVals.length, destVals[0].length);
  destRange.setValues(destVals);
}
function getExcelData() {
  var folderId = '1pjTSNCCVxKajDMNAZnPETR8WmthlgcOX';
  var convertedFolderId = '1bEZ2V7SrFoXUDOQ694q5Ns31qMuVQEy1';
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var folderIdArray = [convertedFolderId];
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    if (fileName.indexOf('xlsx') != -1) {
      var fileId = file.getId();
      var fileBlob = file.getBlob().setContentTypeFromExtension();
      var converted = convertExcel2Sheets(fileBlob, fileName, folderIdArray);
      var sheet = ss.getSheetByName('allPupils');
      var convId = converted.getId();
      converted = SpreadsheetApp.openById(convId);
      break;
    }
  }
  var newData = converted.getSheetByName('Sheet1').getDataRange().getValues();
  for (var i = 0; i < newData.length; i++) {
    var element = newData[i];
    element.splice(0, 1, element[0].toString());
  }
  var destSheet = ss.getSheetByName('allPupils');
  var destRange = destSheet.getRange(1, 1, newData.length, newData[0].length);
  destSheet.getRange(1, 1, 1000, 50).clearContent();
  SpreadsheetApp.flush();
  destRange.setValues(newData);
  var headersAndFormulas = [[
    '=ArrayFormula(iferror(vlookup(tchrNum, teacherCodes!$B$1:$H, 7,false),if(row($M1:$M) = 1, "teachEmail","")))	',
    '=ArrayFormula(iferror(vlookup(tchrNum,{teacherCodes!$B$1:$I34 }, 8,false),if(row($M$1:$M) = 1,"teachName","")))	',
    '=ArrayFormula(if(row($Z$1:$Z) <> 1, if(isBlank($A$1:$A),,if(($M$1:$M = 21) + ($M$1:$M = 100) + ($M$1:$M = 105) + sum($S$1:$S = "X") > 0, 1, 0)),"sdc||rsp"))	',
    // '=ArrayFormula(if(row(A1:A)=1,"nmJdob",regexreplace(if(isblank(A1:A),, REGEXREPLACE(C1:C & D1:D, "[ \'-]", "") & right(year(G1:G),2) & days(\"12/31/\"&(year(G1:G)-1), G1:G)),"-","")))',
    '=ArrayFormula(if(row(id)=1,"nmJdob",regexreplace(if(isblank(id),, REGEXREPLACE(lastName & firstName, "[ \'-]", "") & right(year(dob),2) & days("12/31/"&(year(dob)-1), dob)),"-","")))',
    '=ArrayFormula(if(isblank(id),, regexreplace(lastName & "_" & firstName & "_" & id, "[ \'-]", "")))',
    '=ArrayFormula(if(isblank(id),, REGEXREPLACE(lastName & "_" & firstName & "_dob_" & dob, "[ \'-]", "")))',
    '=ArrayFormula(if(isblank(id),, REGEXREPLACE(lastName & "_" & firstName, "[ \'-]", "")))',
    '=ArrayFormula(if(isblank(id),, REGEXREPLACE(firstName & "_" & lastName, "[ \'-]", "")))',
    '=ARRAYFORMULA((H1:H)&", "&(V1:V))'
  ]];
  ss.getRangeByName('grade');
  var formulaRng = destSheet.getRange(1, newData[0].length + 1, 1, headersAndFormulas[0].length);
  formulaRng.setFormulas(headersAndFormulas);
  ss.getSheetByName('frequency distribution').getRange("E14").setValue(new Date());
}
/**
 * Convert Excel file to Sheets
 * @param {Blob} excelFile The Excel file blob data; Required
 * @param {String} filename File name on uploading drive; Required
 * @param {Array} arrParents Array of folder ids to put converted file in; Optional, will default to Drive root folder
 * @return {Spreadsheet} Converted Google Spreadsheet instance
 **/
function importXLS() {
  var folderID = "1CZK4YhSS3uiihM-7D-m3sgZWVATWfBK0"; // Added // Please set the folder ID of "FolderB".
  var folder = DriveApp.getFolderById('1CZK4YhSS3uiihM-7D-m3sgZWVATWfBK0');
  var files = DriveApp.getFolderById(folderID).getFiles();
  while (files.hasNext()) {
    var xFile = files.next();
    var name = xFile.getName();
    if (name.indexOf('xlsx') > -1) {
      var ID = xFile.getId();
      var xBlob = xFile.getBlob();
      var convertedFile = {
        title: (name + '_converted_' + new Date().toUTCString()).replace(/\.xlsx/g, ""),
        parents: [{ id: folderID }] //  Added
      };
      var file = Drive.Files.insert(convertedFile, xBlob, {
        convert: true
      });
      var fileId = file.id;
      // Drive.Files.remove(ID); // Added // If this line is run, the original XLSX file is removed. So please be careful this.
    }
  }
  var converted = DriveApp.getFileById(fileId);
  var convertedSS = SpreadsheetApp.openById(fileId);
  var newData = convertedSS.getSheetByName('Sheet1').getDataRange().getValues();
  for (var i = 0; i < newData.length; i++) {
    var element = newData[i];
    element.splice(0, 1, element[0].toString());
  }
  var destSheet = ss.getSheetByName('allPupils');
  var destRange = destSheet.getRange(1, 1, newData.length, newData[0].length);
  destSheet.getRange(1, 1, 1000, 50).clearContent();
  SpreadsheetApp.flush();
  destRange.setValues(newData);
  var headersAndFormulas = [[
    '=ArrayFormula(iferror(vlookup($M1:$M, teacherCodes!$B$1:$H, 7,false),if(row($M1:$M) = 1, "teachEmail","")))	',
    '=ArrayFormula(iferror(vlookup($M1:$M,{teacherCodes!$B$1:$I34 }, 8,false),if(row($M$1:$M) = 1,"teachName","")))	',
    '=ArrayFormula(if(row($Z$1:$Z) <> 1, if(isBlank($A$1:$A),,if(($M$1:$M = 21) + ($M$1:$M = 100) + ($M$1:$M = 105) + sum($S$1:$S = "X") > 0, 1, 0)),"sdc||rsp"))	',
    '=ArrayFormula(if(row(A1:A)=1,"nmJdob",regexreplace(if(isblank(A1:A),, REGEXREPLACE(C1:C & D1:D, "[ \'-]", "") & right(year(G1:G),2) & days(\"12/31/\"&(year(G1:G)-1), G1:G)),"-","")))',
    '=ArrayFormula(if(isblank(id),, regexreplace(C1:C & "_" & firstName & "_" & A1:A, "[ \'-]", "")))',
    '=ArrayFormula(if(isblank(id),, REGEXREPLACE(C1:C & "_" & firstName & "_dob_" & dob, "[ \'-]", "")))',
    '=ArrayFormula(if(isblank(id),, REGEXREPLACE(C1:C & "_" & firstName, "[ \'-]", "")))',
    '=ArrayFormula(if(isblank(id),, REGEXREPLACE(D1:D & "_" & lastName, "[ \'-]", "")))',
    '=ARRAYFORMULA((H1:H)&", "&(V1:V))'
  ]];
  
  
  var formulaRng = destSheet.getRange(1, newData[0].length + 1, 1, headersAndFormulas[0].length);
  formulaRng.setFormulas(headersAndFormulas);

  var saiFormulaRange = destSheet.getRange('AP1');
  saiFormulaRange.setFormula('=arrayformula(if(row(AP1:AP)=1,"sai_exceptions",(iferror(VLOOKUP(A1:A, SAI_exceptions!A1:B, 2, false),0))))');

  var saiFormulaRange = destSheet.getRange('AQ1');
  saiFormulaRange.setFormula('=arrayformula(if(row(AQ1:AQ)=1,"sai_filter",(n(not(isblank(AC1:AC)) +  n(AC1:AC <> 240)) + AP1:AP) >= 2))');

  ss.getSheetByName('frequency distribution').getRange("E14").setValue(new Date());
  converted.setTrashed(true);
}
function convertExcel2Sheets(excelFile, filename, arrParents) {
  var parents = arrParents || []; // check if optional arrParents argument was provided, default to empty array if not
  //   if ( !parents.isArray ) parents = []; // make sure parents is an array, reset to empty array if not
  // Parameters for Drive API Simple Upload request (see https://developers.google.com/drive/web/manage-uploads#simple)
  var uploadParams = {
    method: 'post',
    contentType: 'application/vnd.ms-excel',
    contentLength: excelFile.getBytes().length,
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: excelFile.getBytes()
  };
  // Upload file to Drive root folder and convert to Sheets
  // @ts-ignore
  var uploadResponse = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v2/files/?uploadType=media&convert=true', uploadParams);
  // Parse upload&convert response data (need this to be able to get id of converted sheet)
  var fileDataResponse = JSON.parse(uploadResponse.getContentText());
  // Create payload (body) data for updating converted file's name and parent folder(s)
  var payloadData = {
    title: filename,
    parents: []
  };
  if (parents.length) { // Add provided parent folder(s) id(s) to payloadData, if any
    for (var i = 0; i < parents.length; i++) {
      try {
        var folder = DriveApp.getFolderById(parents[i]); // check that this folder id exists in drive and user can write to it
        payloadData.parents.push({ id: parents[i] });
      }
      catch (e) { } // fail silently if no such folder id exists in Drive
    }
  }
  // Parameters for Drive API File Update request (see https://developers.google.com/drive/v2/reference/files/update)
  var updateParams = {
    method: 'put',
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
    contentType: 'application/json',
    payload: JSON.stringify(payloadData)
  };
  // Update metadata (filename and parent folder(s)) of converted sheet
  // @ts-ignore
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v2/files/' + fileDataResponse.id, updateParams);
  return SpreadsheetApp.openById(fileDataResponse.id);
}
/**
 * Sample use of convertExcel2Sheets() for testing
 **/
function testConvertExcel2Sheets() {
  var xlsId = "0B9**************OFE"; // ID of Excel file to convert
  var xlsFile = DriveApp.getFileById(xlsId); // File instance of Excel file
  var xlsBlob = xlsFile.getBlob(); // Blob source of Excel file for conversion
  var xlsFilename = xlsFile.getName(); // File name to give to converted file; defaults to same as source file
  var destFolders = []; // array of IDs of Drive folders to put converted file in; empty array = root folder
  var ss = convertExcel2Sheets(xlsBlob, xlsFilename, destFolders);
  Logger.log(ss.getId());
}
function test() {
  var sheet = ss.getSheetByName("allPupils");
  Logger.log(sheet.getLastRow());
}
function addTeacherNames() {
  var sheet = ss.getSheetByName("allPupils");
  var sheetNames = ss.getSheetByName("teacherCodes");
  // find column where teacher number is and And check to see if the column already exists or the name 
  var sheetHeadings = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  if (sheetHeadings[0].indexOf("Tchr Num") != -1) {
    var codeCol = sheetHeadings[0].indexOf("Tchr Num") + 1;
    if (sheetHeadings[0].indexOf("teacherName") != -1) {
      var nameCol = sheetHeadings[0].indexOf("teacherName") + 1;
    }
    else {
      var nameCol = sheet.getLastColumn() + 1;
    }
  }
  else {
    return;
  }
  var codeValues = sheet.getRange(2, codeCol, sheet.getLastRow() - 1, 1).getValues();
  var codeLookupValues = sheetNames.getRange(2, 1, sheetNames.getLastRow() - 1, 2).getValues();
  // transpose the names and ID number array easy look up
  var codeLookupTrans = [];
  for (var i = 0; i < codeLookupValues[0].length; i++) {
    for (var j = 0; j < 1; j++) {
      var thisItem = codeLookupValues[j][i];
      codeLookupTrans.push([thisItem]);
      for (var k = 1; k < codeLookupValues.length; k++) {
        var thisItem = codeLookupValues[k][i];
        codeLookupTrans[i].push(thisItem);
      }
    }
  }
  Logger.log(codeLookupTrans);
  var array = [];
  for (var i = 0; i < codeValues.length; i++) {
    var code = parseInt(codeValues[i][0].toString(), 10);
    var codeIndex = codeLookupTrans[1].indexOf(code);
    var teacherName = codeIndex != -1 ? codeLookupTrans[0][codeIndex] : "";
    array.push([teacherName]);
  }
  var dest = sheet.getRange(2, nameCol, sheet.getLastRow() - 1, 1);
  dest.setValues(array);
  var heading = sheet.getRange(1, nameCol, 1, 1);
  heading.setValue("teacherName");
}
function findPaights() {
  addTeacherNames();
  var sheet = ss.getSheetByName("allPupils");
  var sheetSeis = ss.getSheetByName("seisPaight");
  // find the headings with ssid and percent OUT in them
  var headingsSeis = sheetSeis.getRange(1, 1, 1, sheetSeis.getLastColumn()).getValues();
  var ssidSeisIndex = headingsSeis[0].indexOf("Student SSID");
  var percentOutSeisIndex = headingsSeis[0].indexOf("Percent OUT Regular Class");
  var rangeSsidSeis = sheetSeis.getRange(2, ssidSeisIndex + 1, sheetSeis.getLastRow() - 1, 1);
  var valuesSsidSeis = rangeSsidSeis.getValues();
  var headingsAp = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var ssidApIndex = headingsAp[0].indexOf("State Student ID");
  var rangeSsidAp = sheet.getRange(2, ssidApIndex + 1, sheet.getLastRow() - 1, 1);
  var valuesSsidAp = rangeSsidAp.getValues();
  var students = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var rangePercentOutSeis = sheetSeis.getRange(2, percentOutSeisIndex + 1, sheetSeis.getLastRow() - 1, 1);
  var valuesPercentOutSeis = rangePercentOutSeis.getValues();
  var array = [];
  for (var i = 0; i < valuesSsidAp.length; i++) {
    for (var j = 0; j < valuesSsidSeis.length; j++) {
      var allP = parseInt(valuesSsidAp[i][0].toString());
      var seis = valuesSsidSeis[j][0];
      if (allP == seis || students[i][12] == '21') {
        var percentOut = valuesPercentOutSeis[j][0] || 'NA';
        array.push([1, percentOut]);
        break;
      }
      else {
        if (j == valuesSsidSeis.length - 1) {
          array.push([0, ""]);
        }
      }
    }
  }
  // see if the dest column already exists
  var headings = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var a = headings[0].indexOf("dpaightStudents");
  if (a != -1) {
    var dpaightColumn = a + 1;
  }
  else {
    var dpaightColumn = headings[0].length + 1;
  }
  //  var dpaightColumn = a = -1 ? headings[0].length + 1 : a ;
  var dest = sheet.getRange(2, dpaightColumn, array.length, array[0].length);
  dest.setValues(array);
  var heading = sheet.getRange(1, dpaightColumn, 1, 1);
  heading.setValue("dpaightStudents");
  var heading = sheet.getRange(1, dpaightColumn + 1, 1, 1);
  heading.setValue("Percent OUT Regular Class");
}
//# sourceMappingURL=module.jsx.map