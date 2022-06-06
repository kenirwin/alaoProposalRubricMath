/*
Code for averaging (etc) scores from individual sheets, plus data validation and copying the master review sheet
Author: Ken Irwin, irwinkr@miamioh.edu
Date: 1 April 2020
Last Revised: 6 June 2022, kri
Save updates to Git Repo: https://github.com/kenirwin/alaoProposalRubricMath

Note: all sheets must be laid out the same. Everything relies on each cell (e.g. "R2") meaning the same thing and being averagable.
*/

// note: reviewerNames should not contain any spaces. Use camelCase if needed,e.g. KatyT,AnnMarie, etc
var reviewerNames = ['Allen','AnnMarie','Cara','Don','Jerry','KathyF','KatieM','KatyT','Ken','Mark','Melissa','Paul','Peggy','Sara','Seth','Rob'].sort();
// var reviewerNames = ['Ken','Cara','Mark'].sort();
// enforce no spaces in reviewerNames:
reviewerNames = reviewerNames.map(i => i.replaceAll(' ',''));

const sourceId = '15deVp7UdoC8EvaYpRhlCZCMpRGRwSc2rkrRfdbGhD5E';
const numSheetsBeforeReviewers = 4;
/* 
getSheets should know the name of every individual sheet with scores that "count"
*/

function getSheets() {
 return reviewerNames;
}

/*
getScoreArray returns an array with the same score (cell value) from each sheet mentioned in getSheets()
this array is used by the AveScore(), NumScores(), and StDevScores() functions
*/
function getScoreArray(cell) {
  var sheets = getSheets();
  var scores = [];
  sheets.forEach(name => {
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
   let row = cell.substring(1, 2);
   let rowCompleteCell = "V"+row;
   let rowComplete = sheet.getRange(rowCompleteCell).getValue(); 
   Logger.log(rowCompleteCell, rowComplete);
   value = sheet.getRange(cell).getValue(); /* if you get an error in this line, probably one of sheets in getSheets() doesn't exist */
  if (value !=0 && rowComplete) { scores.push(value); } //don't push empty scores
  });
  return scores;
}

function isRowComplete() {
  var input = "R6";
  var row = input.substring(1,2);
  var completeCell = "V"+row;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ken');
  var cell = sheet.getRange(input).getValue();
  var completeness = sheet.getRange(completeCell).getValue();
  Logger.log(cell, completeness);
}

function AveScore(cell, bogus) {
  var scores = getScoreArray(cell);
  return (average(scores)); //uses the average() function below 
}

function NumScores(cell, bogus) {
  var scores = getScoreArray(cell);
  return scores.length;
}

function StDevScores(cell, bogus) {
  var scores = getScoreArray(cell);
  return (standardDeviation(scores)); //uses the standardDeviation() function below
}


/* 
the following functions are from:
https://derickbailey.com/2014/09/21/calculating-standard-deviation-with-array-map-and-array-reduce-in-javascript/
*/

function standardDeviation(values){
  var avg = average(values);
  
  var squareDiffs = values.map(function(value){
    var diff = value - avg;
    var sqrDiff = diff * diff;
    return sqrDiff;
  });
  
  var avgSquareDiff = average(squareDiffs);

  var stdDev = Math.sqrt(avgSquareDiff);
  return stdDev;
}

/*
  The Faster Score Totals
*/
function setupFasterScores() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Faster Score Totals');
    let reviewerScores = reviewerNames.map(r => r+'!R2').join(',');
    let rowCompletion = reviewerNames.map(r => r+'!V2').join(','); // array of V2: row complete T/F
    let titleScores = reviewerNames.map(r => r+'!F2').join(',');
    let clarityScores = reviewerNames.map(r => r+'!H2').join(',');
    let filterForNonZeroFinishedScores = 'FILTER({'+reviewerScores+'},{'+reviewerScores+'}>0, {'+rowCompletion +'}=TRUE)';
    let d2Value = '=COUNT('+filterForNonZeroFinishedScores+')';    
    let e2Value = '=SUM('+filterForNonZeroFinishedScores+')/D2';
    let f2Value = '=MEDIAN('+filterForNonZeroFinishedScores+')';
    let g2Value = '=STDEVP('+filterForNonZeroFinishedScores+')';
    let h2Value = '=MIN('+filterForNonZeroFinishedScores+')';
    let i2Value = '=MAX('+filterForNonZeroFinishedScores+')';
    let j2Value = '=I2-H2';
    let k1Value = '=CONCATENATE("Average Spread ",AVERAGE(FILTER(J:J,IFNA({J:J}))))';
    let m2Value = '=AVERAGE(FILTER({'+titleScores+'},{'+titleScores+'}>0))';
    let n2Value = '=MEDIAN(FILTER({'+titleScores+'},{'+titleScores+'}>0))';
    let o2Value = '=MIN(FILTER({'+titleScores+'},{'+titleScores+'}>0))';
    let p2Value = '=AVERAGE(FILTER({'+clarityScores+'},{'+clarityScores+'}>0))';
    let q2Value = '=MEDIAN(FILTER({'+clarityScores+'},{'+clarityScores+'}>0))';
    let r2Value = '=MIN(FILTER({'+clarityScores+'},{'+clarityScores+'}>0))';
    sheet.getRange('D2').setValue(d2Value);
    sheet.getRange('E2').setValue(e2Value);
    sheet.getRange('F2').setValue(f2Value);
    sheet.getRange('G2').setValue(g2Value);
    sheet.getRange('H2').setValue(h2Value);
    sheet.getRange('I2').setValue(i2Value);
    sheet.getRange('J2').setValue(j2Value);
    sheet.getRange('K1').setValue(k1Value);
    sheet.getRange('L2').setValue('=getIncompletes()');
    sheet.getRange('M2').setValue(m2Value);
    sheet.getRange('N2').setValue(n2Value);
    sheet.getRange('O2').setValue(o2Value);
    sheet.getRange('P2').setValue(p2Value);
    sheet.getRange('Q2').setValue(q2Value);
    sheet.getRange('R2').setValue(r2Value);
    
    let fillDownRange = sheet.getRange('D3:J100');
    sheet.getRange('D2:J2').copyTo(fillDownRange);
  
    fillDownRange = sheet.getRange('M2:R2');
    sheet.getRange('M3:R100').copyTo(fillDownRange);
  }

function average(data){
  var sum = data.reduce(function(sum, value){
    return sum + value;
  }, 0);

  var avg = sum / data.length;
  return avg;
}

/* 
 recalculate
 * forces a redo/update of the math on the "Score Totals" sheet
 * also forces an update of the getIncompletes() list on the Score Total
 * works by adding then removing a column at the beginning of the sheet - it's a hack to get it to recalculate

function recalculate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Score Totals');
  
  // This inserts a column in the first column position
  sheet.insertColumnBefore(1);
  sheet.deleteColumn(1);
}
*/

function getIncompletes() {
  var sheets = getSheets();
  var incompletes = ["Incomplete data for:"];
  var complete = ["Score entry complete"];
  sheets.forEach(name => {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    value = sheet.getRange("W1").getValue();
    if (value.includes('FALSE')) {
      incompletes.push(name);
    } //push if not complete
  });
  if (incompletes.length > 1) {
    return incompletes;
  } else {
    return complete;
  }
}

/**************************** SETUP STUFF **********************/
function LogUsefulForumulas() {
  arr = [];
  reviewerNames.forEach(i => arr.push(i));
  Logger.log(arr);
  var ranges = arr.join("!R2,");
  ranges = "{" + ranges + "!R2}";
  Logger.log(ranges);
  Logger.log('=COUNTIF('+ranges+',">0")');
  Logger.log('=SUMIF('+ranges+',">0")/D2');   
  Logger.log('=STDEV(FILTER('+ranges+','+ranges+'>0))');
}


/**************************** DANGER ***************************/
/*
The stuff below here is dangerous. Don't mess with it and definitely don't run it unless you have read and understand the documentation.
*/

function makeReviewerSheets() {
  let insertAfter = numSheetsBeforeReviewers+1; 
  reviewerNames.forEach(name =>{
                        insertAfter++;
                        copyMasterReviewSheet(name, insertAfter);
                        });
  setupFasterScores();
}

function protectReviewerRanges(sheet_to_protect) {
  var ranges = ['A:E', 'G:G', 'I:I', 'K:K', 'M:M', 'O:O', 'Q:R', 'U:W'];
  ranges.forEach(range => {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(sheet_to_protect);
      var range = sheet.getRange(range);
      var protection = range.protect().setDescription('Sample protected range');

      // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
      // permission comes from a group, the script throws an exception upon removing the group.

      var me = Session.getEffectiveUser();
      protection.addEditor(me);
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }
    });
}


function testSheetInfo() {
  var source = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(source.getSheetName());
}

function copyMasterReviewSheet(newName, newPosition) {
  /* 
  NOTE: IF THIS BREAKS...
  probably it isn't broken, you are.
  I found that for reasons unknown to me, I could log into the spreadsheet as a privileged user (my work email), then when I open this script 
  I was logged in as my personal email without privileges - I was getting errors saying that I needed to select a sheet. But the real 
  problem was the script wasn't thinking of itself as being associated with the current active spreadsheet, so getActiveSpreadsheet failed 
  and did so quietly. Logging out of all identities and back under the right one fixed the problem.
  /*
  modeled on:
  https://stackoverflow.com/questions/19791132/google-apps-script-copy-one-spreadsheet-to-another-spreadsheet-with-formatting
  */

  Logger.log('Making sheet for reviewer: ',newName);
  var master = 'Data Entry Review Form';
  var source = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Active spreadsheet: '+ source);
  Logger.log('SheetName: ' +source.getSheetName());
  Logger.log(Object.getOwnPropertyNames(source));
  // var sourceId = '15deVp7UdoC8EvaYpRhlCZCMpRGRwSc2rkrRfdbGhD5E';
  Logger.log('SourceID: '+sourceId);
  var template = source.getSheetByName(master);
  template.copyTo(source).setName(newName);
  protectReviewerRanges(newName);
  source.setActiveSheet(source.getSheetByName(newName))
  source.moveActiveSheet(newPosition);
}

function DANGERdeleteReviewSheets() {
    reviewerNames.forEach(name =>{
                        DANGERdeleteSheet(name);
                        });
}
function DANGERdeleteSheet(name) {
  var source = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = source.getSheetByName(name);
  source.deleteSheet(sheet);
}
