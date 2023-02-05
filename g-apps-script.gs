/*
When the spreadsheet is opened, it will create a menu
on the top bar that will give the user the ability to run the functions
*/
function onOpen(){
  SpreadsheetApp.getUi()
  .createMenu('Sacrament Bulletin')
  .addItem('Generate Sacrament Bulletin For Upcoming Sunday', 'generateBulletin_')
  .addSeparator()
  .addItem('Email Complete Bulletin', 'emailBulletin_')
  .addToUi();
}


function generateBulletin_() {

  var ss = SpreadsheetApp.getActiveSpreadsheet()

  // get sheet data
  var mysheet = ss.getActiveSheet()
  var sheetData = mysheet.getDataRange().getValues();

  // current time
  var now = new Date()

  // Logger.log(mysheet)
  Logger.log(sheetData)
  Logger.log(now)


  for (var i=1; i<sheetData.length; i++){
      //
      // var files = DriveApp.getFilesByName(filenames[ff]);
      // Logger.log(sheetData[i])
      // Logger.log('date')
      // Logger.log(sheetData[i][0])
      // Logger.log(sheetData[i][0] > now)

      futureDate = sheetData[i][0] > now
      if (futureDate){
        var row = i
        break
      }

  }
  // Logger.log('the row we are going to use is:')
  // Logger.log(row)
  const friendlyDate = new Date(sheetData[row][0]).toLocaleDateString();
  var presiding = sheetData[row][1]
  var conducting = sheetData[row][2]
  var open_hymn = sheetData[row][3]
  var sacr_hymn = sheetData[row][4]
  var speaker1 = sheetData[row][5]
  var speaker2 = sheetData[row][6]
  var speaker3 = sheetData[row][7]
  var music_num = sheetData[row][8]
  var close_hymn = sheetData[row][9]
  var open_prayer = sheetData[row][10]
  var close_prayer = sheetData[row][11]

  var cong_hymn = 999
  var cong_name_sp = ''
  var cong_hymn_sp = ''
  var cong_name = ''

  const googleDocTemplate = DriveApp.getFileById('198Am0qVBtJfElctc75KWclH9Fogb5JWOA68RM2JkxcI');

  const hymn_ss = SpreadsheetApp.openById('1i0jvgIqMnMiuGEFl_BjJ9QUA6BMrS7T0DTI4CJNdT6w')
  hymndata = hymn_ss.getActiveSheet().getDataRange().getValues();
  for (var i=0; i<hymndata.length; i++){
    if (hymndata[i][0] == open_hymn){
      Logger.log(hymndata[i])
      var open_name = hymndata[i][1]
      var open_hymn_sp = hymndata[i][2]
      var open_name_sp = hymndata[i][3]
    }
    if (hymndata[i][0] == sacr_hymn){
      Logger.log(hymndata[i])
      var sacr_name = hymndata[i][1]
      var sacr_hymn_sp = hymndata[i][2]
      var sacr_name_sp = hymndata[i][3]
    }
    if (hymndata[i][0] == cong_hymn){
      Logger.log(hymndata[i])
      var cong_name = hymndata[i][1]
      var cong_hymn_sp = hymndata[i][2]
      var cong_name_sp = hymndata[i][3]
    }
    if (hymndata[i][0] == close_hymn){
      Logger.log(hymndata[i])
      var close_name = hymndata[i][1]
      var close_hymn_sp = hymndata[i][2]
      var close_name_sp = hymndata[i][3]
    }

  }

  //This value should be the id of the folder where you want your completed documents stored
  const destinationFolder = DriveApp.getFolderById('1dL2pfGkRQX-YPh7CyO2NkZrt7AQh2Vf6')

  //Start processing each spreadsheet row

  //Using the row data in a template literal, we make a copy of our template document in our destinationFolder
  const copy = googleDocTemplate.makeCopy(`${now} test bulletin` , destinationFolder)
  //Once we have the copy, we then open it using the DocumentApp
  const doc = DocumentApp.openById(copy.getId())
  //All of the content lives in the body, so we get that for editing
  const body = doc.getBody();

  //In these lines, we replace our replacement tokens with values from our spreadsheet row
  body.replaceText('{{date}}', friendlyDate);
  body.replaceText('{{presiding}}', presiding);
  body.replaceText('{{conducting}}', conducting);
  body.replaceText('{{open_hymn}}', open_hymn);
  body.replaceText('{{sacr_hymn}}', sacr_hymn);
  body.replaceText('{{speaker1}}', speaker1);
  body.replaceText('{{speaker2}}', speaker2);
  body.replaceText('{{speaker3}}', speaker3);
  body.replaceText('{{conducting}}', music_num);
  body.replaceText('{{close_hymn}}', close_hymn);
  body.replaceText('{{open_prayer}}', open_prayer);
  body.replaceText('{{close_prayer}}', close_prayer);

  body.replaceText('{{close_name}}', close_name);
  body.replaceText('{{open_name}}', open_name);
  body.replaceText('{{cong_name}}', cong_name);
  body.replaceText('{{sacr_name}}', sacr_name);

  body.replaceText('{{close_hymn_sp}}', close_hymn_sp);
  body.replaceText('{{open_hymn_sp}}', open_hymn_sp);
  body.replaceText('{{cong_hymn_sp}}', cong_hymn_sp);
  body.replaceText('{{sacr_hymn_sp}}', sacr_hymn_sp);

  body.replaceText('{{close_name_sp}}', close_name_sp);
  body.replaceText('{{open_name_sp}}', open_name_sp);
  body.replaceText('{{cong_name_sp}}', cong_name_sp);
  body.replaceText('{{sacr_name_sp}}', sacr_name_sp);

  //We make our changes permanent by saving and closing the document
  doc.saveAndClose();
  //Store the url of our new document in a variable
  const url = doc.getUrl();
  //Write that value back to the 'Document Link' column in the spreadsheet.
  // sheet.getRange(index + 1, 6).setValue(url)
  mysheet.getRange(row+1,13).setValue(url)
  Logger.log(url)



}

function emailBuletin_(){



}
