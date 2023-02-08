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

  const ss = SpreadsheetApp.getActiveSpreadsheet()
  
  // get sheet data
  const mysheet = ss.getActiveSheet()
  const sheetData = mysheet.getDataRange().getValues();

  // current time
  const now = new Date()

  // Logger.log(mysheet)
  Logger.log(sheetData)
  Logger.log(now)

  // find the next sunday in the data 
  for (let i=1; i<sheetData.length; i++){

      futureDate = sheetData[i][0] > now
      Logger.log(sheetData[i][0])
      Logger.log(futureDate)
      if (futureDate){
        var row = i
        break
      }
  }


  // Logger.log('the row we are going to use is:')
  // Logger.log(row)
  const friendlyDate = new Date(sheetData[row][0]).toLocaleDateString();
  const fast_sunday = sheetData[row][1]
  const  presiding = sheetData[row][2]
  const conducting = sheetData[row][3]
  const organist = sheetData[row][4]
  const chorister = sheetData[row][5]
  const open_hymn = sheetData[row][6]
  const sacr_hymn = sheetData[row][7]
  const cong_hymn = sheetData[row][8]
  const close_hymn = sheetData[row][9]
  const open_prayer = sheetData[row][10]
  const close_prayer = sheetData[row][11]
  const speaker1 = sheetData[row][12]
  const speaker2 = sheetData[row][13]
  const speaker3 = sheetData[row][14]
  
  
  
  let cong_name_sp = ''
  let cong_hymn_sp = ''
  let cong_name = ''

  // const googleDocTemplate = DriveApp.getFileById('198Am0qVBtJfElctc75KWclH9Fogb5JWOA68RM2JkxcI');
  const googleDocTemplate = DriveApp.getFileById('1W_l9ck5cnN7LWdAvRZQOh5GP4AF0MOwABNmneA1_-dw');
  

  const hymn_ss = SpreadsheetApp.openById('1i0jvgIqMnMiuGEFl_BjJ9QUA6BMrS7T0DTI4CJNdT6w')
  hymndata = hymn_ss.getActiveSheet().getDataRange().getValues();
  for (let i=0; i<hymndata.length; i++){
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
      cong_name = hymndata[i][1]
      cong_hymn_sp = hymndata[i][2]
      cong_name_sp = hymndata[i][3]
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
  const copy = googleDocTemplate.makeCopy(`${now} Bulletin for ${friendlyDate}` , destinationFolder)
  //Once we have the copy, we then open it using the DocumentApp
  const doc = DocumentApp.openById(copy.getId())
  //All of the content lives in the body, so we get that for editing
  const body = doc.getBody();

  // this does work for fast sunday but it removes all the text fonts and formatting..
  // it might be better to do a seperate template for fast sundays.
  if (fast_sunday){
    fast_match_string = body.getText().match(/##Y[\s\S]*}}##/m);
    body.setText(body.getText().replace(fast_match_string[0],'Bearing of Testimonies'))
  }
  else {
    body.replaceText('##','')
  }


  //In these lines, we replace our replacement tokens with values from our spreadsheet row
  body.replaceText('{{date}}', friendlyDate);

  body.replaceText('{{presiding}}', presiding);
  body.replaceText('{{conducting}}', conducting);

  body.replaceText('{{organist}}', organist);
  body.replaceText('{{chorister}}', chorister);

  body.replaceText('{{open_hymn}}', open_hymn);
  body.replaceText('{{sacr_hymn}}', sacr_hymn);
  body.replaceText('{{cong_hymn}}', cong_hymn);
  body.replaceText('{{close_hymn}}', close_hymn);

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

  body.replaceText('{{open_prayer}}', open_prayer);
  body.replaceText('{{close_prayer}}', close_prayer);

  body.replaceText('{{speaker1}}', speaker1);
  body.replaceText('{{speaker2}}', speaker2);
  body.replaceText('{{speaker3}}', speaker3);
  

  //We make our changes permanent by saving and closing the document
  doc.saveAndClose();
  //Store the url of our new document
  const url = doc.getUrl();
  //Write that value back to the 'Document Link' column in the spreadsheet. 
  mysheet.getRange(row+1,16).setValue(url)

  
}

function testing(){

  // generateBulletin_()
  emailBulletin_()

}


function emailBulletin_(){

  const ss = SpreadsheetApp.getActiveSpreadsheet()

  var ui = SpreadsheetApp.getUi();
  button_result = ui.alert('Save PDF and Email?','Do you want to create a PDF of the upcoming bulletin and email to interested parties?',ui.ButtonSet.YES_NO)
  if (button_result == ui.Button.NO){
    ui.alert('Canceled','PDF generation cancelled',ui.ButtonSet.OK)
    return
  }
  // get sheet data
  const mysheet = ss.getActiveSheet()
  const sheetData = mysheet.getDataRange().getValues();

  // current time
  const now = new Date()

  // find the next sunday in the data 
  for (let i=1; i<sheetData.length; i++){

      futureDate = sheetData[i][0] > now
      if (futureDate){
        var row = i
        break
      }
  }

  const url = mysheet.getRange(row+1,16).getValue()
  const fileId = url.slice(32)

  const bulletinFile = DriveApp.getFileById(fileId)
  Logger.log(bulletinFile)

  // 

  docblob = bulletinFile.getAs('application/pdf');
  /* Add the PDF extension */
  docblob.setName(bulletinFile.getName() + ".pdf");
  var file = DriveApp.createFile(docblob);
  file.moveTo(DriveApp.getFolderById('1dL2pfGkRQX-YPh7CyO2NkZrt7AQh2Vf6'))

  Logger.log('Your PDF file is available at ' + file.getUrl());

  sendEmailAddress = ['jaredeoliphant@gmail.com','joliphant@brightwaterfoundation.org']
  for (let i=0; i<sendEmailAddress.length; i++){
    MailApp.sendEmail(sendEmailAddress[i],'Bulletin','See attached bulletin.',{attachments: [file.getAs(MimeType.PDF)]})
  }
}
