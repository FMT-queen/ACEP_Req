// Deploy value
const getDeployURL = 'https://script.google.com/a/macros/fujielectric.com/s/AKfycbx0uEXVXLriin5NXWck3aNvXj4jX8unl3iy5pXOFjC_5BtjscuWigN5tjxIIRqo3LoE/exec'

//initial database sheet
const databaseID = "1SDLGjb5Kf54h75vyJZlOfDW5YIHvVo-9p4gHIdnPxb0";
const databaseSheetName = "Data_RunDoc";
const databaseRunningDoc = "Doc.Num_Data";

//initial input sheet
const sheetFormName = "Form_ACEP"
const sheetMasterName = "Master"
const formRunningDocName = "IMP_Doc.Num"
const formPosted = "Form_Post";

//initial email
var sheetSection = "IT"
var emailApproveName = "Approve Test User*"; //Approval section name
var emailApprove = "k.suphattra@fujielectric.com" //Approval section email  
var emailCCsec = "u.amorntep@fujielectric.com"
var emailAcountName = "accounting Test User*" //Accounting name
var emailAccounting = "k.suphattra@fujielectric.com" //Accounting email


//Declare range
var rangeDocNo = 'C3';
var rangeAdvance = 'D5';
var rangeClear = 'D6';
var rangeExpense = 'D7';
var rangeClearTo = 'E9';
var rangeEmpCode = 'J5';
var rangeSection = 'J7:K7';
var rangeReqName = 'N7:P7';
var rangeVenderCode = 'J9';
var rangeVenderName = 'N9:P9';
var rangePaymentDate = 'E13';
var rangePaymentMet = 'J13:K13';
var rangeCurrency = 'N13:P13';
var rangeAttach = 'E16:L16'
var rangeDetail = 'D21:P40';
var rangeTotal = 'P44';
var rangeSeachDoc = 'G52:K52';
var rangeSecAppName = 'D57:F57';
var rangeSecComment = 'D60:F62';
var rangeSecStatus = 'E65';
var rangeAccoAppName = 'K57:M57';
var rangeAccoPayDate = 'P57';
var rangeAccoComment = 'K61:N62';
var rangeAccoStatus = 'P62';
var rangeAccoClearStatus = 'K65:M65';
var currentDateTime = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy, HH:mm").toString();
var currentDate = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy").toString();
var rangeCancel ='O3:Q4';
var rangeRefAttach = 'E17:L17'


//------------------------------------------------------------------------------------------
//var ui = SpreadsheetApp.getUi();

function doGet(e) {
  const getParameter = e.parameter['docNo'];
  //Logger.log(getParameter);
  
  const reqDoc = getParameter.slice(0,11);
  const typeButton = getParameter.slice(12,19)
  
  var getHtml;
  // Approve part.
  if(typeButton == 'Approve'){
    const logAppStatus = approveReq(reqDoc);
      Logger.log("Test")
      if(logAppStatus == true) {
        //Logger.log("logAppStatus = true");
       getHtml = 'AppSuccess.html'
      };
      if(logAppStatus == false) {
        //Logger.log("logAppStatus = false");
      getHtml = 'AppFailed.html'
      };
  }
  // Reject part.
   if(typeButton == 'Reject'){
    getHtml = 'Reject.html'
  }
  // const popUp = HtmlService.createHtmlOutputFromFile(getHtml);   
  const html = HtmlService.createTemplateFromFile(getHtml);
    html.value = reqDoc;
    var output = html.evaluate()
    return output;
}

function rejectFun(x,y) {
    var getDoc = x;
    var getComment = y;
     const logRejectStatus = rejectReq(getDoc, getComment);
      Logger.log(logRejectStatus)
      if(logRejectStatus == true) {
       return "Reject successfully";
      };
      if(logRejectStatus == false) {
        return "Reject failed. Please, re-check document status or you have no authorization.";
      };
}

function onOpen(e) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input
  var docNumData = spreadsheet.getSheetByName(formRunningDocName); // Sheet IMP_Doc.Num
  //Browser.msgBox('ระบบกำลังประมวลผล..');
  docNumData.getRange(1,1,docNumData.getMaxRows(),docNumData.getMaxColumns()).clearContent();
  spreadsheet.getRange(rangeAdvance).clearContent();
  spreadsheet.getRange(rangeClear).clearContent();
  spreadsheet.getRange(rangeExpense).clearContent();
  spreadsheet.getRange(rangeDocNo).clearContent();
 
  //resetButt();
  spreadsheet.getRange('E5').setValue("Advance"); //set value advance
  spreadsheet.getRange('E6').setValue("Clear"); //set value clear
  spreadsheet.getRange('E7').setValue("Expense"); //set value Expense
  sheetInput.getRange(rangeSection).setValue(sheetSection);

  
   var activeUser = e.user.getEmail();
   var result = temLockSheet(activeUser);
 
  //sheetInput.getRange('L5').setValue(activeUser);
  var result = temLockSheet(activeUser);
  if(result != activeUser){
    Browser.msgBox('Another user ('+ result+ ')is currently using.')
    //ui.alert('Sheet temporarylock by '+ result)
  }
}

function onClose(e){
  cancelLockSheet();

}

function onEdit(e){
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input
  var sheetImpDoc = spreadsheet.getSheetByName(formRunningDocName); //Sheet IMP_Doc.Num
  var sheetMaster = spreadsheet.getSheetByName(sheetMasterName); //Sheet Master
  var sheetPosted = spreadsheet.getSheetByName(formPosted); //Sheet post
  // Entry
  sheetInput.getRange("D5:E7").setBackground('#ffe599');
  sheetInput.getRange(rangeEmpCode).setBackground('#f3f3f3');
  sheetInput.getRange(rangeVenderName).setBackground('#f3f3f3');
  //sheetInput.getRange(rangeReqName).setBackground('#f3f3f3');
  sheetInput.getRange(rangePaymentDate).setBackground('#f3f3f3');
  sheetInput.getRange(rangePaymentMet).setBackground('#f3f3f3');
  sheetInput.getRange(rangeCurrency).setBackground('#f3f3f3');
  //sheetInput.getRange(rangeAttach).setBackground('#f3f3f3');
  sheetInput.getRange("D21:P21").setBackground('#f3f3f3');
  sheetInput.getRange("E9").setBackground('#ffffff');

  //----auto append email accounting
  if(e.range.getColumn() == 16 && e.range.getRow() == 57){
    var appendEmail = Session.getActiveUser().getEmail();
    sheetInput.getRange(rangeAccoAppName).setValue(appendEmail);
  }
  if(10<=e.range.getColumn()<= 12 && e.range.getRow() == 65){
    var appendEmail = Session.getActiveUser().getEmail();
    sheetInput.getRange(rangeAccoAppName).setValue(appendEmail);
  }
   if(10<= e.range.getColumn() <= 13 && e.range.getRow()== 61){ //Reject
    var appendEmail = Session.getActiveUser().getEmail();
    sheetInput.getRange(rangeAccoAppName).setValue(appendEmail);
    //Logger.log("Loop 3")
  }
  //This IF statement ensures that this onEdit macro only runs when cells J5 (Emp. Code) are edited
  if ( e.source.getSheetName() == sheetFormName && e.range.getColumn() == 10 && e.range.getRow() == 5) {
    //Logger.log("Find Emp. Name Loop")
      var inputEmpCode = spreadsheet.getRange(rangeEmpCode).getValue();
      //Logger.log("Emp.code input =" +  inputEmpCode)
      //find Emp code
    var empName = sheetMaster.getDataRange().getValues();
    var empNameValue;
    var empNameFound = false; 
    for (var i = 0; i < empName.length; i++) {
    var rowValue = empName[i];
    //Logger.log(rowValue)
        if (rowValue[0] == inputEmpCode && rowValue[2] == sheetSection){
        var findRow = i + 1;
        //Browser.msgBox(findRow);
        empNameValue = sheetMaster.getRange("B" + findRow).getValue();
        empNameFound = true;
      } sheetInput.getRange(rangeReqName).setValue(empNameValue)
    } 
    if (empNameFound == false) {
      sheetInput.getRange(rangeReqName).clearContent();
    ui.alert("Error", "Emp.Code ไม่ถูกต้อง", ui.ButtonSet.OK);
    return false;
    }
  }
 //This IF statement ensures that this onEdit macro only runs when cells N9:P9 are edited
   if ( e.source.getSheetName() == sheetFormName && e.range.getColumn() == 14 && e.range.getRow() == 9) {
    //Logger.log("Find Vendor Loop")
      var inputVender = spreadsheet.getRange(rangeVenderName).getValue();
      //Logger.log("Vendor Code =" +  inputVender)
      //find Vendor Code
    var verdorCode = sheetMaster.getDataRange().getValues(); //verdorCode
    var verdorCodeValue;
    var verdorCodeFound = false; 
    for (var i = 0; i < verdorCode.length; i++) {
    var rowValue = verdorCode[i];
    //Logger.log(rowValue)
        if (rowValue[7] == inputVender){
        var findRow = i + 1;
        //Browser.msgBox(findRow);
        verdorCodeValue = sheetMaster.getRange("G" + findRow).getValue();
        verdorCodeFound = true;
      } sheetInput.getRange(rangeVenderCode).setValue(verdorCodeValue)
    } 
    if (verdorCodeFound == false) {
    ui.alert("Error", "Vender Name ไม่ถูกต้อง", ui.ButtonSet.OK);
    return false;
    }
  }  


  //This IF statement ensures that this onEdit macro only runs when cells A1:A2 are edited
  if (
    e.source.getSheetName() == sheetFormName &&
    e.range.getColumn() == 4 &&
    e.range.getRow()> 4 &&
    e.range.getRow()< 8
  ) { //Logger.log("Check box edited");
    Logger.log("test")
    var range = e.range;
    var sheet = range.getSheet();
    

      var chkAdvance = sheetInput.getRange(rangeAdvance).getValue();
      var chkClear = sheetInput.getRange(rangeClear).getValue();
      var chkExpense = sheetInput.getRange(rangeExpense).getValue();
      var typeDoc = 0 ;
        //Browser.msgBox('ระบบกำลัง Generate Doc.No');
      if(chkAdvance == 'A'&& e.range.getColumn() == 4 && e.range.getRow()==5){ //e.range.getColumn() == 3 && e.range.getRow()<=5
        spreadsheet.getRange(rangeClear).clearContent();
        spreadsheet.getRange(rangeExpense).clearContent(); 
        typeDoc = chkAdvance;
        spreadsheet.getRange(rangeClearTo).clearContent()
        spreadsheet.getRange(rangeClearTo).setBackground(null)
        //For ref Attach
        sheetInput.getRange("C17:D17").clearContent();
        sheetInput.getRange(rangeRefAttach).clearContent();
        sheetInput.getRange("C17:D17").setBackground('#ffffff');
        sheetInput.getRange(rangeAttach).setBackground('#f3f3f3');
        sheetInput.getRange(rangeRefAttach).setBackground('#ffffff');
        sheetInput.getRange("C16:D16").setBackground('#ffe599');
        sheetInput.getRange("D16").setValue("Attached file :") 
      }
      if(chkClear == 'C'&& e.range.getColumn() == 4 && e.range.getRow()==6){ //e.range.getColumn() == 3 && e.range.getRow()<=6
        spreadsheet.getRange(rangeAdvance).clearContent();
        spreadsheet.getRange(rangeExpense).clearContent();
        typeDoc = chkClear;
        ui.alert("Information", "กรุณากรอกเลขที่ Advance ที่ต้องการ Clear", ui.ButtonSet.OK);
        spreadsheet.getRange(rangeClearTo).setBackground('#d3e3fd');
        spreadsheet.getRange(rangeClearTo).clearContent();
        //For ref Attach
        sheetInput.getRange("C16:D16").setBackground('#ffffff');
        sheetInput.getRange(rangeAttach).setBackground('#ffffff');
        sheetInput.getRange("D16").setValue("Advance Ref.:") 
        sheetInput.getRange("C17:D17").setBackground('#ffe599');
        sheetInput.getRange("D17").setValue("Attached file :")
        sheetInput.getRange(rangeRefAttach).setBackground('#f3f3f3');
      }
      if(chkExpense == 'E'&& e.range.getColumn() == 4 && e.range.getRow()==7){ //e.range.getColumn() == 3 && e.range.getRow()<=7
        spreadsheet.getRange(rangeAdvance).clearContent();
        spreadsheet.getRange(rangeClear).clearContent();
        typeDoc = chkExpense;
        spreadsheet.getRange(rangeClearTo).clearContent()
        spreadsheet.getRange(rangeClearTo).setBackground(null)
        //For ref Attach
        sheetInput.getRange("C17:D17").clearContent();
        sheetInput.getRange(rangeRefAttach).clearContent();
        sheetInput.getRange("C17:D17").setBackground('#ffffff');
        sheetInput.getRange(rangeAttach).setBackground('#f3f3f3');
        sheetInput.getRange("C16:D16").setBackground('#ffe599');
        sheetInput.getRange(rangeRefAttach).setBackground('#ffffff');
        sheetInput.getRange("D16").setValue("Attached file :") 
      } 
      
      var getDocLocal = generateDocNum_Local(typeDoc);
      
      //Logger.log("getDocLocal =" + getDocLocal);
      
    
      sheetInput.getRange(rangeDocNo).setValue(getDocLocal)
    
    return(0);
  }

  //Total amount edited
  if (
    e.source.getSheetName() == sheetFormName &&
    e.range.getColumn() == 16 &&
    21<=e.range.getRow()<=40
  ) { //Logger.log("Amount edited");
      sheetInput.getRange(rangeTotal).setValue('=SUM(P21:P40)')
  }

}

function myOnEdit(e){
    //sheet Posted find doc
  if ( e.source.getSheetName() == sheetPosted && e.range.getColumn() == 1 && e.range.getRow() == 2) {
    var sheetPosted = spreadsheet.getSheetByName(formPosted); //Sheet post
    var docPost = sheetPosted.getRange("C3").getValue();
    Logger.log("searchButt()")
    searchButt(docPost)
  }
}




function generateDocNum_Local(typeDoc){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var docNumData = spreadsheet.getSheetByName(formRunningDocName); // Sheet IMP_Doc.Num
  docNumData.getRange(1,1,docNumData.getMaxRows(),docNumData.getMaxColumns()).clearContent()
  docNumData.getRange("A1").setFormula(
    '=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1SDLGjb5Kf54h75vyJZlOfDW5YIHvVo-9p4gHIdnPxb0/edit#gid=0","Doc.Num_Data")'
  ) 
  const values = docNumData.getDataRange().getValues();
  docNumData.getDataRange().setValues(values);
  
//-------------------------------------------------------------------------------------------------------------
  var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input

  var chkAdvance = sheetInput.getRange(rangeAdvance).getValue();
  var chkClear = sheetInput.getRange(rangeClear).getValue();
  var chkExpense = sheetInput.getRange(rangeExpense).getValue();
    Logger.log("Check box Advance = " + chkAdvance)
    Logger.log("Check box Clear = " + chkClear)
    Logger.log("Check box Expense = " + chkExpense)

  var Alast = 0;
  var Clast = 0;
  var Elast = 0;
  var lastRow = 0;
 //test get last row of column+
 if (typeDoc == "A"){
    var Avals = docNumData.getRange("A1:A").getValues();
    Alast = Avals.filter(String).length;
    lastRow = Alast;
  } //var data = ss.getDataRange().getDisplayValues();
 if (typeDoc == "C"){
   var Cvals = docNumData.getRange("C1:C").getValues();
   Clast = Cvals.filter(String).length;
   lastRow = Clast;
  } 
  if (typeDoc == "E"){
    var Evals = docNumData.getRange("E1:E").getValues();
    Elast = Evals.filter(String).length;
    lastRow = Elast;
  }
    Logger.log("find a last = "+ Alast)
    Logger.log("find c last = "+ Clast)
    Logger.log("find e last = "+ Elast)

  // var lastRow = docNumData.getLastRow();
  Logger.log("last Row = " + lastRow);
  var valueLastRow = docNumData.getRange(typeDoc+lastRow).getValue();
  var getreserveDoc = sheetInput.getRange(rangeDocNo).getValue().slice(7);
  
  var typeReserveDoc = sheetInput.getRange(rangeDocNo).getValue().slice(0,1);
  Logger.log("typeReserveDoc = " + typeReserveDoc);
  Logger.log("valueLastRow = " + valueLastRow)
//----------------------------------------------------------------------------- Check Resevre doc
  if(valueLastRow == "A" || valueLastRow == "C" || valueLastRow == "E"){
    reserveDoc = 0;
    valueLastRow = 0;
  } else var reserveDoc = parseInt(getreserveDoc);
  Logger.log("reserveDoc = " + reserveDoc)
  
  if(typeDoc == typeReserveDoc){
  while(valueLastRow > reserveDoc || valueLastRow == reserveDoc){
      reserveDoc = reserveDoc+1;
    } 
  } else reserveDoc = valueLastRow+1;
  docNumData.getRange(typeDoc+(parseInt(lastRow)+1)).setValue(reserveDoc); //  write last doc
  Logger.log(reserveDoc);
    //Rundoc
  var yearDoc = Utilities.formatDate(new Date(), "GMT+7", "yyyy").toString();
  var monthDoc = Utilities.formatDate(new Date(), "GMT+7", "MM").toString();
   var docNum = reserveDoc;
   //check digit
   if (docNum < 10 ){ // 0-9
    var charDocNum = "000" + docNum.toString();
    } else if (docNum < 100) { // 00 - 99
        var charDocNum = "00" + docNum.toString();
      } else if (docNum < 1000) { // 000 - 999 
        var charDocNum = "0" + docNum.toString();
  }else charDocNum = docNum.toString(); //  > 1000
  Logger.log(charDocNum)
  var docNo = typeDoc + yearDoc + monthDoc + charDocNum;
  sheetInput.getRange(rangeDocNo).clearContent();

  return(docNo);
}

function getDocType(typeDoc){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input
  var typeDoc = "";
  var chkAdvance = sheetInput.getRange(rangeAdvance).getValue();
  var chkClear = sheetInput.getRange(rangeClear).getValue();
  var chkExpense = sheetInput.getRange(rangeExpense).getValue();
  if(chkAdvance == 'A' && chkClear == '' && chkExpense == ''){
    typeDoc = chkAdvance;
  } else if (chkAdvance == '' && chkClear == 'C' && chkExpense == ''){
      typeDoc = chkClear;
    } else if (chkAdvance == '' && chkClear == '' && chkExpense == 'E'){
      typeDoc = chkExpense;
      }  
  return typeDoc;
}

function fileAttachment() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('uploadForm.html')
        .setWidth(500) //200
        .setHeight(200); //50
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Upload PDF');
    }
    function uploadFileToDrive(base64Data, fileName) {
      var ui = SpreadsheetApp.getUi();
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input
      var decodedData = Utilities.base64Decode(base64Data);
      //fileBlob.getContentType()
      var getDocNo = sheetInput.getRange(rangeDocNo).getValue();
      var getTypeDoc = getDocNo.slice(0,1)
      var folderName = '';
    //if a newly ***Change folder name and folder ID *** 
    // if(getDocNo != '') {
    //   folderName = getDocNo; // ชื่อโฟลเดอร์ที่ต้องการใช้งาน
    // } else { 
    //    ui.alert("Error", 'กรุณาเลือกประเภทเอกสาร', ui.ButtonSet.OK); 
    //    return 0;
    //   }
      folderName = 'TempFileAttach'
    try{ DriveApp.getFolderById("1NI3WI00CTM85xf-tFHxVMKN5dZYmrdgz")
    } catch(e) {
      ui.alert("Error", 'You do not have permission to access the requested folder.', ui.ButtonSet.OK); 
      return 0;
    }
    var rootFolder = DriveApp.getFolderById("1NI3WI00CTM85xf-tFHxVMKN5dZYmrdgz");
    var folder = DriveApp.getFoldersByName(folderName);

      var targetFolder;
      if (folder.hasNext()) {
       targetFolder = folder.next();
      } else targetFolder = rootFolder.createFolder(folderName);
      
      var blob = Utilities.newBlob(decodedData).setContentType('application/pdf').setName(fileName);
      //Logger.log("blob.getContentType()" + blob.getContentType())
      var file = targetFolder.createFile(blob);
      //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var linkCell = "";
      if(getTypeDoc == 'A' || getTypeDoc == 'E'){
        linkCell = sheetInput.getRange(rangeAttach);
        linkCell.setValue(file.getUrl());
        sheetInput.getRange(rangeAttach).setBackground('#f3f3f3');
      }
      if(getTypeDoc == 'C'){
        linkCell = sheetInput.getRange(rangeRefAttach);
        linkCell.setValue(file.getUrl());
        sheetInput.getRange(rangeRefAttach).setBackground('#f3f3f3');
      } ui.alert("Complete", "อัปโหลดไฟล์ PDF เสร็จสิ้น", ui.ButtonSet.OK);
      return file.getUrl();
    }

//Save and Running Doc.-----------------------------------------------------------------------------------------------------------------------------


function generateDocNum(docNo){
  var datasheet = SpreadsheetApp.openById(databaseID); //Sheet database
  var docNumData = datasheet.getSheetByName(databaseRunningDoc); // Sheet Doc.Num_Data
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //var docNumData = datasheet.getSheetByName(databaseRunningDoc);
  var sheetMain = datasheet.getSheetByName(databaseSheetName); // Sheet Data_RunDoc (main)
  var sheetTest = SpreadsheetApp.openById("1pznwMgNJDTxetrYCUIu-Kf8Cach4Ulg1Xu2MSyUXaMo"); //Sheet test
  var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input

  var typeDoc = getDocType(); 
  var chkAdvance = sheetInput.getRange(rangeAdvance).getValue();
  var chkClear = sheetInput.getRange(rangeClear).getValue();
  var chkExpense = sheetInput.getRange(rangeExpense).getValue();
    Logger.log("Check box Advance = " + chkAdvance)
    Logger.log("Check box Clear = " + chkClear)
    Logger.log("Check box Expense = " + chkExpense)

  var Alast = 0;
  var Clast = 0;
  var Elast = 0;
  var lastRow = 0;
 //test get last row of column+
 if (typeDoc == "A"){
    var Avals = docNumData.getRange("A1:A").getValues();
    Alast = Avals.filter(String).length;
    lastRow = Alast;
  } //var data = ss.getDataRange().getDisplayValues();
 if (typeDoc == "C"){
   var Cvals = docNumData.getRange("C1:C").getValues();
   Clast = Cvals.filter(String).length;
   lastRow = Clast;
  } 
  if (typeDoc == "E"){
    var Evals = docNumData.getRange("E1:E").getValues();
    Elast = Evals.filter(String).length;
    lastRow = Elast;
  }
    // Logger.log("find a last = "+ Alast)
    // Logger.log("find c last = "+ Clast)
    // Logger.log("find e last = "+ Elast)

  // var lastRow = docNumData.getLastRow();
  
  var valueLastRow = docNumData.getRange(typeDoc+lastRow).getValue();
  var getreserveDoc = sheetInput.getRange(rangeDocNo).getValue().slice(7);
  // Logger.log("getreserveDoc = " + getreserveDoc)
  var typeReserveDoc = sheetInput.getRange(rangeDocNo).getValue().slice(0,1);
  var reserveDoc = parseInt(getreserveDoc);
  // Logger.log("typeReserveDoc = " + typeReserveDoc);
  // Logger.log("valueLastRow = " + valueLastRow)
  // Logger.log("reserveDoc = " + reserveDoc)
//----------------------------------------------------------------------------- Check Resevre doc
  if(valueLastRow == "A" || valueLastRow == "C" || valueLastRow == "E"){
    reserveDoc = 0;
    valueLastRow = 0;
  } else reserveDoc = parseInt(getreserveDoc);

if(typeDoc == typeReserveDoc){
  while(valueLastRow > reserveDoc || valueLastRow == reserveDoc){
      reserveDoc = reserveDoc+1;
    } 
  } else reserveDoc = valueLastRow+1;
  docNumData.getRange(typeDoc+(parseInt(lastRow)+1)).setValue(reserveDoc); //  write last doc
  // Logger.log(reserveDoc);
 var docNum = reserveDoc;
  
  //Rundoc
  var yearDoc = Utilities.formatDate(new Date(), "GMT+7", "yyyy").toString();
  var monthDoc = Utilities.formatDate(new Date(), "GMT+7", "MM").toString();
  if (docNum < 10 ){ // 0-9
    var charDocNum = "000" + docNum.toString();
    } else if (docNum < 100) { // 00 - 99
        var charDocNum = "00" + docNum.toString();
      } else if (docNum < 1000) { // 000 - 999 
        var charDocNum = "0" + docNum.toString();
  }else charDocNum = docNum.toString(); //  > 1000
  // Logger.log(charDocNum)
  var docNo = typeDoc + yearDoc + monthDoc + charDocNum;
  
  
  return(docNo);

}

//sheetInput
function postButt(){
var ui = SpreadsheetApp.getUi();
  try{ SpreadsheetApp.openById(databaseID); // check permission to sheet database
  } catch(e) {
    ui.alert("Error", "You do not have permission to access the requested function.", ui.ButtonSet.OK);
    return 0;
  }
  
  var datasheet = SpreadsheetApp.openById(databaseID); //Sheet database
  var docNumData = datasheet.getSheetByName(databaseRunningDoc); // Sheet Doc.Num_Data
  var sheetMain = datasheet.getSheetByName(databaseSheetName); // Sheet Data_RunDoc (main)
  var sheetTest = SpreadsheetApp.openById("1pznwMgNJDTxetrYCUIu-Kf8Cach4Ulg1Xu2MSyUXaMo"); //Sheet test
  var sheetInput = sheetTest.getSheetByName(sheetFormName); //Sheet input
  
  //Check data yellow fileds
    if (sheetInput.getRange(rangeDocNo).isBlank() == true) {
      ui.alert("Error", "กรุณาเลือกประเภทเอกสาร", ui.ButtonSet.OK);
      sheetInput.getRange("D5:E7").activate();
      sheetInput.getRange("D5:E7").setBackground('#FF0000');
      return false;
      }
    if(sheetInput.getRange(rangeEmpCode).isBlank() == true || sheetInput.getRange(rangeReqName).getValue() == ""){
      ui.alert("Error", "กรุณาใส่รหัสประจำตัวหรือชื่อผู้กรอกเอกสาร", ui.ButtonSet.OK);
      sheetInput.getRange(rangeEmpCode).activate();
      sheetInput.getRange(rangeEmpCode).setBackground('#FF0000');
      sheetInput.getRange(rangeReqName).activate();
      sheetInput.getRange(rangeReqName).setBackground('#FF0000');
      return false;
      }
    if (sheetInput.getRange(rangeVenderName).isBlank() == true) {
      ui.alert("Error", "กรุณาใส่ชื่อ Vendor", ui.ButtonSet.OK);
      sheetInput.getRange(rangeVenderName).activate();
      sheetInput.getRange(rangeVenderName).setBackground('#FF0000');
      return false;
      }
     if (sheetInput.getRange(rangePaymentDate).isBlank() == true) {
      ui.alert("Error", "กรุณาใส่วันที่ต้องการจ่ายเงิน ", ui.ButtonSet.OK);
      sheetInput.getRange(rangePaymentDate).activate();
      sheetInput.getRange(rangePaymentDate).setBackground('#FF0000');
      return false;
      }
     if (sheetInput.getRange(rangePaymentMet).isBlank() == true) {
      ui.alert("Error", "กรุณาเลือกวิธีการจ่ายเงิน", ui.ButtonSet.OK);
      sheetInput.getRange(rangePaymentMet).activate();
      sheetInput.getRange(rangePaymentMet).setBackground('#FF0000');
      return false;
      }
    if (sheetInput.getRange(rangeCurrency).isBlank() == true) {
      ui.alert("Error", "กรุณาเลือกสกุลเงิน", ui.ButtonSet.OK);
      sheetInput.getRange(rangeCurrency).activate();
      sheetInput.getRange(rangeCurrency).setBackground('#FF0000');
      return false;
      }
    if (sheetInput.getRange(rangeAttach).isBlank() == true) {
      ui.alert("Error", "กรุณาแนบไฟล์เอกสาร Advance นามสกุล (.pdf) เท่านั้น", ui.ButtonSet.OK);
      sheetInput.getRange(rangeAttach).activate();
      sheetInput.getRange(rangeAttach).setBackground('#FF0000');
      return false;
      }
    if (sheetInput.getRange("D21:P21").isBlank() == true) {
      ui.alert("Error", "กรุณากรอกรายละเอียด อย่างน้อย 1รายการ", ui.ButtonSet.OK);
      sheetInput.getRange("D21:P21").activate();
      sheetInput.getRange("D21:P21").setBackground('#FF0000');
      return false;
      }
    if (sheetInput.getRange(rangeRefAttach).isBlank() == true && sheetInput.getRange('D6').getValue() == 'C') {
      ui.alert("Error", "กรุณาแนบไฟล์เอกสาร Clear นามสกุล (.pdf) เท่านั้น", ui.ButtonSet.OK);
      sheetInput.getRange(rangeRefAttach).activate();
      sheetInput.getRange(rangeRefAttach).setBackground('#FF0000');
      return false;
    }
    if (sheetInput.getRange(rangeClearTo).isBlank() == true && sheetInput.getRange('D6').getValue() == 'C') {
      ui.alert("Error", "กรุณากรอกเลขที่ Advance ที่ต้องการ Clear", ui.ButtonSet.OK);
      sheetInput.getRange(rangeClearTo).activate();
      sheetInput.getRange(rangeClearTo).setBackground('#FF0000');
      return false;
    }

  var docNo = generateDocNum();
  var saveStatus = "Requester posted";
  var appStatus = "Waiting"

  const jSonStored = new Object();
  const arrDetail = [];
  
  for (var i = 0; i< 4; i++ ){
    jSonStored.Discription = sheetInput.getRange('D' + (21+i).toString()).getValue(); //:F21 '"'+ 'C' + (21+i).toString() + '"'
    jSonStored.AcountCode = sheetInput.getRange('I' + (21+i).toString()).getValue(); //:H21 
    jSonStored.UnitCode = sheetInput.getRange('L' + (21+i).toString()).getValue(); //:J21
    jSonStored.Amount = sheetInput.getRange('P' + (21+i).toString()).getValue(); //:M21
    if(jSonStored.Discription != "" || jSonStored.AcountCode != "" || jSonStored.UnitCode != "" || jSonStored.Amount != "") {
        arrDetail.push(JSON.stringify(jSonStored)); //.Discription,jSonStored.AcountCode,jSonStored.UnitCode,jSonStored.Amount
    }
  }
  //Logger.log(arrDetail)

  //var detail = sheetInput.getRange("rangeDetail").getValues();
  //Logger.log("Detail =" + JSON.stringify(jSonStored));
  var typeDocmail = " ";
  var attachmentFile = "";
  var refClearDoc = '' ;
  var folderTargetID = '';
  if (docNo.slice(0,1) == 'A'){
    typeDocmail = "advance payment";
    attachmentFile = rangeAttach;
    folderTargetID = '1_XBjqJwKAicpLR2k8oNHBhOWampkdEmY';
  } else if (docNo.slice(0,1) == 'C') {
    typeDocmail = "clear payment";
    attachmentFile = rangeRefAttach;
    refClearDoc = sheetInput.getRange(rangeClearTo).getValue(); 
    //folderTargetID = '1a6gv0Ze93K8mEgcKX6uHv4xr5HaLJ5e5';
  } else if (docNo.slice(0,1) == 'E'){
    typeDocmail = "expense payment";
    attachmentFile = rangeAttach;
    folderTargetID = '1kuKvCTO7TstoVT9jTQ9CP6TRLTKX6_R8';
  }
  //SetValue >> save value to sheet database
  var reqName = sheetInput.getRange(rangeReqName).getValue();
  var reqCurrecy = sheetInput.getRange(rangeCurrency).getValue();
  var reqVender = sheetInput.getRange(rangeVenderName).getValue();
  var reqAttach = sheetInput.getRange(attachmentFile).getValue();
  var reqTotal = sheetInput.getRange(rangeTotal).getValue();
  var values = [[
    saveStatus, //Status
    currentDateTime, //date
    docNo, //running doc C3
    sheetInput.getRange(rangeEmpCode).getValue(), //Emp code
    sheetInput.getRange(rangeSection).getValue(), //Section
    reqName, //Requester name
    sheetInput.getRange(rangeVenderCode).getValue(), //Vender Code
    reqVender, //Vender name
    sheetInput.getRange(rangePaymentDate).getValue(), //Payment date
    sheetInput.getRange(rangePaymentMet).getValue(), //Payment methode
    reqCurrecy, //Currency
    reqAttach, //Attached file
    JSON.stringify(arrDetail), //Detail
    reqTotal //Total
    , , 
    appStatus, 
    , ,
    appStatus // accoStatus
    , , , ,
    refClearDoc, //Clear to
    // sheetInput.getRange("D51").getValue(), //Aprove 1
    // sheetInput.getRange("J57").getValue(), //Comment 1
    // sheetInput.getRange("M57").getValue(), //Status 1
    // sheetInput.getRange("D61").getValue(), //Aprove 2
    // sheetInput.getRange("H61").getValue(), //Comment 2
    // sheetInput.getRange("M61").getValue(), //Status 2
    // "waitting", //Account confirm
    // sheetInput.getRange("C16").getValue(), //Account user
    // sheetInput.getRange("C16").getValue(), //Payment Date

  ]]
//------finding Attach file move to tartet folder---------------------------------------------------------------------------------------

var getFile = reqAttach.split("/")
var getFileID = getFile[5];
  try {DriveApp.getFileById(getFileID);
  } catch(e) {
    ui.alert("Error", "กรุณาแนบไฟล์ Reference", ui.ButtonSet.OK);
      return 0;
    }
var files = DriveApp.getFileById(getFileID);
var folderName = docNo // สำหรับ Advance และ Expense
//---find and get advance folder-สำหรับเอกสาร Clear-----------------
  var advancefolders = DriveApp.getFoldersByName(refClearDoc);
    if(docNo.slice(0,1) == 'C'){
      if (advancefolders.hasNext()){
        var advancefolderID = advancefolders.next().getId();
        folderTargetID = advancefolderID
      } else {
         ui.alert("Error", "ไม่พบโฟลเดอร์เก็บ Advance เบอร์ "+ refClearDoc, ui.ButtonSet.OK);
         return 0;
      }
    } 

var folder = DriveApp.getFoldersByName(folderName);
var targetFolder;
var rootFolder;
 rootFolder = DriveApp.getFolderById(folderTargetID);
  if (folder.hasNext()) {
       targetFolder = folder.next();
    } else targetFolder = rootFolder.createFolder(folderName);
files.moveTo(targetFolder);


//----------------------------------------------------------------------------------------------
const headers = ["No","Discription", "AcountCode", "UnitCode", "Amount"];
var table = "<html><head> <style> table, td, th { border: 1px solid;} table {width: 50%; border-collapse: collapse;}</style></head><body><table><tr><th>No</th><th>Discription</th><th>AcountCode</th><th>UnitCode</th><th>Amount</tr>"; //border-collapse: collapse
//</br>
//
//the body of the table is build in 2D (two foor loops)
//Logger.log("arrDetail.lngth"+arrDetail.length)

//const jsonForMail = JSON.parse(arrDetail);
    for(var j = 0; j<arrDetail.length;j++){ 
      const jsonValues = arrDetail[j];
      const jsonItem = JSON.parse(jsonValues);
      table = table + "<tr>";
        for(var u = 0; u < 1; u++){  //Column 
           table =  table + (j+1).toString()+ "<td>" + jsonItem.Discription+"</td>" + jsonItem.AcountCode+"</td>" + jsonItem.UnitCode+"</td><td align=\"right\">" + jsonItem.Amount+"</td>"  //;
           //Logger.log("jsonItem.Discription = "+jsonItem.Discription)
        }
      table = table + "</tr>"
    } 
   table = table+"</table></body></html>";
//----button style----------------------------------------------------------------------------------
  const button ="<html><head>" +
                "</head><body><h2>Approval part</h2>" + 
                "<button><a href=" + '"' + getDeployURL +'?docNo='+docNo +';Approve'+ '"' + ">Approve</a></button>" + 
                "<span>&#8199;</span>" + 
                "<button><a href=" + '"' + getDeployURL +'?docNo='+docNo +';Reject'+ '"' + ">Reject</a></button>" + 
                "</body></html>"

   //---------------------------------------------------------------------------------------

  var shortCurrency = reqCurrecy.slice(0,3);

    //Header
  var mailHeader = "Requestision: " + typeDocmail+ "- " + docNo + " | " + currentDate ;

  //button
   //var yesButton = '<button style = "background-color:#4CAF50;" id='yesButton' onclick='myfunction(True)'>'

    //mail body
  var bodyHeader ="Dear, " + emailApproveName +'<br/><br/>' + "This is an auto-generated notification from Advance Clear/Expense Payment Requesition Form, please do not reply.<br/><br/>";
  var bodyData = reqName + " was created a " + typeDocmail +" document as a detail below.<br/>"
  var mailBody = bodyHeader + bodyData + "<strong>Document No. : </strong>"+ docNo + '<br/>' + "<strong>Requester name : </strong>" + reqName + '<br/>' + "<strong>Pay to : </strong>" +reqVender + '<br/>' +"<strong>Details</strong>" + table  + '<br/>' +"<strong>Total : </strong>"+ reqTotal + " " +shortCurrency+ '<br/>' +"<strong>Attached file : </strong>"+ reqAttach + '<br/>'+ button; //+ buttonReject

  //Send the email:
 MailApp.sendEmail({
    to: emailApprove, 
    cc: emailCCsec,
    subject: mailHeader,
    htmlBody: mailBody }); 
//---------------------------------------------------------------------------------------------
  sheetMain.getRange(sheetMain.getLastRow()+ 1, 1, 1, 23).setValues(values); // 3= Total column that insert
Browser.msgBox('Document ' + '"'+ docNo +'"'+' is Posted. ' );

//Clear range
resetButt();
//sheetInput.getRange(rangeSection).setValue('IT'); // Section


cancelLockSheet();




} // End postButt

function resetButt(){
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input

sheetInput.getRange(rangeDocNo).clearContent();
sheetInput.getRange(rangeAdvance).clearContent();
sheetInput.getRange(rangeClear).clearContent();
sheetInput.getRange(rangeExpense).clearContent();
spreadsheet.getRange(rangeClearTo).setBackground(null);
sheetInput.getRange(rangeClearTo).clearContent();
sheetInput.getRange(rangeEmpCode).clearContent();
//sheetInput.getRange(rangeSection).clearContent();
sheetInput.getRange(rangeReqName).clearContent();
sheetInput.getRange(rangeVenderCode).clearContent();
sheetInput.getRange(rangeVenderName).clearContent();
sheetInput.getRange(rangePaymentDate).clearContent();
sheetInput.getRange(rangePaymentMet).clearContent();
sheetInput.getRange(rangeCurrency).clearContent();
sheetInput.getRange(rangeAttach).clearContent();
sheetInput.getRange(rangeDetail).clearContent();
sheetInput.getRange(rangeTotal).clearContent();
sheetInput.getRange(rangeSeachDoc).clearContent();
sheetInput.getRange(rangeSecAppName).clearContent();
sheetInput.getRange(rangeSecComment).clearContent();
sheetInput.getRange(rangeSecStatus).clearContent();
sheetInput.getRange(rangeAccoAppName).clearContent();
sheetInput.getRange(rangeAccoPayDate).clearContent();
sheetInput.getRange(rangeAccoComment).clearContent();
sheetInput.getRange(rangeAccoStatus).clearContent();
sheetInput.getRange(rangeAccoClearStatus).clearContent();
sheetInput.getRange(rangeCancel).clearContent();

sheetInput.getRange("D5:E7").setBackground('#ffe599');
sheetInput.getRange(rangeEmpCode).setBackground('#f3f3f3');
sheetInput.getRange(rangeVenderName).setBackground('#f3f3f3');
sheetInput.getRange(rangePaymentDate).setBackground('#f3f3f3');
sheetInput.getRange(rangePaymentMet).setBackground('#f3f3f3');
sheetInput.getRange(rangeCurrency).setBackground('#f3f3f3');
sheetInput.getRange(rangeAttach).setBackground('#f3f3f3');
sheetInput.getRange("D21:P21").setBackground('#f3f3f3');
sheetInput.getRange("E9").setBackground('#ffffff');
sheetInput.getRange("C16:D16").setBackground('#ffe599');


// Ref Attach 
sheetInput.getRange("D17").clearContent();
sheetInput.getRange(rangeRefAttach).clearContent();
sheetInput.getRange("C17:D17").setBackground('#ffffff');
sheetInput.getRange(rangeRefAttach).setBackground('#ffffff');

cancelLockSheet();

} // End resetButt

function trigApproveMail() {
Logger.log("Approve trig")

}

function cancelButt(){
  var ui = SpreadsheetApp.getUi();
    try{ SpreadsheetApp.openById(databaseID); // check permission to sheet database
    } catch(e) {
    ui.alert("Error", "You do not have permission to access the requested function.", ui.ButtonSet.OK);
    return 0;
    }

  var sheetActive = SpreadsheetApp.openById(databaseID); //Sheet Database
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input
  var datasheet = sheetActive.getSheetByName(databaseSheetName); //Sheet Database
  var cancelDocF = sheetInput.getRange(rangeDocNo).getValue(); //get doc no.
  var cellDocNo = sheetInput.getRange(rangeDocNo);
  var appendEmail = Session.getActiveUser().getEmail();
  var strUser = appendEmail.split('@', true)

   if (cellDocNo.isBlank() == true ) { // if ไม่ใส่ค่า 
    ui.alert("Please enter value to Cancel")
    resetButt();
    return false;
  }
  var values = datasheet.getDataRange().getValues();
  var valueFound = false;
  // Logger.log("cancelDocF" + cancelDocF);
  // Logger.log("values.length" + values.length);

  for (var i = 0; i < values.length; i++) {
    var rowValue = values[i];
    //Logger.log(rowValue[2]);
    
    if (rowValue[2] == cancelDocF) {
      if(rowValue[0] == 'Requester posted'){
       var foundRow = (i+1).toString();
       datasheet.getRange('A'+foundRow).setValue("Cancel by requester"); 
       datasheet.getRange('B'+foundRow).setValue(currentDateTime);
       datasheet.getRange('V'+foundRow).setValue(strUser);
       datasheet.getRange('S'+foundRow).setValue('-');
       datasheet.getRange('P'+foundRow).setValue('-');

       ui.alert("Cancel complete", "เอกสาร "+ cancelDocF + " ถูก Cancel เรียบร้อยแล้ว" , ui.ButtonSet.OK);
       //Logger.log("i =" + i);
      resetButt();
       valueFound = true
       return;
      } else { ui.alert("Cancel failed", "กรุณาตรวจสอบสถานะเอกสาร "+ cancelDocF, ui.ButtonSet.OK);
        valueFound = true
        return;

      }
     
    } 
  }
    if (valueFound == false) {
      ui.alert("Error", "ไม่พบเอกสาร "+ cancelDocF + " ในระบบ" , ui.ButtonSet.OK);
     //ui.alert("ไม่พบข้อมูลในระบบ");
     return;
    } 
  
}


function searchButt(docPost) {
  
  var ui = SpreadsheetApp.getUi();
  var sheetActive = SpreadsheetApp.openById(databaseID); //Sheet Database
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input
  var datasheet = sheetActive.getSheetByName(databaseSheetName); //Sheet Database

  Logger.log("Doc to find" + docNoF)
  var values = datasheet.getDataRange().getValues();
  var valueFound = false;
  var docNoF = ""
  if(docPost == ""){
    docNoF = sheetInput.getRange(rangeSeachDoc).getValue(); //get doc no.
  } else  docNoF = docPost;
  


  if (sheetInput.getRange(rangeSeachDoc).isBlank() == true ) { // if ไม่ใส่ค่า 
    ui.alert("Please enter value to find")
    return false;
  } resetButt();
  //get type doc
  var sTypeDoc = docNoF.slice(0,1); // slice(before start,start End -->)
  //Logger.log("sTypeDoc"  + sTypeDoc);
  sheetInput.getRange(rangeSeachDoc).setValue(docNoF);
  resetButt();
  for (var i = 0; i < values.length; i++) {
    var rowValue = values[i];
    //Logger.log(rowValue[2]);
    if (rowValue[2] == docNoF) { // && rowValue[1] == lineNo
      //show data
      sheetInput.getRange(rangeDocNo).setValue(rowValue[2]); //DocNo
       if (sTypeDoc == 'A') {
        sheetInput.getRange(rangeAdvance).setValue(sTypeDoc);
        sheetInput.getRange(rangeAttach).setValue(rowValue[11]); //Attach file
       }
       if (sTypeDoc == 'C') {
        sheetInput.getRange(rangeClear).setValue(sTypeDoc);
        sheetInput.getRange("C16:D16").setBackground('#ffffff');
        sheetInput.getRange(rangeAttach).setBackground('#ffffff');
        sheetInput.getRange("C17:D17").setBackground('#ffe599');
        sheetInput.getRange("D17").setValue("Reference file :")
        sheetInput.getRange(rangeRefAttach).setBackground('#f3f3f3');
        sheetInput.getRange(rangeRefAttach).setValue(rowValue[11]); //Attach file
        //Logger.log("rowValue[22] = " + rowValue[22]);
        for(var k = 2; k <= i ; k++){
          var refValues = datasheet.getRange('C' + k.toString()).getValue();
          //Logger.log('refValues' + k + '=' + refValues);
          if(rowValue[22] == refValues){
            var attachFound = datasheet.getRange('L' + k.toString()).getValue();
            sheetInput.getRange(rangeAttach).setValue(attachFound); //Attach file
          } else valueFound = false;
        }

       }
       if (sTypeDoc == 'E'){
        sheetInput.getRange(rangeExpense).setValue(sTypeDoc);
        sheetInput.getRange(rangeAttach).setValue(rowValue[11]); //Attach file
       } 
      sheetInput.getRange(rangeEmpCode).setValue(rowValue[3]); //EmpCode
      //sheetInput.getRange(rangeSection).setValue(rowValue[4]); //EmpSection
      sheetInput.getRange(rangeReqName).setValue(rowValue[5]); //Emp Req Name
      sheetInput.getRange(rangeVenderCode).setValue(rowValue[6]); //VenderCode
      sheetInput.getRange(rangeVenderName).setValue(rowValue[7]); //VenderName
      sheetInput.getRange(rangePaymentDate).setValue(rowValue[8]); //PaymentDate
      sheetInput.getRange(rangePaymentMet).setValue(rowValue[9]); //PaymentMet
      sheetInput.getRange(rangeCurrency).setValue(rowValue[10]); //Currency
      
      
      if(rowValue[0] == "Cancel by requester"){
        sheetInput.getRange(rangeCancel).setValue(rowValue[0]); //Attach file
        }
      
       var jsonDetail = JSON.parse(rowValue[12]);
        for(var j = 0; j<jsonDetail.length;j++){ 
          //const headers = ["Discription", "AcountCode", "UnitCode", "Amount"];
          const jsonValues = jsonDetail[j];
          const jsonItem = JSON.parse(jsonValues);
          sheetInput.getRange('D'+ (21+j)).setValue(jsonItem.Discription);
          sheetInput.getRange('I'+ (21+j)).setValue(jsonItem.AcountCode);
          sheetInput.getRange('L'+ (21+j)).setValue(jsonItem.UnitCode);
          sheetInput.getRange('P'+ (21+j)).setValue(jsonItem.Amount);
        } 
      //sheetInput.getRange(rangeDetail).setValue()); //Detail
      //Logger.log(JSON.parse(rowValue[12]).length);
      sheetInput.getRange(rangeTotal).setValue(rowValue[13]); //Attach file
      sheetInput.getRange(rangeSecAppName).setValue(rowValue[14]); //Attach file
      sheetInput.getRange(rangeSecStatus).setValue(rowValue[15]); //Attach file
      sheetInput.getRange(rangeSecComment).setValue(rowValue[16]); //Attach file
      sheetInput.getRange(rangeAccoAppName).setValue(rowValue[17]); //Accounting Approve Name
      sheetInput.getRange(rangeAccoStatus).setValue(rowValue[18]); //Accounting Status
      sheetInput.getRange(rangeAccoComment).setValue(rowValue[19]); //Accounting Comment
      sheetInput.getRange(rangeAccoPayDate).setValue(rowValue[20]); //Accounting PayDate
      sheetInput.getRange(rangeClearTo).setValue(rowValue[22]); //Accounting PayDate
       valueFound = true
       return;
      //Logger.log("Test" + i)
    }
  }
  if (valueFound == false) {
    ui.alert("ไม่พบข้อมูลในระบบ");
    //submitClear();
  } 
}

function rejectReq(inputDocNo,getComment){
  //Logger.log("From function reject"+ inputDocNo)

  // var inputDocNo = x;
  // var getComment = y;
 
  var sheetActive = SpreadsheetApp.openById(databaseID); //Sheet Database
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input
  var datasheet = sheetActive.getSheetByName(databaseSheetName); //Sheet Database
  //var cancelDocF = sheetInput.getRange(rangeDocNo).getValue(); //get doc no.
  //var cellDocNo = ;
  var appendEmail = Session.getActiveUser().getEmail();
  //var strUser = appendEmail.split('@', true)
  Logger.log('appendEmail')

   if (inputDocNo == "" || appendEmail != emailApprove) { // if ไม่ใส่ค่า 
    return false;
  }
  var values = datasheet.getDataRange().getValues();
  var valueFound = false;

  for (var i = 0; i < values.length; i++) {
    var rowValue = values[i];
    if (rowValue[2] == inputDocNo && rowValue[0] == 'Requester posted') {
      var foundRow = (i+1).toString();
       datasheet.getRange('A'+foundRow).setValue("Section rejected"); 
       datasheet.getRange('B'+foundRow).setValue(currentDateTime);
       datasheet.getRange('O'+foundRow).setValue(appendEmail);
       datasheet.getRange('Q'+foundRow).setValue(getComment);
       datasheet.getRange('P'+foundRow).setValue("Rejected");
       valueFound = true
       return true;
    }
  }
    if (valueFound == false) {
     return false;
  } 
}

function approveReq(inputDocNo){
  Logger.log("From function approve"+ inputDocNo)
 
  var sheetActive = SpreadsheetApp.openById(databaseID); //Sheet Database
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input
  var datasheet = sheetActive.getSheetByName(databaseSheetName); //Sheet Database
  //var cancelDocF = sheetInput.getRange(rangeDocNo).getValue(); //get doc no.
  //var cellDocNo = ;
  var appendEmail = Session.getActiveUser().getEmail();
  //var strUser = appendEmail.split('@', true)
  Logger.log(appendEmail) 

   if (inputDocNo == "" || appendEmail != emailApprove) { // if ไม่ใส่ค่า 
    return false;
  }
  var values = datasheet.getDataRange().getValues();
  var valueFound = false;

  for (var i = 0; i < values.length; i++) {
    var rowValue = values[i];
    if (rowValue[2] == inputDocNo && rowValue[0] == 'Requester posted') {
      var foundRow = (i+1).toString();
       datasheet.getRange('A'+foundRow).setValue("Section approved"); 
       datasheet.getRange('B'+foundRow).setValue(currentDateTime);
       datasheet.getRange('O'+foundRow).setValue(appendEmail);
       datasheet.getRange('P'+foundRow).setValue("Approved");
       valueFound = true

       var typeDocmail = " ";
         if (inputDocNo.slice(0,1) == 'A'){
          typeDocmail = "advance payment";
          } else if (inputDocNo.slice(0,1) == 'C') {
            typeDocmail = "clear payment";
          } else if (inputDocNo.slice(0,1) == 'E'){
            typeDocmail = "expense payment";
          }

      var mailHeader = "Approval: " + typeDocmail+ "- " + inputDocNo + " | " + currentDate ;
      var bodyHeader ="Dear, " + emailAcountName +'<br/><br/>' + "This is an auto-generated notification from Advance Clear/Expense Payment Requesition Form, please do not reply.<br/><br/>";
      var mailBody = bodyHeader + "Document No."+'<strong>' + inputDocNo + '</strong>' + " was approved. Please, check the details with the form below." +'<br/>' + "https://docs.google.com/spreadsheets/d/1pznwMgNJDTxetrYCUIu-Kf8Cach4Ulg1Xu2MSyUXaMo/edit?usp=sharing"

        MailApp.sendEmail({
          to: emailAccounting,
          cc: emailCCsec, 
          subject: mailHeader,
          htmlBody: mailBody 
          }); 

       return true;
    }
  }
    if (valueFound == false) {
     return false;
  } 
}

function accountAppButts(){
    var ui = SpreadsheetApp.getUi();
    var sheetActive = SpreadsheetApp.openById(databaseID); //Sheet Database
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input
    var datasheet = sheetActive.getSheetByName(databaseSheetName); //Sheet Database

    var appendEmail = Session.getActiveUser().getEmail();
    
    var strUser = appendEmail.split('@', true)
    var accountPayDate = sheetInput.getRange(rangeAccoPayDate).getValue();
    var accountComment = sheetInput.getRange(rangeAccoComment).getValue();
    var inputDocNo = sheetInput.getRange(rangeDocNo).getValue();
    var typeDocCheck = inputDocNo.slice(0,1)
    var docClear = sheetInput.getRange(rangeClearTo).getValue();
    var statusClearInput = sheetInput.getRange(rangeAccoClearStatus).getValue();
    var ref = "";
 
   if (inputDocNo == "" ) { // if ไม่ใส่ค่า 
    ui.alert("Error", "กรุณาเลือก Document No.", ui.ButtonSet.OK);
    return false;
    }
    if (appendEmail != emailAccounting) { // if ไม่ใส่ค่า 
    ui.alert("Error", "Authorize error", ui.ButtonSet.OK);
    return false;
    }
    if (accountPayDate == ''&& typeDocCheck != 'C') { // if ไม่ใส่ค่า 
    ui.alert("Error", "กรุณาเลือก Payment date", ui.ButtonSet.OK);
    return false;
    }
    var foundRow = ""
  var values = datasheet.getDataRange().getValues();
  var valueFound = false;
  for (var i = 0; i < values.length; i++) {
    var rowValue = values[i];
    //ของเอกสาร Clear ไป Finding หาคู่ Advance
    if(typeDocCheck == 'C' && rowValue[2] == docClear){
      if(rowValue[0] != 'Accounting approved'){
        ui.alert("Error", "Advance ref." + docClear + "ยังไม่ถูก Approve", ui.ButtonSet.OK);
        return 0;
      }
      //Logger.log("test")
      foundRow = (i+1).toString();
      ref = datasheet.getRange('w'+foundRow).getValue();
        if(statusClearInput == ''){
          statusClearInput = "Clear by accounting";
        } 
        if (rowValue[0] == "Remaining"){
          var valueRef = ref +", " +inputDocNo
          datasheet.getRange('w'+foundRow).setValue(valueRef);
          datasheet.getRange('A'+foundRow).setValue(statusClearInput); 
          datasheet.getRange('B'+foundRow).setValue(currentDateTime);
          datasheet.getRange('T'+foundRow).setValue(accountComment); 
        }
        if (rowValue[0] == "Accounting approved"){
        datasheet.getRange('w'+foundRow).setValue(inputDocNo);
        datasheet.getRange('A'+foundRow).setValue(statusClearInput); 
        datasheet.getRange('B'+foundRow).setValue(currentDateTime);
        datasheet.getRange('T'+foundRow).setValue(accountComment); 
        }
    } 
    //ทุกเอกสาร
    if (rowValue[2] == inputDocNo && rowValue[0] == 'Section approved' && appendEmail == emailAccounting) {
      foundRow = (i+1).toString();
       datasheet.getRange('A'+foundRow).setValue("Accounting approved"); 
       datasheet.getRange('B'+foundRow).setValue(currentDateTime);
       datasheet.getRange('R'+foundRow).setValue(appendEmail); 
       datasheet.getRange('S'+foundRow).setValue("Approved");
       datasheet.getRange('T'+foundRow).setValue(accountComment);
       datasheet.getRange('U'+foundRow).setValue(accountPayDate);
    ui.alert("Success", "Document "+ inputDocNo +" approve successfully", ui.ButtonSet.OK);
       valueFound = true
       resetButt();
       return true;
    }
  }
    if (valueFound == false) {
       ui.alert("Error", "Approve failed", ui.ButtonSet.OK);
     return false;
  } 
}

function accountRejectButts(){
    var ui = SpreadsheetApp.getUi();
    var sheetActive = SpreadsheetApp.openById(databaseID); //Sheet Database
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input
    var datasheet = sheetActive.getSheetByName(databaseSheetName); //Sheet Database

    var appendEmail = Session.getActiveUser().getEmail();
    
    var strUser = appendEmail.split('@', true)
    var accountPayDate = sheetInput.getRange(rangeAccoPayDate).getValue();
    var accountComment = sheetInput.getRange(rangeAccoComment).getValue();
    var inputDocNo = sheetInput.getRange(rangeDocNo).getValue();
   
   if (inputDocNo == "" ) { // if ไม่ใส่ค่า 
    ui.alert("Error", "กรุณาเลือก Document No.", ui.ButtonSet.OK);
    return false;
    }
    if (appendEmail != emailAccounting) {
    ui.alert("Error", "Authorize error", ui.ButtonSet.OK);
    return false;
    }
    if (accountComment == '') { // if ไม่ใส่ค่า 
    ui.alert("Error", "กรุณาใส่ Comment", ui.ButtonSet.OK);
    return false;
    }
   
  var values = datasheet.getDataRange().getValues();
  var valueFound = false;
  
  for (var i = 0; i < values.length; i++) {
    var rowValue = values[i];
   
    if (rowValue[2] == inputDocNo && rowValue[0] == 'Section approved' && appendEmail == emailAccounting) {
      var foundRow = (i+1).toString();
       datasheet.getRange('A'+foundRow).setValue("Accounting rejected"); 
       datasheet.getRange('B'+foundRow).setValue(currentDateTime);
       datasheet.getRange('R'+foundRow).setValue(appendEmail); 
       datasheet.getRange('S'+foundRow).setValue("Rejected");
       datasheet.getRange('T'+foundRow).setValue(accountComment);
      //  datasheet.getRange('U'+foundRow).setValue(accountPayDate);
    ui.alert("Success", "Document "+ inputDocNo +" rejected successfully", ui.ButtonSet.OK);
       valueFound = true
       resetButt();
       return true;
    }
  }
    if (valueFound == false) {
       ui.alert("Error", "Reject failed", ui.ButtonSet.OK);
     return false;
  } 
}

function temLockSheet(activeUser){
  // var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input
  var sheetProtect = SpreadsheetApp.getActiveSheet();
  
  // Protect the active sheet, then remove all other users from the list of editors.
  //firstUser = sheet.protect().getDescription();
  var firstUser;
  try { firstUser = sheetProtect.protect().getDescription(); //sheetInput.getRange('L5').getValues()
    } catch {firstUser = sheetProtect.protect().getDescription();}
 
  if(firstUser == ''){ //activeUser == firstUser
      firstUser = activeUser;
      var protection = sheetProtect.protect().setDescription(firstUser); // set protection
      var me = Session.getEffectiveUser();
      // permission comes from a group, the script throws an exception upon removing the group.
      // | protection.getEditors() --> output array all editer

      protection.addEditor(me);
      protection.removeEditors(protection.getEditors());
      firstUser = sheet.protect().getDescription();
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    } 
     Logger.log("Des. protect = " + firstUser)
  } 
  return firstUser;
}


function cancelLockSheet(){
  // var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // var sheetInput = spreadsheet.getSheetByName(sheetFormName); //Sheet input
  var sheetProtect = SpreadsheetApp.getActiveSheet();

      sheetProtect.protect().remove();

}









