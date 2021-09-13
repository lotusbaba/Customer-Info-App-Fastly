function myFunction() {

var myDog = "{\"name\": \"Rhino\", \"breed\": \"pug\", \"age\": 8}";
var myDogObj = JSON.parse(myDog);
var myJSON = JSON.stringify(myDog);
  Logger.log(myDogObj['name']);
  //abcObj = JSON.parse(abc);
  //console.log(abc.value['results']);
  //Logger.log(abc.value['results']);
}

function getRightApp()
{
  var uiTypeSpreadsheet = null;
  var uiTypeDocument = null;
  var uiTypeSlides = null;
  var retVal = {};
  
  /* One of the following will succeed */
  
  try {
    uiTypeSpreadsheet = SpreadsheetApp.getUi(); // Function to time.
    Logger.log("This is the spreadsheet getUI output");
    Logger.log(uiTypeSpreadsheet);
    retVal["thisApp"] = SpreadsheetApp;
    retVal["thisAppString"] = "SpreadsheetApp";
  } catch (err) {
    // Logs an ERROR message.
    console.error('SpreadsheetApp.getUi() yielded an error: ' + err + " with UI type " + uiTypeSpreadsheet);
  }

  /*try {
    uiTypeDocument = DocumentApp.getUi(); // Function to time.
    Logger.log("This is the document getUI output");
    Logger.log(uiTypeDocument);
    retVal["thisApp"]  = DocumentApp;
    retVal["thisAppString"] = "DocumentApp";
  } catch (err) {
    // Logs an ERROR message.
    console.error('DocumentApp.getUi() yielded an error: ' + err + " with UI type " + uiTypeDocument);
  }
  
  try {
    uiTypeSlides = SlidesApp.getUi(); // Function to time.
    Logger.log("This is the document getUI output");
    Logger.log(uiTypeSlides);
    retVal["thisApp"]  = SlidesApp;
    retVal["thisAppString"] = "SlidesApp";
  } catch (err) {
    // Logs an ERROR message.
    console.error('SlidesApp.getUi() yielded an error: ' + err + " with UI type " + uiTypeSlides);
  }*/
  return retVal;
}

var workingSheetColumnNames = ["Column Titles", "Customer Name", "Customer ID", "Service Name", "Service ID", "CHR", "Hits", "Misses", "Edge Response Header Bytes", "Edge Response Body Bytes", "Origin Fetch Response Header Bytes", "Origin Fetch Response Body Bytes", "Passes", "Video traffic (Response count)", "OTFP Traffic (Response count)", "WAF Requests Logged", "WAF Requests Blocked", "Bandwidth total in GB", "Image Opto Traffic", "Requests", "Edge Requests", "Origin Fetches",  "Uses Shielding", "Uses Image Opto", "Uses GZIP", "Uses WAF", "Uses Logging", "Uses Deliver Stale", "Uses Auto Director", "Uses Segmented Caching", "Uses OTFP", "Snippets", "Streaming miss", "ACLs", "Edge dictionary", "Synthetic response", "ESI response", "Crypto functions", "Uses Tokens/Signatures", "Uses Geo location", "Uses Device Detection", "Uses BBR"];

var workingSheetColumnValues = ["Values", "-", "-", "-", "-", "hit_ratio", "hits", "miss", "edge_resp_header_bytes", "edge_resp_body_bytes", "origin_fetch_resp_header_bytes", "origin_fetch_resp_body_bytes", "pass", "video", "otfp", "waf_logged", "waf_blocked", "bandwidth", "imgopto", "requests", "edge_requests", "origin_fetches", "shield_[s]*[s]*[l]*[_]*cache", "set req.http.[xX]-[fF]astly-[iI]mageopto-[aA]pi = ", "set beresp.gzip = true", "waf", "log *{\\\\\"", "deliver_stale", "autodirector_", "set req.enable_segmented_caching", "vpop", "[Ss]nippet", "do_stream", "acl", "table", "[Ss]ynthetic", "[Ee][Ss][Ii]", "crypto", "digest", "client.geo", "client.((?!geo)identified|display|browser|bot|class|platform|client|os)", "set client.socket.congestion_algorithm = \"bbr\""];

var vclRowInfo = 23;
var chrRowInfo = 6;

var otherInfoSheetCells = ["Title", "Is there Custom VCL?", "What kind of certs are is the customer using, and how are they managed?", "Does the customer use the Fastly UI or APIs", "Which user is making most configuration changes on customer side?", "How often does this user make changes?", "Uses dedicated IP", "Origin Peering", "Subscriber Prefixes"];

var custId = null;
var CUSTOMER_NAME_COLUMN = 1;
var CUSTOMER_ID_COLUMN = 2;
var SERVICE_NAME_COLUMN = 3;
var SERVICE_ID_COLUMN = 4;
var CHR_COLUMN = 6;
var VIDEO_COLUMN = 7;
var OTFP_COLUMN = 8;
var WAF_LOG_COLUMN = 9;
var WAF_BLOCK_COLUMN = 10;
var BADWIDTH_COLUMN = 11;

var MAINT_CUSTOMER_NAME_ROW = 3;
var MAINT_CUSTOMER_ID_ROW = 4;
var MAINT_CUSTOMER_STAT_ROW = 6;
var MAINT_CUSTOMER_VALUE_COLUMN = 2;
var fetchValues = Array();
var vclRow = 0;
var chrRow = 0;
var fastlyKey = 0;
var customerId = null;
var varFastlyCustID = 0;
var values = 0;

function setupSpreadSheet(region)
{  
  var ss = SpreadsheetApp.getActive();
  var maintSheet;
  var workingSheet;
  var otherInfoSheet;
  var certInfoSheet;
  var sheets;
  if (!region){
    region = "all";
  }
  maintSheet = ss.getSheetByName("Maintenance");
  workingSheet = ss.getSheetByName("Data - "+region);
  otherInfoSheet = ss.getSheetByName("Other customer info");
  certInfoSheet = ss.getSheetByName("Cert info");
  

  if (workingSheet == "" || workingSheet == undefined || workingSheet == null) {
    
    var sheet = ss.insertSheet();
    workingSheet = sheet.setName("Data - "+region);
    
  }
  
  if (otherInfoSheet == "" || otherInfoSheet == undefined || otherInfoSheet == null) {
    
    var sheet = ss.insertSheet();
    otherInfoSheet = sheet.setName("Other customer info");
  }
  
  if (maintSheet == "" || maintSheet == undefined || maintSheet == null) {
    
    var sheet = ss.insertSheet();
    maintSheet = sheet.setName("Maintenance");
  }
  
  if (certInfoSheet == "" || certInfoSheet == undefined || certInfoSheet == null) {
    
    var sheet = ss.insertSheet();
    certInfoSheet = sheet.setName("Cert info");
  }
  
  var deleteSheet = null;
  try {
    deleteSheet = ss.getSheetByName("Sheet1");
    if (deleteSheet) 
    ss.deleteSheet(deleteSheet);
  } catch (err) {
    
    console.error ("No Sheet1 found");
   }

  workingSheet.activate();
  
  var sheets = ss.getSheets();
  var htmlOutputFromFile;
  var htmlOutput;
 
  values = maintSheet.getDataRange().getValues();

  if (values.length > 1) {
    for(var j=1, jLen=values.length; j<jLen; j++) {
      Logger.log (values[j][0]);
      workingSheet.getRange(1, j).setValue(values[j][0]);
      fetchValues.push(values[j][1]);
    }
  } else {
    
    for (var i =0; i<workingSheetColumnNames.length; i++) {
      
      maintSheet.getRange(i+1, 1).setValue(workingSheetColumnNames[i].valueOf());    
    }
    
    for (var i =0; i<workingSheetColumnNames.length; i++) {
      
      maintSheet.getRange(i+1, 2).setValue(workingSheetColumnValues[i].valueOf());    
    }
    maintSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    maintSheet.getRange(1, 1, 1, 4).setBackground("#cccc00");
    
    maintSheet.getRange(1, 3).setValue("VCL Row Info");
    maintSheet.getRange(2, 3).setValue(vclRowInfo);
    
    maintSheet.getRange(1, 4).setValue("CHR Row Info");
    maintSheet.getRange(2, 4).setValue(chrRowInfo);
    
    var values = maintSheet.getDataRange().getValues();

    
    for (var i =0; i<otherInfoSheetCells.length; i++) {
      
      otherInfoSheet.getRange(i+1, 1).setValue((otherInfoSheetCells[i].valueOf()));    
    }
  }
  values = maintSheet.getDataRange().getValues();

  vclRow = values[1][2];
  chrRow = values[2][3];

  if (!vclRow) {
    vclRow = vclRowInfo;
  }
  if (!chrRow) {
    chrRow = chrRowInfo;
  }
  
  if (values.length > 1) {
    for(var j=1, jLen=values.length; j<jLen; j++) {
      Logger.log (values[j][0]);
      workingSheet.getRange(1, j).setFontWeight('bold');
      workingSheet.getRange(1, j).setBackground("#cccc00");
      workingSheet.getRange(1, j).setWrap(true);
      workingSheet.getRange(1, j).setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);
      workingSheet.getRange(1, j).setValue(values[j]);
      fetchValues.push(values[j][1]);
    }
  }
  
 var ui = SpreadsheetApp.getUi();
  
  htmlOutputFromFile = HtmlService.createHtmlOutputFromFile('proteus').setTitle("Please enter Customer ID & Key");
  
  ui.showSidebar(htmlOutputFromFile);
  
}

function showSite() //Disabled
{
  var ss = SpreadsheetApp.getActive();
  var htmlOutputFromFile;
  var htmlOutput;
  var ui = SpreadsheetApp.getUi();
  htmlOutput = HtmlService.createHtmlOutput(UrlFetchApp.fetch('https://proteus.app.secretcdn.net/tls/sans/').getBlob());

  ui.showModalDialog(htmlOutput, "Get TLS Info");
  
}

function getMeCustInfo(formObj)
{
  var customerName = null;
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  
  var workingSheet = ss.getSheetByName("Data - " + formObj.region);
  var maintSheet = ss.getSheetByName("Maintenance");
  var certInfoSheet = ss.getSheetByName("Cert info");
  if (!workingSheet || !maintSheet || !certInfoSheet) {
    setupSpreadSheet(formObj.region);
    var workingSheet = ss.getSheetByName("Data - " + formObj.region);
    var maintSheet = ss.getSheetByName("Maintenance");
    var certInfoSheet = ss.getSheetByName("Cert info");
  }
  var retVal = getRightApp();
  var app = retVal.thisApp;
  var ui = app.getUi();
  var fastlyLogo = HtmlService.createTemplateFromFile('fastlyLogo');
  var proteusContent = HtmlService.createTemplateFromFile('proteus');
  var svcIdFrom = 0;
  var svcIdTo = 0;
  var svcNumFrom = 0;
  var svcNumTo = 0;
  var billingInfo;
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  

  customerId = formObj.customer_id;
  var resultObj;
  fastlyKey = formObj.fastly_key;
  svcIdFrom = formObj.svc_id_from;
  svcIdTo = formObj.svc_id_to;
  svcNumFrom = parseInt(formObj.svc_no_from);
  svcNumTo = parseInt(formObj.svc_no_to);

  workingSheet.getRange(2, 1, workingSheet.getMaxRows()-1, workingSheet.getMaxColumns()).clear();
  
  if (customerId == null || customerId == undefined || !customerId) {
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .alert('You can\'t leave customer ID blank!');
  }
  
  resultObj = populateCustInfo(customerId);
  
  if (resultObj == null || resultObj == undefined) {
    showAlert("Couldn't find customer info");
    return;
  }
  
  var customerName = resultObj['customer']['name'];
  
  var now = new Date();
  var timezone = Session.getTimeZone();

  var shortDate = Utilities.formatDate(now, "GMT", 'yyyy-MM-dd');
  var fromDate = 0;
  var fromDatePlainText;
  if (svcIdFrom) {
    var frmTime = new Date(svcIdFrom).getTime();
    
    
    fromDatePlainText = Utilities.formatDate(new Date(frmTime), "GMT", 'MM dd yyyy'); // For some reason if you don't add one day it will return date input minus 1 day
    fromDate = Utilities.formatDate(new Date(frmTime), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
    
  }
  var toDate = 0;
  var toDatePlainText;
  if (svcIdTo) {
    
    var toTime = new Date(svcIdTo).getTime(); 
    
    toTime = toTime - MILLIS_PER_DAY;
    toDatePlainText = Utilities.formatDate(new Date(toTime), "GMT", 'MM dd yyyy'); // For some reason if you don't add one day it will return date input minus 1 day
    
    toTime = toTime + MILLIS_PER_DAY;
    toDate = Utilities.formatDate(new Date(toTime), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  }
  
  DriveApp.getFileById(ss.getId()).setName("Report for: " + customerName + " CID: " + customerId + " - Last run date - " + shortDate);
  
  resultObj = populateSvcInfo(customerId);
  
  values = maintSheet.getDataRange().getValues(); //If user entered new values (i.e. VCL) we fetch them once more here
  if (values.length > 1) {
    for(var j=1, jLen=values.length; j<jLen; j++) {
      Logger.log (values[j][1]);
      workingSheet.getRange(1, j).setFontWeight('bold');
      workingSheet.getRange(1, j).setBackground("#cccc00");
      workingSheet.getRange(1, j).setWrap(true);
      workingSheet.getRange(1, j).setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);
      workingSheet.getRange(1, j).setValue(values[j]);
      fetchValues.push(values[j][1]);
    }
  } else {
    showAlert("From getMeCustInfo 1:" + "Couldn't find info on Maintenance sheet please run \"Get started\" again from the Add on menu");
    return;
  }
  vclRow = values[1][2];
  chrRow = values[1][3];
  
  if (!vclRow) {
    vclRow = vclRowInfo;
  }
  if (!chrRow) {
    chrRow = chrRowInfo;
  }
  
  if (resultObj == null || resultObj == undefined) {
    showAlert("Couldn't find customer services info");
    return;
  }
  
  if (svcNumTo > Object.keys(resultObj).length || svcNumFrom > Object.keys(resultObj).length) {
    showAlert(svcNumTo > Object.keys(resultObj).length?"Service number \"To\" value larger than number of services on customer account":"Service number \"From\" value larger than number of services on customer account");
  }
  
  for (var i in resultObj) {
    Logger.log(resultObj[i]["name"]);
    Logger.log(resultObj[i]["id"]);
  }
 
  var serviceInfoFromVCL;
  var statInfo;
 
  var htmlOutput = HtmlService
  .createHtmlOutput(fastlyLogo.getRawContent() +'<div><label for="file">Services covered: 0% </label> <progress value="0" max="100"> </progress></div>')
  .setWidth(250)
  .setHeight(300).setTitle('Progress made');
  
  ui.showSidebar(htmlOutput);
  
  workingSheet.activate();

  var indexFrom = svcNumFrom?(svcNumFrom-1):0;
  var indexTo = svcNumTo?(svcNumTo):Object.keys(resultObj).length;
  var maxSvcNum = Object.keys(resultObj).length;
   
  if (formObj.fastly_svc_id) {
    var csv_fastly_svc_ids = formObj.fastly_svc_id.toString();
    
    csv_fastly_svc_ids = csv_fastly_svc_ids.replace(/\s/g, "");
    //showAlert(csv_fastly_svc_ids);
    var fastly_svc_ids = csv_fastly_svc_ids.split(",");
    
    var svcName = "";
    for(var i=0;i<fastly_svc_ids.length;i++){
      
      for (var j = 0; j<maxSvcNum; j++) {
        if (resultObj[j]["id"] == fastly_svc_ids[i]) {
          svcName = resultObj[j]["name"];
        }
      }
      workingSheet.getRange(i+2-indexFrom, CUSTOMER_NAME_COLUMN).setValue(customerName);
      workingSheet.getRange(i+2-indexFrom, CUSTOMER_ID_COLUMN).setValue(customerId);
      workingSheet.getRange(i+2-indexFrom, SERVICE_NAME_COLUMN).setValue(fastly_svc_ids[i]);
      workingSheet.getRange(i+2-indexFrom, SERVICE_ID_COLUMN).setValue(svcName);    
    }
  } else {
    for (var i = indexFrom; i<indexTo; i++) {
      //showAlert ("Value of i " + i);
      workingSheet.getRange(i+2-indexFrom, CUSTOMER_NAME_COLUMN).setValue(customerName);
      workingSheet.getRange(i+2-indexFrom, CUSTOMER_ID_COLUMN).setValue(customerId);
      workingSheet.getRange(i+2-indexFrom, SERVICE_NAME_COLUMN).setValue(resultObj[i]["name"]);
      workingSheet.getRange(i+2-indexFrom, SERVICE_ID_COLUMN).setValue(resultObj[i]["id"]);    
    }
  }
    
  if (formObj.fastly_svc_id) {
    var csv_fastly_svc_ids = formObj.fastly_svc_id.toString();
    csv_fastly_svc_ids = csv_fastly_svc_ids.split(" ").join("");
    
    var fastly_svc_ids = csv_fastly_svc_ids.split(",");
    var version = 1;
    for(var i=0;i<fastly_svc_ids.length;i++){
      
      for (var j = 0; j<maxSvcNum; j++) {
        if (resultObj[j]["id"] == fastly_svc_ids[i]) {
          version = resultObj[j]["version"];
        }
      }
      
      serviceInfoFromVCL = getServiceInfoFromVCL(fastly_svc_ids[i], version, fetchValues, vclRow, fastlyKey);
      if (serviceInfoFromVCL == undefined)
        return;    
      
      for (var j = 0; j<serviceInfoFromVCL.length; j++) {
        
        workingSheet.getRange(i+2-indexFrom, vclRow-1+j).setValue(serviceInfoFromVCL[j]);
      }
      var htmlOutput = HtmlService
      .createHtmlOutput(fastlyLogo.getRawContent() + '<div> <p> Currently gathering info on Service #'+ Math.round(i+1) + ' out of: ' + Object.keys(resultObj).length + ' services </p> </div>')
      .setWidth(250)
      .setHeight(300).setTitle('Progress made');
      
      ui.showSidebar(htmlOutput);
    }
  } else {
    for (var i = indexFrom; i<indexTo; i++)
    {
      serviceInfoFromVCL = getServiceInfoFromVCL(resultObj[i]["id"], resultObj[i]["version"], fetchValues, vclRow, fastlyKey);
      if (serviceInfoFromVCL == undefined)
        return;    
      
      for (var j = 0; j<serviceInfoFromVCL.length; j++) {
        
        workingSheet.getRange(i+2-indexFrom, vclRow-1+j).setValue(serviceInfoFromVCL[j]);
      }
      var htmlOutput = HtmlService
      .createHtmlOutput(fastlyLogo.getRawContent() + '<div> <p> Currently gathering info on Service #'+ Math.round(i+1) + ' out of: ' + Object.keys(resultObj).length + ' services </p> </div>')
      .setWidth(250)
      .setHeight(300).setTitle('Progress made');
      
      ui.showSidebar(htmlOutput);
    }
  }
  
  if (formObj.fastly_svc_id) {
    var csv_fastly_svc_ids = formObj.fastly_svc_id.toString();
    csv_fastly_svc_ids = csv_fastly_svc_ids.replace(/\s/g, "");
    var fastly_svc_ids = csv_fastly_svc_ids.split(",");
    
    for(var i=0;i<fastly_svc_ids.length;i++){
      statInfo = getStatInfo(fastly_svc_ids[i], fetchValues, vclRow, chrRow, fromDate, toDate, formObj.region, fastlyKey);
      if (statInfo == undefined)
        return;
      
      for (var j = 0; j<statInfo.length; j++) {
        workingSheet.getRange(i+2-indexFrom, MAINT_CUSTOMER_STAT_ROW-1+j).setValue(statInfo[j]);
      }
      
      var htmlOutput = HtmlService
      .createHtmlOutput(fastlyLogo.getRawContent() + '<div> <label for="file">Services covered:' + Math.round(((i+1)*100)/Object.keys(resultObj).length) +'% </label> <progress value=' + Math.round(((i+1)*100)/Object.keys(resultObj).length) + ' max="100"> </progress> <p>Currently working on Service #'+ Math.round(i+1) + ' out of: ' + Object.keys(resultObj).length + " services" + '</p> </div>')
      .setWidth(250)
      .setHeight(300).setTitle('Progress made');
      
      ui.showSidebar(htmlOutput);
    }
  } else { 
    for (var i = indexFrom; i<indexTo; i++)
    {
      statInfo = getStatInfo(resultObj[i]["id"], fetchValues, vclRow, chrRow, fromDate, toDate, formObj.region, fastlyKey);
      if (statInfo == undefined)
        return;
      
      for (var j = 0; j<statInfo.length; j++) {
        workingSheet.getRange(i+2-indexFrom, MAINT_CUSTOMER_STAT_ROW-1+j).setValue(statInfo[j]);
      }
      
      var htmlOutput = HtmlService
      .createHtmlOutput(fastlyLogo.getRawContent() + '<div> <label for="file">Services covered:' + Math.round(((i+1)*100)/Object.keys(resultObj).length) +'% </label> <progress value=' + Math.round(((i+1)*100)/Object.keys(resultObj).length) + ' max="100"> </progress> <p>Currently working on Service #'+ Math.round(i+1) + ' out of: ' + Object.keys(resultObj).length + " services" + '</p> </div>')
      .setWidth(250)
      .setHeight(300).setTitle('Progress made');
      
      ui.showSidebar(htmlOutput);
    }
  }
  
  if (!svcIdTo || !svcIdFrom)
    workingSheet.getRange(1, 1).setValue("Customer Name - (All data from past 90 days. Pulled on " + shortDate + ")");
  else 
    workingSheet.getRange(1, 1).setValue("Customer Name (All data from: " + fromDatePlainText + " To: " + toDatePlainText + ")");
  
  var htmlOutput = HtmlService
  .createHtmlOutput(proteusContent.getRawContent() +'<div><label for="file">Services covered:' + Math.round((i*100)/Object.keys(resultObj).length) +'% </label> <progress value=' + Math.round((i*100)/Object.keys(resultObj).length) + ' max="100"> </progress></div>')
  .setWidth(250)
  .setHeight(300).setTitle('Progress made');
  
  ui.showSidebar(htmlOutput);
 
  var nowDate = new Date();
  
  if (formObj.svc_billingMonth && formObj.svc_billingYear) {
    billingInfo = getBillingInfo(formObj.customer_id, formObj.svc_billingMonth, formObj.svc_billingYear, fastlyKey);
  } else {
    var ninetyDaysAgo = new Date(now.getTime() - (MILLIS_PER_DAY * 91)); //Only when you give 91 will it return 90 days for some strange reason
    var billingMonth = Utilities.formatDate(ninetyDaysAgo, timezone, 'MM');
    var billingYear = Utilities.formatDate(ninetyDaysAgo, timezone, 'YYYY');
    billingInfo = getBillingInfo(formObj.customer_id, billingMonth, billingYear, fastlyKey);
  }
  if (formObj.acc_certData == "on") {
   getCertInfo(formObj.customer_id, fastlyKey); 
  }

}

function getCertInfo(customerId, fastlyKey)
{
  var options;
  var options_gs;
  var retValArr = Array();
  var result;
  var resultDomains;
  var totalDomains;
  var globalsignDomains;
  var parsedResult;
  var parsedResultDomains;
  var parsedResultTotalDomains;
  var parsedResultGlobalsignDomains;
  var statInfo;
  var ss = SpreadsheetApp.getActive();
  var certInfoSheet = ss.getSheetByName("Cert info");
  var values;
  var pageNum = 0;

  certInfoSheet.clear();

  if (fastlyKey) {
    options = {
      "method": "get",
      'contentType': 'application/json',
      "headers": {"Fastly-Key": fastlyKey, "Accept": "application/vnd.api+json"},
      "muteHttpExceptions": true
      
    }
    
    options_gs = {
      "method": "get",
      'contentType': 'application/json',
      "headers": {"Fastly-Key": fastlyKey},
      "muteHttpExceptions": true
      
    }
  } else {
    options = {
      "method": "get",
      'contentType': 'application/json',
      "headers": {"Fastly-Key": "", "Accept": "application/vnd.api+json"},
      "muteHttpExceptions": true
    }
    
    options_gs = {
      "method": "get",
      'contentType': 'application/json',
      "headers": {"Fastly-Key": ""},
      "muteHttpExceptions": true
      
    }
  }
  
  do {
    pageNum++;
    Logger.log("This is the page number now" + pageNum);
    try {
      
      result = UrlFetchApp.fetch('https://api.fastly.com/tls/subscriptions?customer_id=' + customerId + "&page%5Bnumber%5D=" + pageNum, options);
      
    }  catch (err) {
      showAlert("From getCertInfo 1:" + err);
      return undefined;
    }
    
    values = certInfoSheet.getDataRange().getValues();
    parsedResult = JSON.parse(result.getContentText());
    
    var errCheck = Object.keys(parsedResult);
    if(errCheck[0] == "errors") {
      return retValArr;
    }
    
    var jIndex = 0;
    var iIndex = 0;
    if (parsedResult['data']) {
      if (pageNum < 2) {
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setValue("Fastly TLS Cert Info");
        
        
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setBackground('#80dfff');
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setFontWeight('bold');
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setWrap(true);
        
        iIndex++
      }
      for (var i in parsedResult['data']) {
        var certAttr = parsedResult['data'][i]['attributes'];
        for (j in certAttr) {
          certInfoSheet.getRange(1+iIndex+jIndex, 1).setValue(j);
          certInfoSheet.getRange(1+iIndex+jIndex, 2).setValue(parsedResult['data'][i]['attributes'][j]);
          certInfoSheet.getRange(1+iIndex+jIndex, 1).setBackground('#80dfff');
          certInfoSheet.getRange(1+iIndex+jIndex, 1).setFontWeight('bold');
          certInfoSheet.getRange(1+iIndex+jIndex, 1).setWrap(true);
          
          certInfoSheet.getRange(1+iIndex+jIndex, 2).setFontStyle('italic');
          certInfoSheet.getRange(1+iIndex+jIndex, 2).setWrap(true);
          certInfoSheet.getRange(1+iIndex+jIndex, 2).setBackground("#ffff00");
          jIndex++;
          
        }
        
        //Get the domains for this cert
        certInfoSheet.getRange(1+iIndex+jIndex, 1).setValue("Fastly TLS Cert ID: "+ Math.round(Math.round(i)+1));
        certInfoSheet.getRange(1+iIndex+jIndex, 1).setBackground('#80dfff');
        certInfoSheet.getRange(1+iIndex+jIndex, 1).setFontWeight('bold');
        certInfoSheet.getRange(1+iIndex+jIndex, 1).setWrap(true);
        certInfoSheet.getRange(1+jIndex+iIndex, 2).setValue(parsedResult['data'][i]['id']);
        
        certInfoSheet.getRange(1+iIndex+jIndex, 2).setFontStyle('italic');
        certInfoSheet.getRange(1+iIndex+jIndex, 2).setWrap(true);
        certInfoSheet.getRange(1+iIndex+jIndex, 2).setBackground("#ffff00");
        for (var k in parsedResult['data'][i]['relationships']['tls_domains']['data']) {
          certInfoSheet.getRange(2+iIndex+jIndex, 2).setValue("Cert domain: " + (Math.round(k)+1));
          certInfoSheet.getRange(2+iIndex+jIndex, 3).setValue(parsedResult['data'][i]['relationships']['tls_domains']['data'][k]['id']);
          certInfoSheet.getRange(2+iIndex+jIndex, 2).setBackground('#80dfff');
          certInfoSheet.getRange(2+iIndex+jIndex, 2).setFontWeight('bold');
          certInfoSheet.getRange(2+iIndex+jIndex, 2).setWrap(true);
          
          certInfoSheet.getRange(2+iIndex+jIndex, 3).setFontStyle('italic');
          certInfoSheet.getRange(2+iIndex+jIndex, 3).setWrap(true);
          certInfoSheet.getRange(2+iIndex+jIndex, 3).setBackground("#ffff00");
          iIndex++;
        }
        jIndex++;
        iIndex++;
        
      }
    }
    Logger.log("This is the value per_page for subscriptions " + parsedResult['meta']['per_page']);
    Logger.log("This is the current page for subscriptions " + parsedResult['meta']['current_page']);
    Logger.log("This is the record count for subscriptions " + parsedResult['meta']['record_count']);
    Logger.log("This is the total pages for subscriptions " + parsedResult['meta']['total_pages']);
  } while (pageNum < parseInt (parsedResult['meta']['total_pages']));
  
  pageNum = 0;
  
  do {
    pageNum++;
    
    try {
      
      resultDomains = UrlFetchApp.fetch('https://api.fastly.com/tls/certificates?customer_id=' + customerId  + "&page%5Bnumber%5D=" + pageNum, options);
      
    }  catch (err) {
      showAlert("From getCertInfo 1:" + err);
      return undefined;
    }
    
    parsedResultDomains = JSON.parse(resultDomains.getContentText());
    
    var errCheck = Object.keys(parsedResultDomains);
    if(errCheck[0] == "errors") {
      return retValArr;
    }
    
    jIndex++;
    if (pageNum < 2) {
      certInfoSheet.getRange(1+jIndex+iIndex, 1).setValue("Fastly Hosted Certs");

      certInfoSheet.getRange(1+jIndex+iIndex, 1).setBackground('#80dfff');
      certInfoSheet.getRange(1+jIndex+iIndex, 1).setFontWeight('bold');
      certInfoSheet.getRange(1+jIndex+iIndex, 1).setWrap(true);
      
      iIndex++
    }
      for (var i in parsedResultDomains['data']) {
        
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setValue("Cert ID: " + Math.round(Math.round(i)+1+((pageNum-1)*20)));
        certInfoSheet.getRange(1+jIndex+iIndex, 2).setValue(parsedResultDomains['data'][i]['id']);
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setBackground('#80dfff');
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setFontWeight('bold');
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setWrap(true);
        
        certInfoSheet.getRange(1+jIndex+iIndex, 2).setFontStyle('italic');
        certInfoSheet.getRange(1+jIndex+iIndex, 2).setWrap(true);
        certInfoSheet.getRange(1+jIndex+iIndex, 2).setBackground("#ffff00");
        iIndex++;
        for (var k in parsedResultDomains['data'][i]['relationships']['tls_domains']['data']) {
          certInfoSheet.getRange(1+jIndex+iIndex, 2).setValue("Cert domain: " + (Math.round(k)+1));
          certInfoSheet.getRange(1+jIndex+iIndex, 3).setValue(parsedResultDomains['data'][i]['relationships']['tls_domains']['data'][k]['id']);
          certInfoSheet.getRange(1+jIndex+iIndex, 2).setBackground('#80dfff');
          certInfoSheet.getRange(1+jIndex+iIndex, 2).setFontWeight('bold');
          certInfoSheet.getRange(1+jIndex+iIndex, 2).setWrap(true);
          
          certInfoSheet.getRange(1+jIndex+iIndex, 3).setFontStyle('italic');
          certInfoSheet.getRange(1+jIndex+iIndex, 3).setWrap(true);
          certInfoSheet.getRange(1+jIndex+iIndex, 3).setBackground("#ffff00");
          iIndex++;
        }
      }
  
    Logger.log("This is the value per_page for domains " + parsedResultDomains['meta']['per_page']);
    Logger.log("This is the current page for domains " + parsedResultDomains['meta']['current_page']);
    Logger.log("This is the record count for domains " + parsedResultDomains['meta']['record_count']);
    Logger.log("This is the total pages for domains " + parsedResultDomains['meta']['total_pages']);
  } while (pageNum < parseInt (parsedResultDomains['meta']['total_pages']));
  
  pageNum = 0;
  iIndex++;
  
  do {
    pageNum++;
    
    try {
      
      totalDomains = UrlFetchApp.fetch('https://api.fastly.com/tls/activations?customer_id=' + customerId + "&page%5Bnumber%5D=" + pageNum, options);
      
    }  catch (err) {
      showAlert("From getCertInfo 1:" + err);
      return undefined;
    }
    
    parsedResultTotalDomains = JSON.parse(totalDomains.getContentText());
    
    var errCheck = Object.keys(parsedResultTotalDomains);
    if(errCheck[0] == "errors") {
      return retValArr;
    }
    
    if (parsedResultTotalDomains['data']) {
      if (pageNum < 2) {
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setValue("Total TLS Activations");
        
        
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setBackground('#80dfff');
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setFontWeight('bold');
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setWrap(true);
        jIndex++;
      }
      
      for (var i in parsedResultTotalDomains['data']) {
        
        //Get the domains for this cert
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setValue("TLS Domain: " + Math.round(Math.round(i)+1+(100*(pageNum-1))));
        certInfoSheet.getRange(1+jIndex+iIndex, 2).setValue(parsedResultTotalDomains['data'][i]['relationships']['tls_domain']['data']['id']);
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setBackground('#80dfff');
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setFontWeight('bold');
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setWrap(true);
        
        certInfoSheet.getRange(1+jIndex+iIndex, 2).setFontStyle('italic');
        certInfoSheet.getRange(1+jIndex+iIndex, 2).setWrap(true);
        certInfoSheet.getRange(1+jIndex+iIndex, 2).setBackground("#ffff00");
        
        iIndex++;
        
      }
    }
    Logger.log("This is the value per_page for activations " + parsedResultTotalDomains['meta']['per_page']);
    Logger.log("This is the current page for activation " + parsedResultTotalDomains['meta']['current_page']);
    Logger.log("This is the record count for activations " + parsedResultTotalDomains['meta']['record_count']);
    Logger.log("This is the total pages for activations " + parsedResultTotalDomains['meta']['total_pages']);
  } while (pageNum < parseInt (parsedResultTotalDomains['meta']['total_pages']));

  pageNum = 0;
  iIndex++;
 
  var carryOverCert = "";
  
    try {
      
      globalsignDomains = UrlFetchApp.fetch('https://api.fastly.com/tls/globalsign/domains?customer_id=' + customerId, options_gs);
      
    }  catch (err) {
      showAlert("From getCertInfo 1:" + err);
      return undefined;
    }
    
    parsedResultGlobalsignDomains = JSON.parse(globalsignDomains.getContentText());
    
    var errCheck = Object.keys(parsedResultGlobalsignDomains);
    if(errCheck[0] == "errors") {
      return retValArr;
    }
    
  if (parsedResultGlobalsignDomains) {
    
      certInfoSheet.getRange(1+jIndex+iIndex, 1).setValue("Total Globalsign Domains");
      
      certInfoSheet.getRange(1+jIndex+iIndex, 1).setBackground('#80dfff');
      certInfoSheet.getRange(1+jIndex+iIndex, 1).setFontWeight('bold');
      certInfoSheet.getRange(1+jIndex+iIndex, 1).setWrap(true);
      jIndex++;
    
    var domainNumber = 1;
    for (var i in parsedResultGlobalsignDomains) {
      //certInfoSheet.getRange(1+jIndex+iIndex, 2).setValue(parsedResultGlobalsignDomains['data'][i]['certificate_id']);
      if (carryOverCert != parsedResultGlobalsignDomains[i]['certificate_id']) {
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setValue("Globalsign Cert ID: " + Math.round(Math.round(i)+1));
        certInfoSheet.getRange(1+jIndex+iIndex, 2).setValue(parsedResultGlobalsignDomains[i]['certificate_id']);
        carryOverCert = parsedResultGlobalsignDomains[i]['certificate_id'];
        
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setBackground('#80dfff');
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setFontWeight('bold');
        certInfoSheet.getRange(1+jIndex+iIndex, 1).setWrap(true);
        
        certInfoSheet.getRange(1+jIndex+iIndex, 2).setFontStyle('italic');
        certInfoSheet.getRange(1+jIndex+iIndex, 2).setWrap(true);
        certInfoSheet.getRange(1+jIndex+iIndex, 2).setBackground("#ffff00");
        
        iIndex++;
        domainNumber = 1;
      }
      
      certInfoSheet.getRange(1+jIndex+iIndex, 2).setValue("Globalsign Domain: " + domainNumber);
      certInfoSheet.getRange(1+jIndex+iIndex, 3).setValue(parsedResultGlobalsignDomains[i]['fqdn']);
      certInfoSheet.getRange(1+jIndex+iIndex, 2).setBackground('#80dfff');
      certInfoSheet.getRange(1+jIndex+iIndex, 2).setFontWeight('bold');
      certInfoSheet.getRange(1+jIndex+iIndex, 2).setWrap(true);
      
      certInfoSheet.getRange(1+jIndex+iIndex, 3).setFontStyle('italic');
      certInfoSheet.getRange(1+jIndex+iIndex, 3).setWrap(true);
      certInfoSheet.getRange(1+jIndex+iIndex, 3).setBackground("#ffff00");
      iIndex++;
      domainNumber++;
    }
    
  }
  
  return retValArr;
}

function getBillingInfo(customerId, month, year, fastlyKey)
{
  var options;
  var retValArr = Array();
  var result;
  var parsedResult = null;
  var parsedResult_preProc = null;
  var statInfo;
  var ss = SpreadsheetApp.getActive();
  var otherInfoSheet = ss.getSheetByName("Other customer info");
  var values;
  var ourMonths = {
    "January" : "01", 
    "February" : "02", 
    "March" : "03", 
    "April" : "04", 
    "May" : "05", 
    "June": "06", 
    "July" : "07", 
    "August" : "08", 
    "September" : "09", 
    "October" : "10", 
    "November" : "11", 
    "December" : "12"};
  var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  otherInfoSheet.clear();
  for (var i =0; i<otherInfoSheetCells.length; i++) {
    otherInfoSheet.getRange(i+1, 1).setValue((otherInfoSheetCells[i].valueOf()));    
  }

  if (fastlyKey) {
    options = {
      "method": "get",
      'contentType': 'application/json',
      "headers": {"Fastly-Key": fastlyKey},
      "muteHttpExceptions": true
    }
  } else {
    options = {
      "method": "get",
      'contentType': 'application/json',
      "headers": {"Fastly-Key": ""},
      "muteHttpExceptions": true
    }
  }
  
  try {
    
    result = UrlFetchApp.fetch('https://api.fastly.com/billing/v2/account_customers/' + customerId + '/invoices', options);
    
  }  catch (err) {
      showAlert("From getBillingInfo 1:" + err);
      return undefined;
  }
  
  var monthStr = ourMonths[months[Math.round(month)-1]];
  parsedResult_preProc = JSON.parse(result.getContentText());
  
   Logger.log(parsedResult_preProc);
  if (parsedResult_preProc == null) {
    return;
  }
  
    for (var i = 0; i < Object.keys(parsedResult_preProc).length; i++)
    {
      try { 
        if (parsedResult_preProc[i]['end_time'].search(year + "-" + monthStr) > -1) {
          parsedResult = parsedResult_preProc[i];
          break;
        }
      } catch (err) {
        Logger.log("Couldn't find the bill in this page");
      } finally {
        continue;
      }
    }

    Logger.log(parsedResult);
  if (parsedResult == null) {
    return;
  }
         
  var errCheck = Object.keys(parsedResult);
  if(errCheck[0] == "errors") {
    return retValArr;
  } else {
    if (parsedResult['total'] != undefined) {
      if (parsedResult['total']['cost'] != undefined) {
        retValArr.push({"Cost": parsedResult['total']['cost']});
      } else {
        retValArr.push({"Cost": "Couldn't fetch. Please rerun the app and try again"});
      }
      
      if (parsedResult['total']['discount'] != undefined) {
        retValArr.push({"Discount": parsedResult['total']['discount']});
      } else {
        retValArr.push({"Discount": "Couldn't fetch. Please rerun the app and try again"});
      }
      
      if (parsedResult['total']['bandwidth'] != undefined) {
        retValArr.push({"Bandwidth in GB": parsedResult['total']['bandwidth']});
      } else {
        retValArr.push({"Bandwidth in GB": "Couldn't fetch. Please rerun the app and try again"});
      }
      
      if (parsedResult['total']['incurred_cost'] != undefined) {
        retValArr.push({"Incurred Cost": parsedResult['total']['incurred_cost']});
      } else {
        retValArr.push({"Incurred Cost": "Couldn't fetch. Please rerun the app and try again"});
      }
      
      if (parsedResult['total']['cost_before_discount'] != undefined) {
        retValArr.push({"Cost Before Discount": parsedResult['total']['cost_before_discount']});
      } else {
        retValArr.push({"Cost Before Discount": "Couldn't fetch. Please rerun the app and try again"});
      }
      
      if (parsedResult['total']['extras'] != undefined) {
        for (var i in parsedResult['total']['extras']) {
        
          retValArr.push({"Extra Item -" :parsedResult['total']['extras'][i]['name']}, {"Recurring cost -" :parsedResult['total']['extras'][i]['recurring']});
        
        }
      } else {
        retValArr.push({"Extra Item -": "Couldn't fetch. Please rerun the app and try again"});
      }
      
      if (parsedResult['total']['line_items'] != undefined) {
        for (var i in parsedResult['line_items']) {
          retValArr.push({"Line Item -" :parsedResult['line_items'][i]['description']});
        }
      }
      
    } else {
        retValArr.push({"Billing Info -": "Couldn't fetch for this run. Please rerun the app once more"});
    }
    
    values = otherInfoSheet.getDataRange().getValues();
    
    otherInfoSheet.getRange(values.length+1, 1).setValue("Billing MTD for "+ months[Math.round(month)-1] + ", " + year);
    otherInfoSheet.getRange(values.length+1, 1).setBackground("#cccc00");
    otherInfoSheet.getRange(values.length+1, 1).setFontWeight('bold');
    otherInfoSheet.getRange(values.length+1, 1).setWrap(true);
    //otherInfoSheet.getRange(values.length+1, 1).setValue(cellValue);
    values = otherInfoSheet.getDataRange().getValues();
    
    for (var i in retValArr) {
         
      for (j in retValArr[i]) {
      
        otherInfoSheet.getRange(values.length+Math.round(i)+1, 1).setBackground('#80dfff');
        otherInfoSheet.getRange(values.length+Math.round(i)+1, 1).setFontWeight('bold');
        otherInfoSheet.getRange(values.length+Math.round(i)+1, 1).setWrap(true);
        otherInfoSheet.getRange(values.length+Math.round(i)+1, 1).setValue(j);
        
        otherInfoSheet.getRange(values.length+Math.round(i)+1, 2).setFontStyle('italic');
        otherInfoSheet.getRange(values.length+Math.round(i)+1, 2).setWrap(true);
        otherInfoSheet.getRange(values.length+Math.round(i)+1, 2).setBackground("#ffff00");
        otherInfoSheet.getRange(values.length+Math.round(i)+1, 2).setValue(retValArr[i][j]);
      }
    }
  }  
  return retValArr;
}

function showAlert(alertMsg)
{

  var retVal = getRightApp();
  var thisApp = retVal.thisApp;
  var thisAppString = retVal.thisAppString;

  var ui = thisApp.getUi();
  ui.alert(alertMsg);
}

function getStatInfo(serviceId, fetchValues, vclRow, chrRow, fromDate, toDate, region, fastlyKey)
{
  var options;
  var retValArr = Array();
  var result;
  var parsedResult;
  var statInfo;
  if (fastlyKey) {
    options = {
      "method": "get",
      'contentType': 'application/json',
      "headers": {"Fastly-Key": fastlyKey},
      "muteHttpExceptions": true
    }
  } else {
    options = {
      "method": "get",
      'contentType': 'application/json',
      "headers": {"Fastly-Key": ""},
      "muteHttpExceptions": true
    }
  }
  
  for (var f=chrRow-2; f<(vclRow-2); f++) {
    statInfo = 0;
    try {
      if (fromDate || toDate) {
        
        result = UrlFetchApp.fetch('https://api.fastly.com/stats/service/'+ serviceId +'/field/'+ fetchValues[f] + '?from='+fromDate + '&to=' + toDate + '&region=' + region, options);
      } else {
        result = UrlFetchApp.fetch('https://api.fastly.com/stats/service/'+ serviceId +'/field/'+ fetchValues[f] + '?from=90+days+ago' + '&region=' + region, options);
      }
      
    } catch (err) {
      showAlert("From getStatInfo 1:" + err);
      return undefined;
    }
    
    parsedResult = JSON.parse(result.getContentText());
    //Logger.log(result);
    if (result.getResponseCode() != 200) {
      //Logger.log("Code " + result.getResponseCode() + "result " + result + " f " + f + " vclRow " + vclRow + " chrRow " + chrRow);
      showAlert("From getStatInfo 2:" + result.getContentText());
      return undefined;
    }
    

    if (fetchValues[f] == "hit_ratio") {
      var totalHits = 0;
      for (var i in parsedResult['data']) {
        totalHits += parsedResult['data'][i]['hit_ratio'];
      }
      if (totalHits) {
        statInfo = totalHits/parsedResult['data'].length;
        statInfo *= 100;
        var tempStat = parseFloat(statInfo);
        statInfo = parseFloat(tempStat.toFixed(2));
        //Math.round(statInfo);
        statInfo = statInfo + "%"
      } else
        statInfo = 0;
      
      //showAlert("This is the total CHR " + statInfo);
    } else if (fetchValues[f] == "bandwidth") {
      var bandwidth = 0;
      for (var i in parsedResult['data']) {
        bandwidth += parsedResult['data'][i]['bandwidth'];
      }
      if (bandwidth) {
        statInfo = bandwidth/(1000*1000*1000);
        //statInfo *= 100;
        //Math.round(statInfo);
        //statInfo = statInfo + "%"
      } else
        statInfo = 0;
      
      //showAlert("This is the total CHR " + statInfo);
    } else {
      var result = 0;
      for (var i in parsedResult['data']) {
        result += parsedResult['data'][i][fetchValues[f]];
        
      }
      statInfo = result;
      
    }  

      retValArr.push(statInfo);
  }
  //Logger.log(retValArr);
  return retValArr;
  
  //Logger.log(result);
  return;
}

function getServiceInfoFromVCL(serviceId, version, fetchValues, vclRow, fastlyKey)
{
  var options;
  var retValArr = Array();
  var result;
  var serviceInfo;
  if (fastlyKey) {
    options = {
      "method": "get",
      'contentType': 'application/json',
      "headers": {"Fastly-Key": fastlyKey},
      "muteHttpExceptions": true
    }
  } else {
    options = {
      "method": "get",
      'contentType': 'application/json',
      "headers": {"Fastly-Key": fastlyKey},
      "muteHttpExceptions": true
    }
  }
  
  try {
    result = UrlFetchApp.fetch('https://api.fastly.com/service/'+ serviceId +'/version/'+ version + '/generated_vcl', options);
  } catch (err) {
    showAlert("From getServiceInfoFromVCL 1 error message: " + err);
    return undefined;
  }
  
  //Logger.log(result);
  if (result.getResponseCode() == 403) {
      retValArr.push("N/A C@E Svc");
      //return retValArr;
   } else if (result.getResponseCode() != 200) {
    //Logger.log("Code " + result.getResponseCode() + "result " + result);
    
    showAlert("From getServiceInfoFromVCL 1:" + result.getContentText() + " for serviceID: " + serviceId + " Code:" + result.getResponseCode() + "result " + result);
    
    return undefined;
  }
  
  serviceInfo = result.getContentText();
    
  //Logger.log ("fetchValues " + fetchValues + " vclRow " + vclRow);
  for (var i=vclRow-2; i<fetchValues.length; i++) {
    var searchTerm = fetchValues[i];
    //Logger.log(searchTerm.test("log *{"));
    //Logger.log(searchVerb.search(searchTerm));
    //if (serviceInfo.indexOf(fetchValues[i]) > -1)
    if (result.getResponseCode() == 403) {
      retValArr.push("N/A C@E Svc");
      continue;
    }
    
    if (serviceInfo.search(searchTerm) > -1)
    {
      //Logger.log ("Index of " + fetchValues[i] + " is " + serviceInfo.indexOf(fetchValues[i]));
      retValArr.push("Yes");
    } else {
      retValArr.push("No");
    }
    //Logger.log("Array value at " + i + " is" + retValArr[i]);
  }
  //Logger.log(retValArr);
  return retValArr;
  
  //Logger.log(result);
  return;
}

function populateCustInfo(customerId)
{
  var options;
  
  if (fastlyKey) {
    options = {
      "method": "get",
      'contentType': 'application/json',
      "headers": {"Fastly-Key": fastlyKey},
      "muteHttpExceptions": true
    }
  } else {
    options = {
      "method": "get",
      'contentType': 'application/json',
      "headers": {"Fastly-Key": ""},
      "muteHttpExceptions": true
    }
  }
  
  var result;
  //custId = maintSheet.getRange(CUSTOMER_NAME_COLUMN, 1).getValue();
  
  try {
    result = UrlFetchApp.fetch('https://api.fastly.com/customer/details/'+ customerId, options);
  } catch (err) {
    showAlert("From populateCustInfo 1:" + err);
    return undefined;
  }
  
  if (result .getResponseCode() != 200) {
    Logger.log("Code " + result.getResponseCode() + "result " + result);
    showAlert("From populateCustInfo 2:" + result.getContentText());
    return undefined;
  }

  return(JSON.parse(result.getContentText()));
}


function populateSvcInfo(customerId)
{
  var options;
  
  if (fastlyKey) {
    options = {
      "method": "get",
      'contentType': 'application/json',
      "headers": {"Fastly-Key": fastlyKey},
      "muteHttpExceptions": true
    }
  } else {
    options = {
      "method": "get",
      'contentType': 'application/json',
      "headers": {"Fastly-Key": ""},
      "muteHttpExceptions": true
    }
  }
  
  var result;
  //custId = maintSheet.getRange(CUSTOMER_NAME_COLUMN, 1).getValue();
  
  try {
    result = UrlFetchApp.fetch('https://api.fastly.com/customer/'+ customerId +'/services', options);
  } catch (err) {
    showAlert("From populateSvcInfo 1:" + err);
    return undefined;
  }
  
  if (result .getResponseCode() != 200) {
    Logger.log("Code " + result.getResponseCode() + "result " + result);
    showAlert("From populateSvcInfo 1:" + result.getContentText());
    return undefined;
  }

  return(JSON.parse(result.getContentText()));
}

function onOpen (e)
{
  var uiTypeSpreadsheet = null;
  var uiTypeDocument = null;
  var uiTypeSlides = null;
  var thisApp = null;
  var thisAppString;
  var retVal = getRightApp();
  thisApp = retVal.thisApp;
  thisAppString = retVal.thisAppString;

  
  var ui = thisApp.getUi();
  
  //var ui = SpreadsheetApp.getUi();
  
  ui.createAddonMenu()
  .addItem('Get started', 'setupSpreadSheet')
  .addSeparator()
  .addItem('Pls contribute to code at - https://github.com/lotusbaba/Customer-Info-App-Fastly', 'https://github.com/lotusbaba/Customer-Info-App-Fastly')
  .addToUi();

}

function writeToDoc()
{
  var thisApp = getRightApp();
  var ss = SpreadsheetApp.getActive();
  var workingSheet = ss.getSheetByName("Data");
}

function onInstall(e) {
  onOpen(e);
}
