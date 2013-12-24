// <!-- This bit by @mhawksey 
// Available under Creative Commons Attribution-ShareAlike 2.5 UK: Scotland License
var msg ="";
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "Configure", functionName: "configureAPI"},
                      {name: "Test Connection", functionName: "aTest"},
                      {name: "Clear Current Sheet", functionName: "clearSheet"},
                      {name: "Get Friends", functionName: "getFriends"},
                      {name: "Get Followers", functionName: "getFollowers"},
                      {name: "Get other persons followers", functionName: "getAnotherFollowers"},
                      {name: "Get other persons friends", functionName: "getAnotherFriends"} ];
  ss.addMenu("Twitter", menuEntries);
}


function getFriends(){
  getFriendAndFo("friends", "friends"); // using part of the Twitter API call also as sheet name
}


function getFollowers(){
  getFriendAndFo("followers", "followers"); // using part of the Twitter API call also as sheet name
}

function getAnotherFollowers(){
  var screenName = Browser.inputBox("Enter screen name of person you'd like to get followers for:");
  if (screenName){
    getAnother(screenName, "followers");
  }
}

function getAnotherFriends(){
  var screenName = Browser.inputBox("Enter screen name of person you'd like to get friends for:");
  if (screenName){
    getAnother(screenName, "friends");
  }
}

function getAnother(screenName, type){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pro = isProtected(screenName);
  if (isProtected(screenName)){
    Browser.msgBox("This users account is protected. Data is not available");
    return;
  }
  if (!ss.getSheetByName(type + " - " + screenName)){
    var temp = ss.getSheetByName("TMP");
    
    // fix for current bug on insertSheet using template
    ss.setActiveSheet(temp);
    var sheet = ss.duplicateActiveSheet();
    ss.setActiveSheet(sheet);
    ss.renameActiveSheet(type + " - " + screenName);
    
    //var sheet = ss.insertSheet(type + " - " + screenName, {template:temp});
  }
  
  getFriendAndFo(type + " - " + screenName, type, screenName);
}

function isProtected(screenName){
  var url = "users/show.json?screen_name="+screenName;
  var o = tw_request("GET", url).protected;
  return o;
}

function getFriendAndFo(sheetName, friendOrFo, optScreenName){
  var statusString = "";
  var your_screen_name = tw_request("GET", "account/verify_credentials.json").screen_name;
  if (typeof optScreenName != "undefined") {
    statusString = "&screen_name="+optScreenName;
    your_screen_name = optScreenName;
  } 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Looking for data ...");
  var sheet = ss.getSheetByName(sheetName); 
  if (sheet.getLastRow()>1) {
    var existing_ids = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues();
    var existing_ids_str = existing_ids.join();
  } else {
    var existing_ids_str = "";
  }
  var cursor = "-1";
  var count = 0;
  var users = [];
  while(cursor != "none"){ // while twitter returns data loop
    try {
      var url = friendOrFo+"/ids.json?cursor=" + cursor + statusString;
      var o = tw_request("GET", url);
      for (j in o.ids){
          // get new ids delete missing ones
        var searchKey = new RegExp(o.ids[j].toString(),"gi");
        var test = existing_ids_str.search(searchKey);
        if (existing_ids_str.search(searchKey) === -1){
          users.push(o.ids[j]);
        }
      }
      if (o.next_cursor!="0" || count < 6){
        cursor = o.next_cursor; // get next cursor
        count++;
      } else {
        cursor = "none"; // break 
      }
    }  catch (e) {
      Logger.log(e);
    }
  }

  // lookup ids from twitter
  var data = [];
  var count = 0;
  var chunks = chunk(users,100);
  for (i in chunks){
    count++;
    if (count > 50){
      msg = "Reaching run-time limit. Run again to get additional data. ";
      break;
    }
    ss.toast("Getting "+i*100+" to "+(parseInt(i)+1)*100+" of "+chunks.length*100);
    try {
      var o = tw_request("GET", "users/lookup.json?user_id=" + chunks[i].join()); // note using sheetname to build api request;
      for (j in o){
          data.push(o[j]);
      }
    } catch(e) {
      ss.toast("Oops something has gone wrong. Inserting existing data");
      insertData(sheet, data);
      Logger.log(e);
    }
  }
  insertData(sheet, data);
}
function insertData(sheet, data){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (data.length>0){
    ss.toast(msg+"Inserting "+data.length+" new users");
    sheet.insertRowsAfter(1, data.length);
    setRowsData(sheet, data);
  } else {
    ss.toast("All done - no new users");
  }  
}

function resetGetLotsOfFriendAndFo(){
  ScriptProperties.setProperty("cursor", "-1");
}

function getLotsOfFriendAndFo(){
  // NOTE: before using running this script clear any existing data (apart from the header row) from the target sheet 
  var sheetName = "Followers"; // enter the sheet name to update
  var friendOrFo = "followers"; // options 'friends' or 'followers
  var optScreenName = "mhawksey"; // if you are getting someone elses friends or followers enter their screen name here 

  var statusString = "";
  var your_screen_name = tw_request("GET", "account/verify_credentials.json").screen_name;
  if (typeof optScreenName != "undefined") {
    statusString = "&screen_name="+optScreenName;
    your_screen_name = optScreenName;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!ScriptProperties.getProperty("cursor")){ 
    ScriptProperties.setProperty("cursor", "-1");
  }
  if (ScriptProperties.getProperty("cursor") != "none"){ // while twitter returns data loop
    var users = [];
    try {
      var url = friendOrFo+"/ids.json?cursor=" + ScriptProperties.getProperty("cursor") + statusString;
      var o = tw_request("GET", url);
      for (j in o.ids){
        users.push(o.ids[j]);
      }
      var data = [];
      var count = 0;
      var chunks = chunk(users,100);
      for (i in chunks){
        ss.toast("Getting "+i*100+" to "+(parseInt(i)+1)*100+" of "+chunks.length*100);
        var d = tw_request("GET", "users/lookup.json?user_id=" + chunks[i].join()); // note using sheetname to build api request;
        for (j in d){
          data.push(d[j]);
        }
      }
      insertData(sheet, data);
      if (o.next_cursor!="0"){
        ScriptProperties.setProperty("cursor", o.next_cursor.toString());
        ss.toast("Run this function again to get more data"); // get next cursor
      } else {
        ScriptProperties.setProperty("cursor", "none"); // break 
        ss.toast("All data collected");
      }
    // lookup ids from twitter
    }  catch (e) {
      Logger.log(e);
    }
  }
}

function aTest(){ // quick test to see if recieving data from twitter API
  var api_request = "account/verify_credentials.json";
  var method = "GET";
  var data = tw_request(method, api_request);
  if(data){
    Browser.msgBox("Connected to Twitter sucessfully");
  } else {
    Browser.msgBox("OOPS - it didn't work");
  }
}

function clearSheet(){ // quick function to clear active sheet leaving headers
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.getRange(2, 1, sheet.getLastRow(), sheet.getMaxColumns()).clear({contentsOnly:true});
}
function authenticate(){
  authorize();
}
function configureAPI() {
  //getID();
  renderAPIConfigurationDialog();
}

function renderAPIConfigurationDialog() {
// modified from Twitter Approval Manager 
// http://code.google.com/googleapps/appsscript/articles/twitter_tutorial.html
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle(
      "Twitter API Authentication Configuration").setHeight(400).setWidth(420);
  app.setStyleAttribute("padding", "10px");

  //var dialog = app.loadComponent("GUIComponent");
  var dialogPanel = app.createFlowPanel().setWidth("400px");
  var label1 = app.createLabel("1. Register for an API key with Twitter at http://dev.twitter.com/apps/new (if you've already registered a Google Spreadsheet/Twitter mashup you can reuse your existing Consumer Key/Consumer Secret).  In the form these are the important bits: ").setStyleAttribute("paddingBottom", "10px");
  var label2 = app.createLabel(" - Application Website = anything you like").setStyleAttribute("textIndent", "30px");
  var label3 = app.createLabel(" - Application Type = Browser").setStyleAttribute("textIndent", "30px");
  var label4 = app.createLabel(" - Callback URL = https://spreadsheets.google.com/macros").setStyleAttribute("textIndent", "30px");
  var label5 = app.createLabel(" - Default Access type = Read-only ").setStyleAttribute("textIndent", "30px").setStyleAttribute("paddingBottom", "10px");
  var label6 = app.createLabel("2. Once finished filling in the form and accepting Twitter's terms and conditions you'll see a summary page which includes a Consumer Key and Consumer Secret which you need to enter below").setStyleAttribute("paddingBottom", "10px");
  var label7 = app.createLabel("3. When your Key and Secret are saved you need to open Tools > Script Editor ... and run the 'authenticate' function").setStyleAttribute("paddingBottom", "10px");
//("<strong>hello</strong><ul><li>one</li></ul>");
  dialogPanel.add(label1);
  dialogPanel.add(label2);
  dialogPanel.add(label3);
  dialogPanel.add(label4);
  dialogPanel.add(label5);
  dialogPanel.add(label6);
  dialogPanel.add(label7);

  var consumerKeyLabel = app.createLabel(
      "Twitter OAuth Consumer Key:");
  var consumerKey = app.createTextBox();
  consumerKey.setName("consumerKey");
  consumerKey.setWidth("90%");
  consumerKey.setText(getConsumerKey());
  var consumerSecretLabel = app.createLabel(
      "Twitter OAuth Consumer Secret:");
  var consumerSecret = app.createTextBox();
  consumerSecret.setName("consumerSecret");
  consumerSecret.setWidth("90%");
  consumerSecret.setText(getConsumerSecret());
  
  var saveHandler = app.createServerClickHandler("saveConfiguration");
  var saveButton = app.createButton("Save Configuration", saveHandler);
  
  var listPanel = app.createGrid(2, 2);
  listPanel.setStyleAttribute("margin-top", "10px")
  listPanel.setWidth("100%");

  listPanel.setWidget(0, 0, consumerKeyLabel);
  listPanel.setWidget(0, 1, consumerKey);
  listPanel.setWidget(1, 0, consumerSecretLabel);
  listPanel.setWidget(1, 1, consumerSecret);


  // Ensure that all form fields get sent along to the handler
  saveHandler.addCallbackElement(listPanel);
  
  //var dialogPanel = app.createFlowPanel();
  //dialogPanel.add(helpLabel);
  dialogPanel.add(listPanel);
  dialogPanel.add(saveButton);
  app.add(dialogPanel);
  doc.show(app);
}

function tw_request(method, api_request){ 
  // general purpose function to interact with twitter API
  // for method and api_request doc see http://dev.twitter.com/doc/
  // retuns object
  var oauthConfig = UrlFetchApp.addOAuthService("twitter");
  oauthConfig.setAccessTokenUrl(
      "https://api.twitter.com/oauth/access_token");
  oauthConfig.setRequestTokenUrl(
      "https://api.twitter.com/oauth/request_token");
  oauthConfig.setAuthorizationUrl(
      "https://api.twitter.com/oauth/authorize");
  oauthConfig.setConsumerKey(getConsumerKey());
  oauthConfig.setConsumerSecret(getConsumerSecret());
  var requestData = {
        "method": method,
        "oAuthServiceName": "twitter",
        "oAuthUseToken": "always"
      };
    try {
      var result = UrlFetchApp.fetch(
          "https://api.twitter.com/1.1/"+api_request,
          requestData);
      var o  = Utilities.jsonParse(result.getContentText());
      return o;
    } catch (e) {
      Logger.log(e);
    }
  
   return false;
}

// Back to the stuff from Google -->

// setRowsData fills in one row of data per object defined in the objects Array.
// For every Column, it checks if data objects define a value for it.
// Arguments:
//   - sheet: the Sheet Object where the data will be written
//   - objects: an Array of Objects, each of which contains data for a row
//   - optHeadersRange: a Range of cells where the column headers are defined. This
//     defaults to the entire first row in sheet.
//   - optFirstDataRowIndex: index of the first row where data should be written. This
//     defaults to the row immediately below the headers.
function setRowsData(sheet, objects, optHeadersRange, optFirstDataRowIndex) {
  var headersRange = optHeadersRange || sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var firstDataRowIndex = optFirstDataRowIndex || headersRange.getRowIndex() + 1;
  var headers = normalizeHeaders(headersRange.getValues()[0]);

  var data = [];
  for (var i = 0; i < objects.length; ++i) {
    var values = []
    for (j = 0; j < headers.length; ++j) {
      var header = headers[j];
      values.push(header.length > 0 && objects[i][header] ? objects[i][header] : "");
    }
    data.push(values);
  }
  var destinationRange = sheet.getRange(firstDataRowIndex, headersRange.getColumnIndex(), 
                                        objects.length, headers.length);
  destinationRange.setValues(data);
}

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    //if (!isAlnum(letter)) {
    //  continue;
    //}
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}
// http://jsfromhell.com/array/chunk
function chunk(a, s){
    for(var x, i = 0, c = -1, l = a.length, n = []; i < l; i++)
        (x = i % s) ? n[c][x] = a[i] : n[++c] = [a[i]];
    return n;
}

// The first part of this code was developed by google and has a copyright statement. 
// Everything after // Archive Twitter Status Updates is by mhawksey and released under CC


// Copyright 2010 Google Inc. All Rights Reserved.
 
/**
 * @fileoverview Google Apps Script demo application to illustrate usage of:
 *     MailApp
 *     OAuthConfig
 *     ScriptProperties
 *     Twitter Integration
 *     UiApp
 *     UrlFetchApp
 *     
 * @author vicfryzel@google.com (Vic Fryzel)
 */

/**
 * Key of ScriptProperty for Twitter consumer key.
 * @type {String}
 * @const
 */
var CONSUMER_KEY_PROPERTY_NAME = "twitterConsumerKey";

/**
 * Key of ScriptProperty for Twitter consumer secret.
 * @type {String}
 * @const
 */
var CONSUMER_SECRET_PROPERTY_NAME = "twitterConsumerSecret";

/**
 * Key of ScriptProperty for tweets and all approvers.
 * @type {String}
 * @const
 */
var TWEETS_APPROVERS_PROPERTY_NAME = "twitterTweetsWithApprovers";

/**
 * @param String Approver email address required to give approval
 *               prior to a tweet going live.  Comma-delimited.
 */
function setApprovers(approvers) {
  ScriptProperties.setProperty(APPROVERS_PROPERTY_NAME, approvers);
}

/**
 * @return String OAuth consumer key to use when tweeting.
 */
function getConsumerKey() {
  var key = ScriptProperties.getProperty(CONSUMER_KEY_PROPERTY_NAME);
  if (key == null) {
    key = "";
  }
  return key;
}

/**
 * @param String OAuth consumer key to use when tweeting.
 */
function setConsumerKey(key) {
  ScriptProperties.setProperty(CONSUMER_KEY_PROPERTY_NAME, key);
}

/**
 * @return String OAuth consumer secret to use when tweeting.
 */
function getConsumerSecret() {
  var secret = ScriptProperties.getProperty(CONSUMER_SECRET_PROPERTY_NAME);
  if (secret == null) {
    secret = "";
  }
  return secret;
}

/**
 * @param String OAuth consumer secret to use when tweeting.
 */
function setConsumerSecret(secret) {
  ScriptProperties.setProperty(CONSUMER_SECRET_PROPERTY_NAME, secret);
}


/**
 * @return bool True if all of the configuration properties are set,
 *              false if otherwise.
 */
function isConfigured() {
  return getConsumerKey() != "" && getConsumerSecret != "" ;
}


/** Retrieve config params from the UI and store them. */
function saveConfiguration(e) {
  setConsumerKey(e.parameter.consumerKey);
  setConsumerSecret(e.parameter.consumerSecret);
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

 
/**
 * Authorize against Twitter.  This method must be run prior to 
 * clicking any link in a script email.  If you click a link in an
 * email, you will get a message stating:
 * "Authorization is required to perform that action."
 */
function authorize() {
  var oauthConfig = UrlFetchApp.addOAuthService("twitter");
  oauthConfig.setAccessTokenUrl(
      "https://api.twitter.com/oauth/access_token");
  oauthConfig.setRequestTokenUrl(
      "https://api.twitter.com/oauth/request_token");
  oauthConfig.setAuthorizationUrl(
      "https://api.twitter.com/oauth/authorize");
  oauthConfig.setConsumerKey(getConsumerKey());
  oauthConfig.setConsumerSecret(getConsumerSecret());
  var requestData = {
    "method": "GET",
    "oAuthServiceName": "twitter",
    "oAuthUseToken": "always"
  };
 var result = UrlFetchApp.fetch(
      "http://api.twitter.com/1.1/account/verify_credentials.json",
      requestData);
  var o  = Utilities.jsonParse(result.getContentText());
  ScriptProperties.setProperty("STORED_SCREEN_NAME", o.screen_name);
}