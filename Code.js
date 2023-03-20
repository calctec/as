function doGet(e) {
  var html =  HtmlService.createTemplateFromFile('login');
   html.message = '';
  ///return html.evaluate();
   return html.evaluate().setTitle("Hr Portal").setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)

}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
function getUrl() {
 var url = ScriptApp.getService().getUrl();
 return url;
}
var z = (num, places) => String(num).padStart(places, '0');

var wss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1biLBBb35QGG0q_UKq0v8gKZuq4MjUoPucWgBWg4k8rU/edit');
var us = wss.getSheetByName("Users");
var sa = wss.getSheetByName("active");
var sv = wss.getSheetByName("view");
var s1 = wss.getSheetByName("Sheet1");
var s2 = wss.getSheetByName("Sheet2");
var s3 = wss.getSheetByName("Sheet3");

var pgi;
function doPost(e) {
    Logger.log(JSON.stringify(e));
    if(e.parameter.LoginButton == 'Login'){
    var username = e.parameter.username;
    var password = e.parameter.password;
    pgi = username;
    var vdate = vlogin(username, password);
    if(vdate == 'TRUE'){
    var html =  HtmlService.createTemplateFromFile('index');
    var nm = us.getDataRange().getValues();
    for(var i=1;i<nm.length;i++){if(username === nm[i][4]){var accv = nm[i][6];var nme = nm[i][7];}}
    html.username = username;
    html.message = nme;
    html.acc = accv;
    return html.evaluate();   
    }else{
    var html =  HtmlService.createTemplateFromFile('login');
    html.message = 'Failed to Login';
    return html.evaluate();     
    }}else if(e.parameter.LogoutButton == 'Logout'){
    louNow(e.parameter.username);
    var html =  HtmlService.createTemplateFromFile('login');
    html.message = 'Logged Out';
    return html.evaluate(); 
    }else if(e.parameter.Pcv == 'Pcv'){
    louNow(e.parameter.username);
    var html =  HtmlService.createTemplateFromFile('login');
    html.message = 'Access denied.';
    return html.evaluate(); 
    }else if(e.parameter.null == null){
    louNow(e.parameter.username);
    var html =  HtmlService.createTemplateFromFile('login');
    html.message = 'Login to continue...';
    return html.evaluate(); 
    }}

    function vlogin(username, password){
    var time = Utilities.formatDate(new Date(),"Asia/Kolkata", 'yyyy-MM-dd\'T\'HH:mm:ss');
    var clr =  sa.getLastRow();
    var vlr =  us.getLastRow();
    var flag ='';
    for(var x =1;x<=clr; x++){
    if(sa.getRange(x,3).getValue()==username&&username!=""){
    flag = 'TRUE';
    sa.getRange(x,5).setValue(time);
    }}if( flag == ''){
    for(var i = 1;i<=vlr; i++){
    if(us.getRange(i,2).getValue()=="Yes"&&us.getRange(i,5).getValue()==username&&us.getRange(i,6).getValue()==password&& username!=""){
    flag = 'TRUE';
    var v2 = us.getRange(i,4).getValue();
    var v3 = us.getRange(i,8).getValue();
    sa.insertRowAfter(clr).appendRow([time,v2,username,v3]);
    }}}if(flag == ''){
    flag = 'FALSE'; 
    }return flag;};

  function louNow(username){
  var clr =  sa.getLastRow();
  for(var b = 1; b <=clr; b++){
  if(sa.getRange(b, 3).getValue()==username){
  sa.deleteRow(b);}}};

 function autoLout(){
  var clr =  sa.getLastRow();
  for(var t = 1; t <= clr; t++){
  if(sa.getRange(t, 1).getValue() < new Date(Date.now() - sa.getRange(t,2).getValue()*60*1000)){
  sa.deleteRow(t);}}};

