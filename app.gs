// LINE Messenging API Token
var CHANNEL_ACCESS_TOKEN = 'INPUT YOUR CHANNEL ACCESS TOKEN';

var ban_arr = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十', '十一'];
var date = new Date();
var kick_message = ["已被無情踢除", "我們懷念他"];

var intro = "指令說明:\n"+
"1. 輸入\"建立名單xxxxx到xxxxx\"建立回報成員名單(xxxxx為學號)\n"+
"2. 輸入\"重新建立名單\"清空回報成員名單\n"+
"3. 輸入\"個別新增人員xxxxx\"可以個別新增人員至名單中\n"+
"4. 輸入回報信息時，格式須包含\"時間:、學號:、姓名:、電話:\"等\n"+
"5. 輸入\"回報\"，可統整所有人的回報訊息，統整完重置訊息。若有人未回報，則顯示未回報人員學號\n"+
"6. 輸入\"重新回報\"，清空已回報訊息\n"+
"7. 輸入\"踢xxxxx\",即可將其從名單移除\n"+
"8. 輸入\"help\"或\"幫助\"，顯示此頁說明\n"+
"9. 輸入不符格式或無相關內容，則不回應";

function doPost(e) {
  
  // 從接收到的訊息中取出 replyToken 和發送的訊息文字
  var msg = JSON.parse(e.postData.contents);
  var replyToken = msg.events[0].replyToken;
  var userMessage = msg.events[0].message.text;
  var event_type = msg.events[0].source.type;
  
  const user_id = msg.events[0].source.userId;
  try {
    var group_id = msg.events[0].source.groupId;
  }
  catch{
    console.log("wrong");
  }
  
  // Google表單設定
  var sheet_url = 'INPUT YOUR SHEET URL';  
  var SpreadSheet = SpreadsheetApp.openByUrl(sheet_url);

  var id_list = SpreadSheet.getSheetByName("IdtoSheet");
  var id_current_list_row = id_list.getLastRow();
  var sheet_name ="";

  if(group_id != undefined){
    for(let i=2; i<=id_current_list_row+1; i++){
      if(id_list.getRange(i,1).getValue() == group_id){
        sheet_name = id_list.getRange(i,2).getValue();
        break;
      }
    }
    if(sheet_name == ""){
      id_list.getRange(id_current_list_row+1, 1).setValue(group_id);
      id_list.getRange(id_current_list_row+1, 2).setValue("group" + id_current_list_row);
      sheet_name = id_list.getRange(id_current_list_row+1, 2).getValue();
    }
  }else{
    for(let i=2; i<=id_current_list_row+1; i++){
      if(id_list.getRange(i,1).getValue() == user_id){
        sheet_name = id_list.getRange(i,2).getValue();
        break;
      }
    }
    if(sheet_name == ""){
      id_list.getRange(id_current_list_row+1, 1).setValue(user_id);
      id_list.getRange(id_current_list_row+1, 2).setValue("user" + id_current_list_row);
      sheet_name = id_list.getRange(id_current_list_row+1, 2).getValue();
    }
  }

  var report_list = SpreadSheet.getSheetByName(sheet_name);
  if(!report_list){
    SpreadSheet.insertSheet(sheet_name);
    report_list = SpreadSheet.getSheetByName(sheet_name);
    report_list.getRange(1,1).setValue("index");
    report_list.getRange(1,2).setValue("number");
    report_list.getRange(1,3).setValue("name");
    report_list.getRange(1,4).setValue("detail");
  }

  var current_list_row = report_list.getLastRow();

  //取得使用者名稱
  switch (event_type) {
    case "user":
      var nameurl = "https://api.line.me/v2/bot/profile/" + user_id;
      break;
    case "group":
      var nameurl = "https://api.line.me/v2/bot/group/" + group_id + "/member/" + user_id;
      break;
  }
  var response = UrlFetchApp.fetch(nameurl, { //使用line API
    "method": "GET",
    "headers": {
      "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN,
      "Content-Type": "application/json"
    },
  });
  var namedata = JSON.parse(response);
  var userName = namedata.displayName; //得到使用者名稱
  var userNumber = userName.substr(0,5);
  var userRealName = userName.substr(5).trim().replace('-', "");    
  
  var reply_message=[];
  var memberCount;

  //1. 輸入\"建立名單xxxxx到xxxxx\"，建立回報成員名單(xxxxx為學號)
  if(userMessage.substr(0,4) == "建立名單"){
    var startIndex = userMessage.substr(4).split("到")[0];
    var endIndex = userMessage.substr(4).split("到")[1];
    memberCount = parseInt(endIndex - startIndex) +1;
    // console.log(typeof memberCount);

    if(isNaN(memberCount)){
      reply_message = [{
      "type": "text",
      "text": "輸入錯誤"
      }]
    }
    else if(!report_list.getRange(2,2).isBlank()){
      reply_message = [{
      "type": "text",
      "text": "名單已有成員!請\"重新建立名單\""
      }]
    }
    else if(memberCount > 50){
      reply_message = [{
      "type": "text",
      "text": "哩很促咪哦~"
      }]
    }
    else{
      for(let i=1 ; i<=memberCount; i++){
      report_list.getRange(i+1, 1).setValue(i);
      report_list.getRange(i+1, 2).setValue(parseInt(startIndex,10)+i-1);
      }
      reply_message = [{
        "type": "text",
        "text": "建立名單:"+ startIndex + "到" + endIndex
      }]
    }
  }

  //2. 輸入\"重新建立名單\"清空回報成員名單
  else if(userMessage == "重新建立名單"){
    report_list.deleteRows(2, current_list_row-1);
    report_list.deleteColumn(5); //班級資訊也重置
    reply_message = [{
      "type": "text",
      "text": "名單已清空"
    }]
  }

  //3. 輸入"個別新增人員xxxxx"可以個別新增人員至名單中
  else if(userMessage.substr(0,6) == "個別新增人員"){
    var idx = userMessage.substr(6);
    if(isNaN(idx)){
      reply_message = [{
      "type": "text",
      "text": "輸入錯誤"
      }]
    }else if(idx >=100000 || idx<11001){
      reply_message = [{
      "type": "text",
      "text": "哩很促咪哦~"
      }]
    }
    else if(report_list.getRange(2,2).isBlank()){
      report_list.getRange(2, 2).setValue(idx);
      reply_message = [{
      "type": "text",
      "text": "新增人員: "+ idx 
      }]
    }
    else{
      for(let i=2; i<=current_list_row; i++){
        if(idx < report_list.getRange(i,2).getValue()){
          report_list.insertRows(i,1);
          report_list.getRange(i, 2).setValue(parseInt(idx));
          break;
        }
        else if(idx > report_list.getRange(i,2).getValue() && i == current_list_row){
          report_list.getRange(i+1 ,2).setValue(parseInt(idx));
          break;
        }
        else if(idx > report_list.getRange(i,2).getValue()){
          continue;
        }else break;
      }
      reply_message = [{
      "type": "text",
      "text": "新增人員: "+ idx 
      }]
    }
  }

  //4. 輸入回報信息時，格式須包含\"時間:、學號:、姓名:、電話:\"等，系統會將訊息紀錄
  else if((userMessage.includes("時間:") || userMessage.includes("時間：")) &&
          (userMessage.includes("學號:") || userMessage.includes("學號：")) &&
          (userMessage.includes("姓名:") || userMessage.includes("姓名：")) &&
          (userMessage.includes("電話:") || userMessage.includes("電話：")) &&
          (userMessage.includes("現在位置:") || userMessage.includes("現在位置：")) &&
          (userMessage.includes("現在在幹嘛:") || userMessage.includes("現在在幹嘛：")) &&
          (userMessage.includes("跟誰:") || userMessage.includes("跟誰：")) &&
          (userMessage.includes("身體狀況:") || userMessage.includes("身體狀況：")) &&
          (userMessage.includes(userNumber) || userMessage.includes(userNumber))
        ){

    userMessage = userMessage.replace(/^\n+/, "");
    userMessage = userMessage.replace(/\n+$/, "");
    for(let i=2; i<=current_list_row; i++){
      if(report_list.getRange(i, 2).getValue() == userNumber){
        report_list.getRange(i, 3).setValue(userRealName);
        report_list.getRange(i, 4).setValue(userMessage);
        break;
      }
    }
  }

  //5. 輸入\"回報\"，可統整所有人的回報訊息，統整完重置訊息。若有人未回報，則顯示未回報人員學號
  else if(userMessage == "回報") {
    var allreport = "";
    var not_reported = "未回報人員:\n";
    var check = true;
    
    for(let j=2; j<=current_list_row; j++){
      if(report_list.getRange(j,4).isBlank()){
        not_reported += report_list.getRange(j,2).getValue() + "、";
        check = false;
      }
    }
    not_reported = not_reported.slice(0,-1);

    //大家都回報了，印出總回報, 並清空回報
    var hours = (date.getUTCHours()+8)%24;
    var ban = report_list.getRange(1,5).getValue();
    if(check){
      if(hours<18 && hours>=9 && ban!= ""){
        allreport += "第"+ ban +"班 15:00回報\n\n";
      }else if(hours>=18 && hours<24  && ban!= ""){
        allreport += "第"+ ban +"班 20:00回報\n\n";
      }
      for(let i=2; i<=current_list_row; i++){
        allreport += report_list.getRange(i,4).getValue() + "\n\n"; 
      }
      allreport = allreport.replace(/\n\n$/, "");

      reply_message = [
      {
        "type": "text",
        "text": allreport
      }]
      report_list.deleteColumn(4);
      var ban = report_list.getRange(1,4).getValue();//防止刪到google表單的班級資訊
      report_list.getRange(1,5).setValue(ban);
      report_list.getRange(1,4).setValue("detail");
    } 
    //還有人未回報，列出未回報人員
    else {
      reply_message=[{
        "type": "text",
        "text": not_reported
      }]
    }
  }

  //6. 輸入\"重新回報\"，清空已回報訊息
  else if(userMessage == "重新回報"){
    report_list.deleteColumn(4);
    var ban = report_list.getRange(1,4).getValue();//防止刪到google表單的班級資訊
    report_list.getRange(1,5).setValue(ban);
    report_list.getRange(1,4).setValue("detail");

    reply_message = [{
      "type": "text",
      "text": "回報資訊已清空"
    }]
  }

  //7. 輸入\"踢xxxxx\",即可踢人
  else if(userMessage.substr(0,1) == "踢"){
    var idx = userMessage.substr(1);
    if(isNaN(idx)){
      reply_message = [{
      "type": "text",
      "text": "輸入錯誤"
      }]
    }else if(idx >=100000 || idx<11001){
      reply_message = [{
      "type": "text",
      "text": "哩很促咪哦~"
      }]
    }else{
      for(let i=2; i<=current_list_row; i++){
        if(idx == report_list.getRange(i,2).getValue()){
          report_list.deleteRow(i);
          reply_message = [{
            "type": "text",
            "text": idx + kick_message[Math.floor(Math.random()*2)]
          }]
          break;
        }
      }
    }
  }

  //8. 輸入\"help\"，顯示此頁說明
  else if(userMessage == "help" || userMessage=="幫助" || userMessage == "Help"){
    reply_message = [{
      "type": "text",
      "text": intro
    }];
  }

  else if(userMessage == "他媽的回報"){
    reply_message = [{
      "type": "text",
      "text": "不要勒~%^&*"
    }];
  }

  //(option)可記錄班級資訊，回報時依時間加入前置回報訊息
  else if(userMessage.substr(0,4) == "輸入班級"){
    var ban = userMessage.substr(4);
    if(isNaN(ban) && !ban_arr.includes(ban)){
      reply_message = [{
      "type": "text",
      "text": "輸入錯誤"
      }]
    }else if(ban >=12 || ban <1){
      reply_message = [{
      "type": "text",
      "text": "哩很促咪哦~"
      }]
    }else{
      report_list.getRange(1,5).setValue(ban);
    }
  }

  //9. 輸入不符格式或無相關內容，則不回應
  else {
    reply_message = [];
  }

  //回傳訊息給line及使用者的實際作用區
  var url = 'https://api.line.me/v2/bot/message/reply';
  UrlFetchApp.fetch(url, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': replyToken,
      'messages': reply_message,
    }),
  });

}