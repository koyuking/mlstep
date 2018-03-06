//テスト用メモ
// Browser.msgBox(a,a);

//共通関数
var mailsheet = SpreadsheetApp.getActive().getSheetByName('メーリングリスト管理表');
var tempsheet = SpreadsheetApp.getActive().getSheetByName('メールテンプレート'); 
var logsheet = SpreadsheetApp.getActive().getSheetByName('ログ');
var logsheet10 = SpreadsheetApp.getActive().getSheetByName('10日前');
var logsheetlast = SpreadsheetApp.getActive().getSheetByName('最終通知'); 
var kanrisha = tempsheet.getRange("H2").getValue(); //管理者用メールアドレス設定箇所
var currentDate = new Date(); //今日の日付

// onOpenのインストール
function onInstall()
{
  onOpen();
}

//実行メニュー追加
function onOpen() {
  var aw = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries.push({name: "最終メール通知", functionName: "MailListlast"});
  aw.addMenu("アクション", menuEntries);
}

function MailListCheck() {
  var today = Utilities.formatDate(currentDate,"JST","yyyy/M/d");
  var check30day; //30日前のデフォルト日付
  var check10day; //10日前のデフォルト日付
  var checktoday; //更新期限のデフォルト日
  var check30date; //30日前の変換後の日付
  var check10date; //10日前の変換後の日付
  var checktodate; //更新期限日

  //メール用情報格納
  var owner;
  var ownername;
  var group;
  var mailadress;
  var maxdate;
  
  //メールテンプレートB列使用 
  //  var mailToB = tempsheet.getRange(4,2).getValue();
  //  var mailCCB = tempsheet.getRange(7,2).getValue();
  //  var mailBCCB = tempsheet.getRange(10,2).getValue();
  var mailSubjectB = tempsheet.getRange(13,2).getValue();
  var mailbodyB = tempsheet.getRange(16,2).getValue();
  var optionB = {from: kanrisha};
  
  //メールテンプレートD列使用 
  //  var mailToD = tempsheet.getRange(4,4).getValue();
  //  var mailCCD = tempsheet.getRange(7,4).getValue();
  //  var mailBCCD = tempsheet.getRange(10,4).getValue();
  var mailSubjectD = tempsheet.getRange(13,4).getValue();
  var mailbodyD = tempsheet.getRange(16,4).getValue();
  var optionD = {from: kanrisha,cc: kanrisha};
  
  //メールテンプレートF列使用
  var mailSubjectF = tempsheet.getRange(13,6).getValue();
  var mailbodyF = tempsheet.getRange(16,6).getValue();
  
  // 最終行から1行ずつ上の行を参照
  for (var i = mailsheet.getLastRow(); i > 6; i--) {
    check30day = mailsheet.getRange("N" + i).getValue();
    check10day = mailsheet.getRange("M" + i).getValue();
    checktoday = mailsheet.getRange("L" + i).getValue();
    
    // 判定カラムの制御
    if(check30day >= 1 || check10day >= 1 || checktoday >= 1){
      check30date = Utilities.formatDate(check30day,"JST","yyyy/M/d");
      check10date = Utilities.formatDate(check10day,"JST","yyyy/M/d");
      checktodate = Utilities.formatDate(checktoday,"JST","yyyy/M/d");
    }
    
    //メール送信処理
    if(check30date == today){
      owner = mailsheet.getRange("E" + i).getValue();
      ownername = mailsheet.getRange("F" + i).getValue();
      group = mailsheet.getRange("C" + i).getValue();
      mailadress = mailsheet.getRange("D" + i).getValue();
      maxdate = mailsheet.getRange("L" + i).getValue();
      maxdate = Utilities.formatDate(maxdate,"JST","yyyy/M/d"); 
      var mailbody = mailbodyB.replace('${"管理者"}',ownername).replace('${"グループ"}',group).replace('${"メールアドレス"}',mailadress).replace('${"利用期限"}',maxdate);
      GmailApp.sendEmail(owner,mailSubjectB,mailbody,optionB);
      logsheet.appendRow([currentDate, mailadress, owner, ownername, "30日前メール"]);
    }
    if(check10date == today){
      owner = mailsheet.getRange("E" + i).getValue();
      ownername = mailsheet.getRange("F" + i).getValue();
      group = mailsheet.getRange("C" + i).getValue();
      mailadress = mailsheet.getRange("D" + i).getValue();
      maxdate = mailsheet.getRange("L" + i).getValue();
      maxdate = Utilities.formatDate(maxdate,"JST","yyyy/M/d");
      var mailbody = mailbodyD.replace('${"管理者"}',ownername).replace('${"グループ"}',group).replace('${"メールアドレス"}',mailadress).replace('${"利用期限"}',maxdate);
      GmailApp.sendEmail(owner,mailSubjectD,mailbody,optionD);
      logsheet10.appendRow([currentDate, mailadress, owner, ownername, "10日前メール"]);
    }
    if(checktodate == today){
      group = mailsheet.getRange("C" + i).getValue();
      mailadress = mailsheet.getRange("D" + i).getValue();
      maxdate = mailsheet.getRange("L" + i).getValue();
      maxdate = Utilities.formatDate(maxdate,"JST","yyyy/M/d"); 
      var mailbody = mailbodyF.replace('${"グループ"}',group).replace('${"メールアドレス"}',mailadress).replace('${"利用期限"}',maxdate);
      GmailApp.sendEmail(kanrisha,mailSubjectF,mailbody,{from: kanrisha});
    }
    
  }
}

function MailListlast() {
  var lastrow = logsheet.getRange(1,19).getValue(); //メール配信最終行格納場所
  var sumirow = logsheet.getRange(2,19).getValue(); //済記載最終行格納場所
  
  //メールテンプレートJ列使用
  var mailSubjectJ = tempsheet.getRange(13,10).getValue();
  var mailbodyJ = tempsheet.getRange(16,10).getValue();
  
  // 最終行から1行ずつ上の行を参照
  for (var i = lastrow; i > sumirow; i--) {
    var sumi = logsheet.getRange("O" + i).getValue(); //済記載チェック
    var kanriadress = logsheet.getRange("L" + i).getValue();
    var kanriname  = logsheet.getRange("M" + i).getValue();
    var maillistname = logsheet.getRange("K" + i).getValue();
    var jyushinlast = logsheet.getRange("N" + i).getValue();
    if(jyushinlast = ""){
      jyushinlast = "受信履歴無し";
      logsheet.getRange(i,14).setValue('受信履歴無し');
    }else{
      jyushinlast = Utilities.formatDate(jyushinlast,"JST","yyyy/M/d");
    }
     if(sumi != "済"){
       var mailbody = mailbodyJ.replace('${"管理者名"}',kanriname).replace('${"メーリスアドレス"}',maillistname).replace('${"最終受信日"}',jyushinlast);
       GmailApp.sendEmail(kanriadress,mailSubjectJ,mailbody,{from: kanrisha});
       logsheetlast.appendRow([currentDate, maillistname, kanriadress, kanriname, "最終通知"]);
     }
    //メール送信したら「済」にする
    logsheet.getRange(i,15).setValue('済');
  }
}