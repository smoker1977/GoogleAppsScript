//----------------------------------------------
// Enum
//----------------------------------------------

var DateFormat = {
  TZ : "JST",
  STYLE : "yyyy/M/d HH:mm:ss"
}

var Color = {
  CHANGE:'#E8FFBD'
};

// 編集された際に色変え処理を動かすスイッチ
var chkSWITCH = {
  NAME : "SWITCH_SETTING",
  ON : "ON",
  OFF : "OFF"
}

// comment間のSeparator
var commentSeparator = "\n-------------\n";

// 設定値の保存場所
var SettingSheet = {
  name : "XXXXXXXXXXX",
  cell : "A1"
}

/**
 * 起動時の処理
 * 
 */
function onOpen(e){
  //  現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // 設定の値を取得
  var value = spreadsheet.getSheetByName(SettingSheet.name).getRange(SettingSheet.cell).getValue();
 
  //ログに出力
  Logger.log(value);

  // ドキュメントキャッシュから設定の読み込み
  var cache = CacheService.getDocumentCache();
  var ui = SpreadsheetApp.getUi();
  
  try{
    switch (value) {
      case chkSWITCH.ON:
        cache.put(chkSWITCH.NAME, chkSWITCH.ON);
        break;
      case chkSWITCH.OFF:
        cache.put(chkSWITCH.NAME, chkSWITCH.OFF);
        break;
      default:
        Browser.msgBox("設定値がありません。");
        break;
    }
  }catch(e){
    Browser.msgBox(
      "下記の内容を管理者へ連絡してください。" 
      + "\\n" 
      + "エラー内容：" 
      + e 
      + "\\n" 
      + "管理者が不明な場合は「XXXXXXXXX」へ連絡してください。"
      );
  }
}

/**
 * ステータスが編集された時に背景色を変更し、コメントを追加する
 * 
 */
function onEdit(event) {
  var sheet = event.source.getActiveSheet();
  var activeCell = sheet.getActiveCell();
  var oldVal = event.oldValue;
  var strComment = "";

  //  現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // 設定の値を取得
  var value = spreadsheet.getSheetByName(SettingSheet.name).getRange(SettingSheet.cell).getValue();

  
  //ログに出力
  Logger.log(oldVal);
  
  //----------------------------------------------
  // スイッチがONの場合のみ稼働する
  //----------------------------------------------
  try { 
    if (value == chkSWITCH.ON) {
      
      //----------------------------------------------
      // コメント作成
      //----------------------------------------------
      if (!event.oldValue) {
        oldVal = "空白";
      }else{
        oldVal = event.oldValue;
      }
      
      strComment = makeComment(
        Session.getActiveUser().getUserLoginId()
        , Utilities.formatDate( new Date(), DateFormat.TZ, DateFormat.STYLE)
        , oldVal
        , activeCell.getComment());
      
      //----------------------------------------------
      // 背景色変更
      //----------------------------------------------
      activeCell.setBackground(Color.CHANGE);
      
      //----------------------------------------------
      // コメント追加
      //----------------------------------------------
      //Browser.msgBox(strComment);
      activeCell.setComment(strComment);
    }
  }catch(e){
    Browser.msgBox(
      "下記の内容を管理者へ連絡してください。" 
      + "\\n" + "エラー内容：" 
      + e 
      + "\\n" 
      + "管理者が不明な場合は「XXXXXXXXX」へ連絡してください。"
      );
  }
  
  return;
}

/**
 * 書き込むコメントを作成する
 * 
 */
function makeComment (id, strTime, oldVal, preComment) {
  var str =  id
    + "\n"
    + "書込時間：" + strTime
    + "\n"
    + "変更前：" + oldVal
    + commentSeparator
    + preComment;
  return str;
}
