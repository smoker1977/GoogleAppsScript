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

// 初回起動 or 設定キャッシュが見つからない場合のコメント
var msgSTART = "設定が見つかりませんでしたので変更チェックをOFFにしています。"
+ "\\n"
+ "変更チェックを有効にする場合は、"
+ "\\n"
+ "アドオンメニュー → 更新したセルのチェック → チェックON を選択してください。";

var strMenu = {
  MENU : '変更チェック',
  ON : "チェックON",
  ON_FUNC : "chkOn",
  OFF : "チェックOFF",
  OFF_FUNC : "chkOff",
  CHK : "✔ "
}


/**
 * インストール時の処理
 * 
 */
function onInstall() {
  onOpen();
}

/**
 * 起動時の処理
 * 
 */
function onOpen(e){
  // ユーザキャッシュから設定の読み込み
  // 初回 or キャッシュが見つからない場合はOFFにする
  var cache = CacheService.getUserCache();
  var ui = SpreadsheetApp.getUi();
  
  if (!cache.get(chkSWITCH.NAME)) {
    Browser.msgBox(msgSTART);
    cache.put(chkSWITCH.NAME, chkSWITCH.OFF);
    makeMenu(cache.get(chkSWITCH.NAME), ui);
  }else{
    makeMenu(cache.get(chkSWITCH.NAME), ui);
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
  var cache = CacheService.getUserCache();
  
  //----------------------------------------------
  // スイッチがONの場合のみ稼働する
  //----------------------------------------------
  if (cache.get(chkSWITCH.NAME) == chkSWITCH.ON) {
    
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
  
  return;
}

/**
 * メニュー → 変更チェック → チェックOFF を選択した場合の処理
 * ユーザのキャッシュに書き込み
 */
function chkOn(){
  var cache = CacheService.getUserCache();
  var ui = SpreadsheetApp.getUi();

  cache.put(chkSWITCH.NAME, chkSWITCH.ON);
  makeMenu(chkSWITCH.ON, ui);
}

/**
 * メニュー → 変更チェック → チェックOFF を選択した場合の処理
 * ユーザのキャッシュに書き込み
 */
function chkOff(){
  var cache = CacheService.getUserCache();
  var ui = SpreadsheetApp.getUi();

  cache.put(chkSWITCH.NAME, chkSWITCH.OFF);
  makeMenu(chkSWITCH.OFF, ui);
}

/**
 * メニュー項目の作成
 * 設定してある方にチェックを入れる
 */
function makeMenu (e, ui) {
  var strOnMsg = "";
  var strOffMsg = "";
  
  switch (e) {
    case (chkSWITCH.ON) :
      strOnMsg = strMenu.CHK + strMenu.ON;
      strOffMsg = strMenu.OFF;
      break;
    case (chkSWITCH.OFF) :
      strOnMsg = strMenu.ON;
      strOffMsg = strMenu.CHK + strMenu.OFF;
      break;
    default:
      strOnMsg = strMenu.ON;
      strOffMsg = strMenu.OFF;
      break;
  }
  
  ui
//  .createMenu(strMenu.MENU)
  .createAddonMenu()
  .addItem(strOnMsg, strMenu.ON_FUNC)
  .addItem(strOffMsg, strMenu.OFF_FUNC)
  .addToUi();

}

/**
 * メニュー項目の作成
 * 設定してある方にチェックを入れる
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
