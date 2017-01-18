//----------------------------------------------
// Enum
//----------------------------------------------
var Color = {
  CHANGE:'#E8FFBD'
};

// 編集された際に色変え処理を動かすスイッチ
var STATUS = {
  SHEET : "シート15",
  CELL : "A32"
}

// comment間のSeparator
var commentSeparator = "\n-------------\n";


/**
 * ステータスが編集された時に背景色を変更し、コメントを追加する
 * イベントトリガー
 * 実行：onEdit
 * イベント：「スプレッドシートから」「値の変更」
 *
 * @param event events {@see https://developers.google.com/apps-script/understanding_events?hl=ja}
 */
function onEdit(event) {
  var sheet = event.source.getActiveSheet();
  var activeCell = sheet.getActiveCell();
  var oldVal = event.oldValue;
  var strComment = "";
  
  //----------------------------------------------
  // スイッチがONの場合のみ稼働する
  //----------------------------------------------
  if (event.source.getSheetByName(STATUS.SHEET).getRange(STATUS.CELL).getValue() == "on") {
    
    //----------------------------------------------
    // コメント作成
    //----------------------------------------------
    if (!event.oldValue) {
      oldVal = "空白";
    }else{
      oldVal = event.oldValue;
    }
    
    strComment =  Session.getActiveUser().getUserLoginId()
    + "\n"
    + "書込時間：" + Utilities.formatDate( new Date(), 'JST', "yyyy/M/d HH:mm:ss")
    + "\n"
    + "変更前：" + oldVal
    + commentSeparator
    + activeCell.getComment();

    //----------------------------------------------
    // 背景色変更
    //----------------------------------------------
    //Browser.msgBox(strComment);
    activeCell.setBackground(Color.CHANGE);

    //----------------------------------------------
    // コメント追加
    //----------------------------------------------
    //Browser.msgBox(strComment);
    activeCell.setComment(strComment);
    
  }
  
  return;
}
