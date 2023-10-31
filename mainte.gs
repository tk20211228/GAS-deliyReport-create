function myName() {
try{
  const myData = getMyname();
  // console.log("myData");
  let fullName = myData[1];
  let familyName =  myData[0];
  let mailAddress =  myData[2];
  Browser.msgBox("性："+ familyName + "\\n"+ "性名："+ fullName  + "\\n" + "メールアドレス："+mailAddress);

}catch(e){
  console.log(e)
  if(e.systemError === "not myprop"){
    Browser.msgBox('ユーザー名が設定されていません。\\nプロパティ設定で設定後、再度実行してください。', Browser.Buttons.YES);
    inputMyprop();
    return;
  }
  // eがオブジェクトの場合、カスタムエラーとシステムエラーを取得する
    const customErrorMessage = e.customError || '';
    const systemErrorMessage = e.systemError || e.message || '';
    createError(customErrorMessage, systemErrorMessage);
}
}

//選択範囲の位置を取得
function mygetRowcolumnActiveRange() {
//アクティブなスプレッドシートのシートを取得する
let mySheet = SpreadsheetApp.getActiveSheet();
//選択されているアクティブな範囲を取得する
let myActiveRange = mySheet.getActiveRange();
//アクティブな範囲から最初のRow:行、Column:列を取得する
let selectedRow = myActiveRange.getRow();
let selectedLastRow = myActiveRange.getLastRow();
//アクティブな範囲から最終のRow:行、Column:列を取得する
let selectedColumn = myActiveRange.getColumn();
let selectedLastColumn = myActiveRange.getLastColumn();
//スプレッドシート上でアクティブなセルをポップアップ表示
Browser.msgBox("セルの選択位置", "最初行："+selectedRow+"、最初列："+selectedColumn+"\n最終行："+selectedLastRow+"、最終列："+selectedLastColumn, Browser.Buttons.OK);
}

//getRangeで使用できる選択範囲の位置を取得
function mygetRowcolumnActiveRange0530() {
 //アクティブなスプレッドシートのシートを取得する
 let mySheet = SpreadsheetApp.getActiveSheet();
 //選択されているアクティブな範囲を取得する
 let myActiveRange = mySheet.getActiveRange();
 //アクティブな範囲から最初のRow:行、Column:列を取得する
 let selectedRow = myActiveRange.getRow();
 let selectedLastRow = myActiveRange.getLastRow();
 let selestedgetRangeRow = selectedLastRow-selectedRow+1;
 //アクティブな範囲から最終のRow:行、Column:列を取得する
 let selectedColumn = myActiveRange.getColumn();
 let selectedLastColumn = myActiveRange.getLastColumn();
 let selectedgetRangeColumn = selectedLastColumn-selectedColumn+1;

 //スプレッドシート上でアクティブなセルをポップアップ表示
 Browser.msgBox("選択範囲の位置", selectedRow+","+selectedColumn+","+selestedgetRangeRow+ "," +selectedgetRangeColumn, Browser.Buttons.OK);
}
