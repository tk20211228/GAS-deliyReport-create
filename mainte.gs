function myName() {
let user = Session.getActiveUser(); // スクリプトの実行者を取得
let contact = ContactsApp.getContact(user); // 実行者の連絡帳データを取得

let fullName = contact.getFullName(); // フルネームを取得
let familyName = contact.getFamilyName(); // 姓（苗字）のみ
let givenName = contact.getGivenName(); // 名（下の名前）のみ
console.log(fullName);
console.log(familyName);
console.log(givenName);

let mailAddress = user.getUserLoginId();

Browser.msgBox("性名："+ fullName + "\\n" + "性："+ familyName + "\\n" + "名："+ givenName+ "\\n" + "メールアドレス："+mailAddress);
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
