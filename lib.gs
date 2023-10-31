function createError(customErrorMessage,systemErrorMessage) {
  let title = 'エラー';
  var output = HtmlService.createTemplateFromFile('Dailog');
  output.customErrorMessage = customErrorMessage;
  output.systemErrorMessage = systemErrorMessage;
  output.inputLib = HtmlService.createHtmlOutputFromFile('cdn').getContent();
  var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(533)
    .setHeight(300);//16：9の比率に設定
  SpreadsheetApp.getUi().showModelessDialog(html, title);//showModelessDialog()でも可
}

function createDailog(body){
  var htmlOutput = HtmlService
      .createHtmlOutput(body)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(533)
      .setHeight(300);//16：9の比率に設定
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '実行結果');
  //showModelessDialog()でも可
}

function changeDayformatjp(){
  let test = SpreadsheetApp.getActiveRange().getValues();
  var value7 = '　完了見込(実績) ：'+Utilities.formatDate(new Date(test), "Asia/Tokyo", "yyyy/MM/dd");
  console.log(value7);
};

function myNameSerach(){
  try{
      var user = Session.getActiveUser();
      var contact = ContactsApp.getContact(user);
      var fullName = contact.getFullName();
      var familyName = contact.getFamilyName();
    }catch(e){
      var getName_answer = Browser.msgBox('ユーザー名を正しく取得できませんでした。\\n処理を継続しますか？', Browser.Buttons.YES_NO);
      if(getName_answer === 'yes'){
        var fullName = '（フルネーム）';
        var familyName = '（姓）';
      }else{
        const contents = '<p>ユーザー名（姓名）を登録していないと実行できません。</p>';
        const referenceLink = '<p><a href="https://docs.google.com/spreadsheets/d/SpreadsheetID/edit#gid=624880260&range=A1" target="blank">初期設定方法</a></p>';
        const errorBody = '<p>【エラー内容】</p><p></p>'+ e.message;
        const body = contents + referenceLink+ errorBody;
        createError(body);
        return;
      }
    }
  return [familyName,fullName];
}

//プロパティから名前を取得
function getProp(value){
  // 取得したメールアドレスでプロパティから名前を取得
  try{
    var scriptProperties = PropertiesService.getScriptProperties();
    var propData = scriptProperties.getProperties();
    return propData[value];
  }catch{
  }
}

//プロパティに苗字を登録
function setProp(key,value) {
  let scriptProperties = PropertiesService.getScriptProperties();
  //'example@gmail.com'は実行ユーザーのメールアドレスに置き換える
  scriptProperties.setProperties({
    [key]: value
  });
}

function allpropkill(){
  //すべてのプロパティを削除
  var prop = PropertiesService.getScriptProperties();
  prop.deleteAllProperties();
}

//自分の苗字を取得
function getMyname(){
  // 実行ユーザーのメールアドレス取得
  var userEmail = Session.getActiveUser().getEmail();
  console.log(userEmail);
  var familyName = getProp(userEmail);
  var fullName = getProp(familyName);

  if(!familyName){
    throw {
      customError: `ユーザー名が設定されていません。プロパティ設定で設定後、再度実行してください。`,
      systemError: "not myprop"
    };     
  }
  if(!fullName){
    throw {
      customError: `ユーザーの名前が設定されていません。プロパティ設定で設定後、再度実行してください。`,
      systemError: "not myprop"
    };     
  }
  
  return [familyName,fullName,userEmail];

}

function createGmail( to, subject, body ){
  GmailApp.createDraft( to, subject , body );
};

function inputPlanCells() {
  const mainsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadSheetId = mainsheet.getId();
  const choicePlace = mainsheet.getActiveRange()
  const choicePlaceRow = choicePlace.getRow();
  const choicePlaceColmmon = choicePlace.getColumn();

  const mainsheet1 = SpreadsheetApp.getActiveSheet();
  const tasklist = mainsheet1.getRange(choicePlaceRow,choicePlaceColmmon-2,1,4).getValues();

  const sheetlist = SpreadsheetApp.getActive().getSheets();
  let sheetNamelist = [];
  for(i=0;i<sheetlist.length;i++) {
    let sheetName = sheetlist[i].getSheetName();
    sheetNamelist[i]= sheetName;

  }
  const sheetJudgement = sheetNamelist.includes(tasklist[0][3]);

  if(!sheetJudgement){
    Browser.msgBox("エラー", "「担当」に一致したシート名が存在しません。\\nシートを作成した後、再実行してください。", Browser.Buttons.OK);
    return;
  }
  const mainSheet2 = mainsheet.getSheetByName(tasklist[0][3]);
  console.log(mainSheet2.getRange(1,1).getValue());
  const valuelastRow = mainSheet2.getLastRow();
  // console.log(valuelastRow);
  const lastRow = mainSheet2.getMaxRows();
  // console.log(lastRow);
  //行を追加
  mainSheet2.insertRowsAfter(lastRow, 17);
  const copySheet = mainsheet.getSheetByName("コピー元");
  const copysheetRange = copySheet.getRange("A1:PA18")
  //コピー対象のセル範囲のデータを貼り付け先のセルにコピーする
  const firstRngen = valuelastRow+1;
  const secondRngen = valuelastRow+18;
  console.log('A'+firstRngen+':PA'+secondRngen);
  copysheetRange.copyTo(mainSheet2.getRange('A'+firstRngen+':PA'+secondRngen), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  const thirdRange = valuelastRow + 6;
  mainSheet2.getRange('B'+ thirdRange).activate();
  mainSheet2.getRange('B'+ thirdRange).setValue(tasklist[0][2]);
  mainSheet2.getRange('G'+ thirdRange).setValue(tasklist[0][3]);
  const setValue = mainSheet2.getRange('L'+ (thirdRange+5) ).getValue();
  mainSheet2.getRange('J'+ thirdRange).setValue('=IMPORTRANGE('+'"'+spreadSheetId+'"' +","+'"\'全体把握用\'!"&'+ setValue + ')');
};

function inputPlanCellsNexst() {

  const mainsheet = SpreadsheetApp.getActiveSpreadsheet();

  const choicePlace = mainsheet.getActiveRange()
  const taskName = choicePlace.getValue();
  console.log(taskName);
  const choicePlaceRow = choicePlace.getRow();
  const choicePlaceColmmon = choicePlace.getColumn();
  console.log(choicePlaceRow+","+choicePlaceColmmon);
  const mainsheet1 = SpreadsheetApp.getActiveSheet();
  const tasklist = mainsheet1.getRange(choicePlaceRow,choicePlaceColmmon-2,1,4).getValues();
  // console.log(tasklist);
  const uniqueId =  tasklist[0][3] + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd/HH:mm:ss");
  // console.log(uniqueId);

  mainsheet1.getRange(choicePlaceRow,choicePlaceColmmon+12,4,1).setValues([[uniqueId],[uniqueId],[uniqueId],[uniqueId]]);
  const sheetlist = SpreadsheetApp.getActive().getSheets();
  let sheetNamelist = [];
  for(i=0;i<sheetlist.length;i++) {
    let sheetName = sheetlist[i].getSheetName();
    sheetNamelist[i]= sheetName;

  }
  // console.log(sheetNamelist);
  // console.log(tasklist[0][3]);
  const sheetJudgement = sheetNamelist.includes(tasklist[0][3]);
  console.log(sheetJudgement);
  if(!sheetJudgement){
    Browser.msgBox("エラー", "「担当」に一致したシート名が存在しません。\\nシートを作成した後、再実行してください。", Browser.Buttons.OK);
    return;
  }
  const mainSheet2 = mainsheet.getSheetByName(tasklist[0][3]);
  console.log(mainSheet2.getRange(1,1).getValue());
  const valuelastRow = mainSheet2.getLastRow();
  // console.log(valuelastRow);
  const lastRow = mainSheet2.getMaxRows();
  // console.log(lastRow);
  //行を追加
  mainSheet2.insertRowsAfter(lastRow, 17);
  const copySheet = mainsheet.getSheetByName("コピー元");
  const copysheetRange = copySheet.getRange("A1:PN18");
  //コピー対象のセル範囲のデータを貼り付け先のセルにコピーする
  const firstRngen = valuelastRow+1;
  const secondRngen = valuelastRow+18;
  // console.log('A'+firstRngen+':PA'+secondRngen);
  copysheetRange.copyTo(mainSheet2.getRange('A'+firstRngen+':PN'+secondRngen), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  const thirdRange = valuelastRow + 6;
  const uniqueIdRange = valuelastRow + 11;
  mainSheet2.getRange('B'+ thirdRange).activate();
  mainSheet2.getRange('B'+ thirdRange).setValue(tasklist[0][2]);
  mainSheet2.getRange('G'+ thirdRange).setValue(tasklist[0][3]);
  mainSheet2.getRange('L'+ uniqueIdRange).setValue(uniqueId);
  // const setValue = mainSheet2.getRange('L'+ (thirdRange+5) ).getValue();
  // mainSheet2.getRange('J'+ thirdRange).setValue('=IMPORTRANGE('+'"'+spreadSheetId+'"' +","+'"\'全体把握用\'!"&'+ setValue + ')');
};

function nowSheetName(){
  const sheetName = SpreadsheetApp.getActiveSheet().getSheetName();
  return sheetName;
  console.log(sheetName);

}

