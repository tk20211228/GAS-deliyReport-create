function uploadFileToDrive(base64Data, fileName,data) {
  try {
    var splitBase = base64Data.split(','),
        type = splitBase[0].split(';')[0].replace('data:', ''),
        byteCharacters = Utilities.base64Decode(splitBase[1]),
        ss = Utilities.newBlob(byteCharacters, type);
    ss.setName(fileName);

    var dropbox = "redmine_月間工数"; // フォルダ名
    const myName = getMyname();
    console.log(myName)
    if(!myName[1]) {
        const body = '<p>エラー内容をご確認ください。</p><p>【エラー内容】</p><p>プロパティに名前が登録されていません。<br/>管理者にお問い合わせください</p>'
        createError(body);
        return
    }

    var subFolderName = myName[0]; // サブフォルダ名
    console.log(subFolderName)



    var folders = DriveApp.getFoldersByName(dropbox);

    // フォルダがない場合はエラーをスロー
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      throw new Error("Folder not found");
    }

    // サブフォルダを取得
    var subFolders = folder.getFoldersByName(subFolderName);
    var subFolder;
    console.log(subFolders.hasNext())
    if (subFolders.hasNext()) {
      subFolder = subFolders.next();
    } else {
      subFolder = folder.createFolder(myName[0]);
    }

    // ファイルをサブフォルダに保存
    // var file = subFolder.createFile(ss); 　

    // 
    // CSVファイルを文字列化
    const csvString = base64Data.getBlob().getDataAsString("Shift_JIS");
    var values = Utilities.parseCsv(csvString);

    // ダイヤログ確認
    const body = '<p>成功</p>' + values ;
    createDailog(body)

    return file.getName();
  } catch (f) {
    return f.toString();
  }
}


function csvInput() {
  let title = 'title';

    const myName = getMyname();
    console.log(myName[0])
  var output = HtmlService.createTemplateFromFile('uploadForm');
  // output.bodyItemJSON = JSON.stringify(bodyItem);
  // output.bodyItem = bodyItem;
  // output.inputsub = title;
  // output.inputCss = HtmlService.createHtmlOutputFromFile('css').getContent();
  // output.inputJs = HtmlService.createHtmlOutputFromFile('js').getContent();

  var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setWidth(700)
  .setHeight(50);
  SpreadsheetApp.getUi().showModelessDialog(html, title);

}
