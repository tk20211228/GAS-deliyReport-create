
function CSVStringToArray(strData) {
    var rows = strData.trim().split("\n");
    return rows.map(function(row) {
        var arrayDate = row.split(",");
        // 二重引用符を取り除く処理
        arrayDate = arrayDate.map(item => {
          if(item.startsWith('"') && item.endsWith('"')){
            return item.slice(1,-1);
          }
          return item;
        })
        return arrayDate;
    });
}
// Dateオブジェクトを'yyyy-MM-dd'形式の文字列に変換する関数
function formatDate(date) {
    return Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM-dd");
}

// 特定の日付のインデックスを取得するための関数
function findDateIndex(dateString, arr) {
    return arr.indexOf(dateString);
}



function uploadFileToDrive(base64Data,fileName) {
  try {
    var date = base64Data.split(",")[1];
    // console.log("csvArray",csvArray);
    // Base64データをバイト配列にデコード
    var decodedBytes = Utilities.base64Decode(date);

    // バイト配列を文字列に変換
    var csvString = Utilities.newBlob(decodedBytes).getDataAsString('Shift_JIS');

    // CSV文字列を配列に変換
    var csvArray = CSVStringToArray(csvString);

    // console.log(csvArray);

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
    // console.log(subFolderName)



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
    // const csvString = base64Data.getBlob().getDataAsString("Shift_JIS");
    // var values = Utilities.parseCsv(csvString);

    // ダイヤログ確認
    const data = new Date();
    const today = Utilities.formatDate(data, "Asia/Tokyo", "yyyy-MM-dd");

    const csvArrayTodayIndex = csvArray[0].indexOf(today);
    const todayTaskTimeList = [];


    // console.log("mainsheet");
    const mainsheet = SpreadsheetApp.getActiveSpreadsheet();
    // console.log(mainsheet);

    const mySheet = mainsheet.getSheetByName(myName[0]);
    // console.log(mySheet);

    const mySheetTaskList = mySheet.getRange("B:B").getValues().flat();
    // console.log(mySheetTaskList);


    const mySheetdayList = mySheet.getRange("5:5").getValues().flat();
    // 配列を更新して、Dateオブジェクトを'yyyy-MM-dd'形式の文字列に変換
    const mySheetdayListFormattedArray = mySheetdayList.map(item => {
        if (item instanceof Date) {
            return formatDate(item);
        }
        return item;
    });
    // console.log(today);
    // console.log(mySheetdayListFormattedArray);
    
    for(i=1;i<csvArray.length;i++){
      if(!csvArray[i][csvArrayTodayIndex]) continue;
      let taskRow = mySheetTaskList.indexOf(csvArray[i][0]) + 4 ;
      let taskCol = findDateIndex(today, mySheetdayListFormattedArray) + 1;
      todayTaskTimeList.push([csvArray[i][0],csvArray[i][csvArrayTodayIndex],taskRow,taskCol]);
      
      
    }
    console.log(todayTaskTimeList);

    


    



    // const body = `
    // <p>成功</p>
    // <p>${csvArray[1][0]}</p>
    // <p>${today}</p>
    // <p>${csvArrayTodayIndex}</p>
    // <p>${csvArray.length}</p>
    // <p>${todayTaskTimeList.toString}</p>
    // `; // 
    // // createDailog(body)

    return file.getName();
  } catch (f) {
    return f.toString();
  }
}


function csvInput() {
  let title = 'CSV入力';

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
