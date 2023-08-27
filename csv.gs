
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

/** 配列内で値が重複してないか調べる **/
function existsSameValue(a){
  var s = new Set(a);
  return s.size != a.length;
}

function copyTaskTable(mainsheet,mySheet){
  // console.log("コピー実行");
    mySheet.insertRowsAfter(1, 18);
    mySheet.getRange('2:18').shiftRowGroupDepth(1);
    const copySheet = mainsheet.getSheetByName("コピー元");
    const copysheetRange = copySheet.getRange("A1:PN18");
    copysheetRange.copyTo(mySheet.getRange("A1:PN18"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  // console.log("コピー完了");
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
    // console.log(myName)
    if(!myName[1]) {
        const body = '<p>エラー内容をご確認ください。</p><p>【エラー内容】</p><p>プロパティに名前が登録されていません。<br/>管理者にお問い合わせください</p>'
        createError(body);
        return
    }

    // var subFolderName = myName[0]; // サブフォルダ名
    // console.log(subFolderName)

    // var folders = DriveApp.getFoldersByName(dropbox);
    // フォルダがない場合はエラーをスロー
    // var folder;
    // if (folders.hasNext()) {
    //   folder = folders.next();
    // } else {
    //   throw new Error("Folder not found");
    // }

    // サブフォルダを取得
    // var subFolders = folder.getFoldersByName(subFolderName);
    // var subFolder;
    // console.log(subFolders.hasNext())
    // if (subFolders.hasNext()) {
    //   subFolder = subFolders.next();
    // } else {
    //   subFolder = folder.createFolder(myName[0]);
    // }

    // ファイルをサブフォルダに保存
    // var file = subFolder.createFile(ss);

    // 
    // CSVファイルを文字列化
    // const csvString = base64Data.getBlob().getDataAsString("Shift_JIS");
    // var values = Utilities.parseCsv(csvString);

    // ダイヤログ確認
    const data = new Date();
    const today = Utilities.formatDate(data, "Asia/Tokyo", "yyyy-MM-dd");
    // console.log(today);
    // console.log(csvArray[0]);

    const csvArrayTodayIndex = csvArray[0].indexOf(today);
    console.log(csvArrayTodayIndex);
    if (csvArrayTodayIndex === -1) {
      throw new Error("Today's date was not found in the CSV array.");
    };


    const todayTaskTimeList = [];


    // console.log("mainsheet");
    const mainsheet = SpreadsheetApp.getActiveSpreadsheet();
    // console.log(mainsheet);

    const mySheet = mainsheet.getSheetByName(myName[0]);
    // console.log(mySheet);

    let mySheetTaskList = mySheet.getRange("B:B").getValues().flat();
    // if(existsSameValue(mySheetTaskList)){
    //   Browser.msgBox("個人の進捗シートに、重複したタスク名の進捗管理表があります。");

    // }
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
      // console.log("try開始");
      for(i=1;i<csvArray.length;i++){
        if(!csvArray[i][csvArrayTodayIndex]) continue;
        let taskRow = mySheetTaskList.indexOf(csvArray[i][0]) + 4 ;
        // console.log(taskRow);
        let taskCol = findDateIndex(today, mySheetdayListFormattedArray) + 1;
        console.log(taskCol);
        if (taskCol === 0) {
          throw new Error(today+"を、進捗シートの”5:5”から見つけることができませんでした。");
        };

        todayTaskTimeList.push([
          csvArray[i][0],
          csvArray[i][csvArrayTodayIndex],
          taskRow,
          taskCol
          ]);
        if(taskRow == 3) {
          copyTaskTable(mainsheet,mySheet);
          mainsheet.getRange('B6').setValue(csvArray[i][0]);
          mySheet.getRange(9,taskCol)
            .setValue(csvArray[i][csvArrayTodayIndex])
            .setFontColor("#0000FF");
          mySheetTaskList = mySheet.getRange("B:B").getValues().flat();

        }else{
          mySheet.getRange(taskRow,taskCol)
            .setValue(csvArray[i][csvArrayTodayIndex])
            .setFontColor("#0000FF");

        };
        // console.log(taskRow,taskCol);
        // console.log(csvArray[i][0]);
        // console.log(csvArray[i][csvArrayTodayIndex]);
      }
      // console.log("ループ完了");
      // copyTaskTable(mainsheet,mySheet);
      // mainsheet.getRange('B6').setValue('123');
      // console.log(todayTaskTimeList);
    // try {

    // }catch(e){
    //   console.log(e);
    //   Browser.msgBox(e.message, Browser.Buttons.OK_CANCEL);
    // }






    // SpreadsheetApp.getActiveSpreadsheet().getSheetByName(myName[0]).getRange('2:3').shiftRowGroupDepth(1);
    // SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('2:3').shiftRowGroupDepth(1);
    // .getRange('2:3').shiftRowGroupDepth(1);
    



    // const body = `
    // <p>成功</p>
    // <p>${csvArray[1][0]}</p>
    // <p>${today}</p>
    // <p>${csvArrayTodayIndex}</p>
    // <p>${csvArray.length}</p>
    // <p>${todayTaskTimeList.toString}</p>
    // `; // 
    // // createDailog(body)

    // return file.getName();
  } catch (e) {
    // return f.toString();
    console.log(e);
    createError(e.message);
    // Browser.msgBox(e.message, Browser.Buttons.OK_CANCEL);
  }
}

function csvInput() {
  let title = 'CSV入力';
  // const myName = getMyname();
  // console.log(myName[0])
  var output = HtmlService.createTemplateFromFile('csvForm');
  output.csvType = "csvInput"; 
  // output.bodyItemJSON = JSON.stringify(bodyItem);
  // output.bodyItem = bodyItem;
  // output.inputsub = title;
  // output.inputCss = HtmlService.createHtmlOutputFromFile('css').getContent();
  // output.inputJs = HtmlService.createHtmlOutputFromFile('js').getContent();
  var html = output.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(700)
    .setHeight(290);
  SpreadsheetApp.getUi().showModelessDialog(html, title);
}

function csvOutput() {
  let title = 'CSV出力';
  var output = HtmlService.createTemplateFromFile('csvForm');
  output.csvType = "csvOutput"; 
  var html = output.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(700)
    .setHeight(290);
  SpreadsheetApp.getUi().showModelessDialog(html, title);
}



