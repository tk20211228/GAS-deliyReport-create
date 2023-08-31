// 共通で使う関数や変数を先に定義する
// Dateオブジェクトを'yyyy-MM-dd'形式の文字列に変換する関数
function formatDate(date) {
    return Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM-dd");
}
// 特定の日付のインデックスを取得するための関数
function findDateIndex(dateString, arr) {
    return arr.indexOf(dateString);
}
// エラーハンドリングの関数
function handleErrors(e) {
  // eがオブジェクトの場合、カスタムエラーとシステムエラーを取得する
    const customErrorMessage = e.customError || '';
    const systemErrorMessage = e.systemError || e.message || '';
    createError(customErrorMessage, systemErrorMessage);
}


// 主要な処理を行う関数
function csvCreateProgress({taskList}){
  try{
    // 前準備と変数定義
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    const myName = getMyname();
    const mySheet = mainSheet.getSheetByName(myName[0]);
    const data = new Date();
    const today = Utilities.formatDate(data, "Asia/Tokyo", "yyyy-MM-dd");
    const mySheetDayList = mySheet.getRange("5:5").getValues().flat();
    // 配列を更新して、Dateオブジェクトを'yyyy-MM-dd'形式の文字列に変換
    const formattedDayList = formatDayList(mySheetDayList);
    // 進捗の記録
    recordProgress(taskList, today, formattedDayList, mainSheet, mySheet);
  } catch (e) {
    console.log("csvCreateProgress",e);
    handleErrors(e);
  } finally {
    return taskList;
  }
}

// Dateオブジェクトを'yyyy-MM-dd'形式の文字列に変換する関数
function formatDayList(dayList) {
    return dayList.map(item => (item instanceof Date) ? formatDate(item) : item);
}
// 進捗を記録する関数
function recordProgress(taskList, today, formattedDayList, mainSheet, mySheet) {
    let taskCol = findDateIndex(today, formattedDayList) + 1;
    for(let n = 0; n < taskList.length; n++) {
        if(!taskList[n][2]) {
            copyTaskTable(mainSheet, mySheet);
            mainSheet.getRange('B6').setValue(taskList[n][0]);
            mySheet.getRange(9, taskCol).setValue(taskList[n][1]).setFontColor("#0000FF");
            taskList[n][2] = true;
        }
    }
}

function CSVStringToArray(strData) {
    var rows = strData.trim().split("\n");
    return rows.map(function(row) {
        var arrayDate = row.split(",");
        // 二重引用符を取り除く処理
        arrayDate = arrayDate.map(item => {
          if(item.startsWith('"') && item.endsWith('"')){
            return item.slice(1,-1);
          }
          let match = item.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
          if (match) {
            const year = match[1];
            const month = match[2].padStart(2, "0");
            const day = match[3].padStart(2, "0");
            return `${year}-${month}-${day}`;
          }
          return item;
        })
        return arrayDate;
    });
}

/** 配列内で値が重複してないか調べる **/
function existsSameValue(a){
  var s = new Set(a);
  return s.size != a.length;
}

function copyTaskTable(mainSheet,mySheet){
  // console.log("コピー実行");
    mySheet.insertRowsAfter(1, 18);
    mySheet.getRange('2:18').shiftRowGroupDepth(1);
    const copySheet = mainSheet.getSheetByName("コピー元");
    const copysheetRange = copySheet.getRange("A1:PN18");
    copysheetRange.copyTo(mySheet.getRange("A1:PN18"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  // console.log("コピー完了");
}

function uploadFile({content,fileName,taskList}) {
  
  try {
    // console.log(taskList);
    var base64Data = content.split(",")[1];
    // console.log("csvArray",csvArray);
    // Base64データをバイト配列にデコード
    var decodedBytes = Utilities.base64Decode(base64Data);

    // バイト配列を文字列に変換
    var csvString = Utilities.newBlob(decodedBytes).getDataAsString('Shift_JIS');

    // CSV文字列を配列に変換
    var csvArray = CSVStringToArray(csvString);


    const myName = getMyname();
    // console.log(myName)
    if(!myName[1]) {
        const body = '<p>エラー内容をご確認ください。</p><p>【エラー内容】</p><p>プロパティに名前が登録されていません。<br/>管理者にお問い合わせください</p>'
        createError(body);
        return
    }
    // ダイヤログ確認
    const data = new Date();
    const today = Utilities.formatDate(data, "Asia/Tokyo", "yyyy-MM-dd");
    const csvArrayTodayIndex = csvArray[0].indexOf(today);
    // console.log(csvArrayTodayIndex);
    // console.log(today);
    // console.log(csvArray[0]);



    if (csvArrayTodayIndex === -1) {
      throw new Error("Today's date was not found in the CSV array.");
    };
    const todayTaskTimeList = [];
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    const mySheet = mainSheet.getSheetByName(myName[0]);
    let mySheetTaskList = mySheet.getRange("B:B").getValues().flat();
    const mySheetDayList = mySheet.getRange("5:5").getValues().flat();
    // 配列を更新して、Dateオブジェクトを'yyyy-MM-dd'形式の文字列に変換
    const mySheetdayListFormattedArray = formatDayList(mySheetDayList);

      for(i=1;i<csvArray.length;i++){
        if(!csvArray[i][csvArrayTodayIndex]) continue;
        let taskRow = mySheetTaskList.indexOf(csvArray[i][0]) + 4 ;
        // console.log(taskRow);
        let taskCol = findDateIndex(today, mySheetdayListFormattedArray) + 1;
        // console.log(taskCol);
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

          for(n=0;n<taskList.length;n++){
            if(taskList[n][2] === undefined ){
              taskList[n].push(false);
              break;
            };
          };
        }else{
          for(n=0;n<taskList.length;n++){
            if(taskList[n][2] === undefined ){
              taskList[n].push(true);
              break;
            }
          };
          mySheet.getRange(taskRow,taskCol)
            .setValue(csvArray[i][csvArrayTodayIndex])
            .setFontColor("#0000FF");
        };
      }
    return taskList;
  } catch (e) {
    console.log(e);
    createError(e.message);
  }
}

function csvInput() {
  let title = 'CSV入力';
  var output = HtmlService.createTemplateFromFile('csvForm');
  output.inputLib = HtmlService.createHtmlOutputFromFile('bootstrap@5.0.2').getContent();
  output.csvType = "csvInput"; 
  var html = output.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(750)
    .setHeight(325);
  SpreadsheetApp.getUi().showModelessDialog(html, title);
}

function csvOutput() {
  let title = 'CSV出力';
  var output = HtmlService.createTemplateFromFile('csvForm');
  output.inputLib = HtmlService.createHtmlOutputFromFile('bootstrap@5.0.2').getContent();
  output.csvType = "csvOutput"; 
  var html = output.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(750)
    .setHeight(325);
  SpreadsheetApp.getUi().showModelessDialog(html, title);
}



