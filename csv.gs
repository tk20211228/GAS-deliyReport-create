function csvCreateProgress({taskList}){
  try{
    const mainsheet = SpreadsheetApp.getActiveSpreadsheet();
    const myName = getMyname();
    const mySheet = mainsheet.getSheetByName(myName[0]);
    const data = new Date();
    const today = Utilities.formatDate(data, "Asia/Tokyo", "yyyy-MM-dd");
    const mySheetdayList = mySheet.getRange("5:5").getValues().flat();
    // 配列を更新して、Dateオブジェクトを'yyyy-MM-dd'形式の文字列に変換
    const mySheetdayListFormattedArray = mySheetdayList.map(item => {
        if (item instanceof Date) {
            return formatDate(item);
        }
        return item;
    });
    let taskCol = findDateIndex(today, mySheetdayListFormattedArray) + 1;
    for(n=0;n<taskList.length;n++){
      if(taskList[n][2] === false){
        copyTaskTable(mainsheet,mySheet);
        mainsheet.getRange('B6').setValue(taskList[n][0]);
        mySheet.getRange(9,taskCol)
          .setValue(taskList[n][1])
          .setFontColor("#0000FF");
        taskList[n][2] = true ;
      };
    };
  }catch(e){
    console.log("csvCreateBody123",e);
    // eがオブジェクトの場合、カスタムエラーとシステムエラーを取得する
    const customErrorMessage = e.customError || '';
    const systemErrorMessage = e.systemError || e.message || '';
    createError(customErrorMessage, systemErrorMessage);
  }finally{
    return taskList;

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

function uploadFile({content,fileName,taskList}) {
  
  try {
    console.log(taskList);
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
    if (csvArrayTodayIndex === -1) {
      throw new Error("Today's date was not found in the CSV array.");
    };
    const todayTaskTimeList = [];
    const mainsheet = SpreadsheetApp.getActiveSpreadsheet();
    const mySheet = mainsheet.getSheetByName(myName[0]);
    let mySheetTaskList = mySheet.getRange("B:B").getValues().flat();
    const mySheetdayList = mySheet.getRange("5:5").getValues().flat();
    // 配列を更新して、Dateオブジェクトを'yyyy-MM-dd'形式の文字列に変換
    const mySheetdayListFormattedArray = mySheetdayList.map(item => {
        if (item instanceof Date) {
            return formatDate(item);
        }
        return item;
    });

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



