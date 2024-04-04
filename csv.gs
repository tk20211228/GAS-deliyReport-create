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

// Dateオブジェクトを'yyyy-MM-dd'形式の文字列に変換する関数
function formatDayList(dayList) {
    return dayList.map(item => (item instanceof Date) ? formatDate(item) : item);
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

function parseCSVRow(row) {
    const cleanItem = item => {
      // 二重引用符を取り除く処理
      if (item.startsWith('"') && item.endsWith('"')) {
          return item.slice(1, -1);
      }
      // \d{1,2}は、\d（任意の数字）が1回から2回まで現れるパターンにマッチします。例えば、「1」、「12」などが該当しますが、「123」は該当しません（「123」の最初の2文字は該当する）
      const match = item.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
      if (match) {
        const year = match[1];
        // メソッドは、指定された長さになるまで文字列の先頭を特定の文字で埋める。このメソッドは2つの引数を取る。目標の長さ：この例では 2　埋める文字：この例では "0"
        const month = match[2].padStart(2, "0");
        const day = match[3].padStart(2, "0");
        return `${year}-${month}-${day}`;
      }
      return item;
    }
    //row.split(","): 文字列 row を , で分割して配列を生成
    //.map(cleanItem): 上記で生成した配列の各要素に cleanItem 関数を適用
    return row.split(",").map(cleanItem);
}

function CSVStringToArray(strData) {
  //strData.trim(): 文字列 strData の先頭および末尾の空白を取り除く
  //.split("\n"): トリムされた文字列を \n（改行）で分割して配列を生成。これにより、多行の文字データを行ごとの配列に変換
  //.map(parseCSVRow): 上記で生成した配列の各要素（各行）に parseCSVRow 関数を適用
    return strData.trim().split("\n").map(parseCSVRow);
}

function copyTaskTable(mainSheet,mySheet){
    mySheet.insertRowsAfter(1, 18);
    mySheet.getRange('2:18').shiftRowGroupDepth(1);
    mainSheet.getSheetByName("コピー元").getRange("A1:PN18")
        .copyTo(mySheet.getRange("A1:PN18"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false)
    mySheet.getRange("B6").activate();
}

function uploadFile({content,taskList}) {
    try {
        const base64Data = content.split(",")[1];
        // Base64データをバイト配列にデコード
        const decodedBytes = Utilities.base64Decode(base64Data);
        // バイト配列を文字列に変換
        const csvString = Utilities.newBlob(decodedBytes).getDataAsString('Shift_JIS');
        // CSV文字列を配列に変換
        const csvArray = CSVStringToArray(csvString);


        const today = formatDate(new Date());
        const myName = getMyname();
        const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
        const mySheet = mainSheet.getSheetByName(myName[0]);
        const mySheetdayListFormattedArray = formatDayList(mySheet.getRange("5:5").getValues().flat());
        const res = validateUploadData(mySheet, csvArray, today, mySheetdayListFormattedArray, taskList);
        return res

    } catch (e) {
        handleErrors(e);
    }
}
function validateUploadData(mySheet, csvArray, today, mySheetdayListFormattedArray, taskList) {
    const todayIndex = csvArray[0].indexOf(today);

    // 変数todayIndexで、配列内に本日の日付がない場合、-1で返される
    if (todayIndex === -1) {
      throw {
          customError: `読み込まれたCSVファイルを読み込みましたが、
          本日[${today}]が入力された値を見つけることができませんでした。
          ファイルをご確認ください。`,
          systemError: null
      };
    }

    // 配列の最初の行（0番目の行、おそらくヘッダ行）を除いて残りのデータ行だけを対象
    const taskCol = findDateIndex(today, mySheetdayListFormattedArray) + 1;
    const mySheetTaskList = mySheet.getRange("B:B").getValues().flat();
    csvArray.slice(1).forEach(row => {
        if (!row[todayIndex]) return;
        const taskRow = mySheetTaskList.indexOf(row[0]) + 4;
        if (taskCol === 0) {
          throw new Error(today+"を、進捗シートの”5:5”から見つけることができませんでした。");
        };
        
        // 変数taskRowで該当のタスク名がない場合、-1+4＝3 となる。
        if (taskRow === 3) {
            taskList.find(task => task[2] === undefined)[2] = false;
        } else {
            taskList.find(task => task[2] === undefined)[2] = true;
            mySheet.getRange(taskRow,taskCol).setValue(row[todayIndex]).setFontColor("#0000FF");
        }
    });
    return taskList;
}

function csvInput() {
  let title = 'CSV入力';
  var output = HtmlService.createTemplateFromFile('csvForm');
  output.inputLib = HtmlService.createHtmlOutputFromFile('cdn').getContent();
  var html = output.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(750)
    .setHeight(325);
  SpreadsheetApp.getUi().showModelessDialog(html, title);
}



