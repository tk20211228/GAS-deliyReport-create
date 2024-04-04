function formatNumberToFixed(value) {
    if (typeof value === 'number') {
    return (value == 100) ? value : value.toFixed(1);
  } 
  return value ?? "0.00";
}

function formatDate(date) {
  return date ? Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd") : 'なし';
}

//正規表現（/\n/g）を用いて文字列内のすべての改行コード（\n）を検索
function formatMemo(memo) {
  return memo ? memo.replace(/\n/g, "\n　") + "\n" : "";
}

function taskBody({activeSheet,taskRow,taskCol}){

    //全作業計画取得
    let selectAllPlanVlales = activeSheet.getRange(taskRow,2,14,425).getValues();
    //本日の作業実績のインデックス値を取得
    let todayAchievementNo = taskCol;//マジックナンバーの”-2”を削除
    // console.log("taskRow",taskRow);
    // console.log("taskCol",taskCol);
    // console.log("todayAchievementNo",todayAchievementNo);
    //翌日の作業予定のインデックス値を取得
    const tasklistNest = [...selectAllPlanVlales[4]];
    // console.log("selectAllPlanVlales",selectAllPlanVlales);
    // console.log("tasklistNest",tasklistNest);
    tasklistNest.splice(0,todayAchievementNo+1);
    var selectNexstColumnNo = tasklistNest.findIndex(currentValue => currentValue > 0);
    // console.log("selectNexstColumnNo",selectNexstColumnNo);

    if(selectNexstColumnNo == -1){
      const taslListPlanNest = [...selectAllPlanVlales[2]];
      taslListPlanNest.splice(0,todayAchievementNo+1);
      // console.log("taslListPlanNest",taslListPlanNest);
      var selectNexstColumnNo = taslListPlanNest.findIndex(currentValue => currentValue > 0);
    }
    if(selectNexstColumnNo == -1){
      var selectNexstColumnNo = 0;
    }

    const nexstdayAchievementNo = selectNexstColumnNo+todayAchievementNo+1;

    const taskName                   = selectAllPlanVlales[1][0];                                         // タスク名
    const startDay                   = formatDate(selectAllPlanVlales[1][1]);                             // 開始日
    const completeDay                = formatDate(selectAllPlanVlales[1][2]);                             // 完了日
    const totalItems                 = formatNumberToFixed(selectAllPlanVlales[1][7]);                    // 消化予定項目数 合計
    const planTotalTime              = formatNumberToFixed(selectAllPlanVlales[2][7]);                    // 工数予定      合計

    const todayPlanUsingItem         = formatNumberToFixed(selectAllPlanVlales[1][todayAchievementNo]);   // 予定項目数  計画(本日)
    const todayPlanUsingTime         = formatNumberToFixed(selectAllPlanVlales[2][todayAchievementNo]);   // 予定工数   計画(本日)
    const todayActualItem            = formatNumberToFixed(selectAllPlanVlales[3][todayAchievementNo]);   // 消化項目   実績(本日)
    // console.log(todayActualItem);

    const todayActualTime            = formatNumberToFixed(selectAllPlanVlales[4][todayAchievementNo]);   // 実工数     実績(本日)
    const todayTotalPlanUsingItem    = formatNumberToFixed(selectAllPlanVlales[5][todayAchievementNo]);   // 累積項目数  計画(本日)
    const todayTotalActualItem       = formatNumberToFixed(selectAllPlanVlales[6][todayAchievementNo]);   // 累積項目数  実績(本日)
    const todayTotalPlanUsingTime    = formatNumberToFixed(selectAllPlanVlales[7][todayAchievementNo]);   // 累積時間    計画(本日)
    const todayTotalActualTime       = formatNumberToFixed(selectAllPlanVlales[8][todayAchievementNo]);   // 累積時間    実績(本日)

    const todayPlanUsingItemProgress      = formatNumberToFixed(selectAllPlanVlales[9][todayAchievementNo]* 100);//進捗率　(計画)
    const todayActualItemProgress         = formatNumberToFixed(selectAllPlanVlales[10][todayAchievementNo]* 100);//進捗率　(実績)
    const todayPlanUsingTimeProgress      = formatNumberToFixed(selectAllPlanVlales[11][todayAchievementNo]* 100);//工数進捗　(計画)
    const todayActualTimeProgress         = formatNumberToFixed(selectAllPlanVlales[12][todayAchievementNo]* 100);//工数進捗　(実績)
    const todayMemo                       = formatMemo(selectAllPlanVlales[13][todayAchievementNo]);      // メモ

    // const tomorrowPlanUsingItem      = Number(selectAllPlanVlales[1][nexstdayAchievementNo]);// 予定項目数  計画(明日)
    // const tomorrowPlanUsingTime      = Number(selectAllPlanVlales[2][nexstdayAchievementNo]);// 予定工数   計画(明日)
    const tomorrowActualItem         = formatNumberToFixed(selectAllPlanVlales[3][nexstdayAchievementNo]);// 消化項目    実績(明日)
    // console.log(tomorrowActualItem);
    const tomorrowActualTime         = formatNumberToFixed(selectAllPlanVlales[4][nexstdayAchievementNo]);// 実工数     実績(明日)
    // console.log(tomorrowActualTime);


    // const tomorrowTotalPlanUsingItem = Number(selectAllPlanVlales[5][nexstdayAchievementNo]);// 累積項目数  計画 (明日)
    const tomorrowTotalActualItem    = formatNumberToFixed(selectAllPlanVlales[6][nexstdayAchievementNo]);// 累積項目数  実績(明日)
    // const tomorrowTotalPlanUsingTime = Number(selectAllPlanVlales[7][nexstdayAchievementNo]);// 累積時間    計画(明日)
    const tomorrowTotalActualTime    = formatNumberToFixed(selectAllPlanVlales[8][nexstdayAchievementNo]);// 累積時間    実績(明日)

    // const tomorrowPlanUsingItemProgress = selectAllPlanVlales[9][nexstdayAchievementNo];//予定進捗率(計画)(次日)
    const tomorrowActualItemProgress       = formatNumberToFixed(selectAllPlanVlales[10][nexstdayAchievementNo]* 100);//予定進捗率(実績)(次日)
    // const tomorrowPlanUsingTimeProgress = selectAllPlanVlales[11][nexstdayAchievementNo];//予定工数進捗(計画)(次日)
    const tomorrowActualUsingTimeProgress  = formatNumberToFixed(selectAllPlanVlales[12][nexstdayAchievementNo]* 100);//予定工数進捗(実績)(次日) 
    const tomorrowMemo                     = formatMemo(selectAllPlanVlales[13][nexstdayAchievementNo]);      // メモ

    const todayPlan = `
・${taskName}
　予定進捗率     ：${todayPlanUsingItemProgress}%[${todayTotalPlanUsingItem}/${totalItems}]
　予定工数進捗   ：${todayPlanUsingTimeProgress}%[${todayTotalPlanUsingTime}/${planTotalTime}h]
　予定実施項目数 ：${todayPlanUsingItem}項目[${todayPlanUsingTime}h]
`;

    const todayActual = `
・${taskName}
　開始予定       ：${startDay}
　完了予定       ：${completeDay}
　進捗率         ：${todayActualItemProgress}%[${todayTotalActualItem}/${totalItems}]／${todayPlanUsingItemProgress}%[${todayTotalPlanUsingItem}/${totalItems}]
　工数進捗       ：${todayActualTimeProgress}%[${todayTotalActualTime}h/${planTotalTime}h]／${todayPlanUsingTimeProgress}%[${todayTotalPlanUsingTime}/${planTotalTime}h]
　今日の実績     ：${todayActualItem}項目[${todayActualTime}h]／${todayPlanUsingItem}項目[${todayPlanUsingTime}h]
　総項目数       ：${totalItems}項目
　${todayMemo}`;
    const tomorrowPlan = `
・${taskName}
　予定進捗率     ：${tomorrowActualItemProgress}%[${tomorrowTotalActualItem}/${totalItems}]
　予定工数進捗   ：${tomorrowActualUsingTimeProgress}%[${tomorrowTotalActualTime}h/${planTotalTime}h]
　予定実施項目数 ：${tomorrowActualItem}項目[${tomorrowActualTime}h]
　${tomorrowMemo}`;

    const bodyItem = [todayPlan,todayActual,tomorrowPlan];
    return bodyItem;
}

function createEmailBody({taskBodyList,myName}) {
  // console.log('taskBodyList',taskBodyList);
  let todayPlanAll = '';
  let todayActualAll = '';
  let tomorrowPlanAll = '';
  for(i=0;i<taskBodyList.length;i++){
    todayPlanAll += taskBodyList[i][0];
    todayActualAll += taskBodyList[i][1];
    tomorrowPlanAll += taskBodyList[i][2];
  }

  return body = `各位

お疲れ様です。
${myName[0]}です。
本日の日報を送付致します。
/-----------------------------------------------------------------/

①プロジェクト名
 【MDM】

/-----------------------------------------------------------------/

②本日の作業計画・・・[目標進捗]
${todayPlanAll}
/-----------------------------------------------------------------/

③本日の作業実績  [実績進捗]/[目標進捗]
${todayActualAll}
/-----------------------------------------------------------------/

④明日の作業予定   [実績進捗]/[目標進捗]
${tomorrowPlanAll}
/-----------------------------------------------------------------/

⑤問題点
・なし

/-----------------------------------------------------------------/

⑥依頼事項
・なし

/-----------------------------------------------------------------/

⑦連絡事項
・なし

/-----------------------------------------------------------------/

以上です。
よろしくお願い致します。`
}

function csvCreateBody({myName,taskList}){
    const activeSheet = SpreadsheetApp.getActiveSheet();
    const sheetName = activeSheet.getSheetName();
    if(sheetName !== myName[0]){
      // カスタムエラーオブジェクトを投げる
      throw {
          customError: `「${myName[0]}」シートで実行してください。`,
          systemError: null
      };
    }

    //日報出力する日付を取得
    const today = new Date()
    const formatDateToSlash = Utilities.formatDate(today, "Asia/Tokyo", "yyyy/MM/dd");
    const formatDateToISO = Utilities.formatDate(today, "Asia/Tokyo", "yyyy-MM-dd");
    console.log(formatDateToISO)

    //題名を作成する。のちほど、メールの件名として扱う
    const subject = '[MDM]【日報】'+ myName[1] + '\ ' + formatDateToSlash;

    const mySheetTaskList = activeSheet.getRange("B:B").getValues().flat();
    const mySheetdayList = activeSheet.getRange("5:5").getValues().flat();
    // 配列を更新して、Dateオブジェクトを'yyyy-MM-dd'形式の文字列に変換
    const mySheetdayListFormattedArray = mySheetdayList.map(item => {
        if (item instanceof Date) {
            return formatDate(item);
        }
        return item;
    });

    const taskCol = findDateIndex(formatDateToISO, mySheetdayListFormattedArray) - 1 ; //ここの”-1”は配列からスプレットシートの列を特定するため、代入している
    if (taskCol === -1) {
      throw {
          customError: `「${formatDateToISO}」を、進捗「${myName[0]}」シートの”5:5”から見つけることができませんでした。`,
          systemError: null
      };
    };
    let taskBodyList = [];
    for(i=0;i<taskList.length;i++){
        let taskRow = findDateIndex(taskList[i][0], mySheetTaskList);
        if(taskRow === -1) continue;
        // 正規表現で検出
          const pattern = /(_朝会_|_機器管理_|_その他_|_CS_)/;
          if (pattern.test(taskList[i][0])) {
            var task = taskBodyOpsion({activeSheet,taskRow,taskCol});
          } else {
            var task = taskBody({activeSheet,taskRow,taskCol});
          }        
        taskBodyList.push(task);
    }
    const reportBody = createEmailBody({taskBodyList,myName});
    const userEmail = Session.getActiveUser().getEmail();
    const destination = getProp('destination-sdm');
    return {reportBody,destination,userEmail,subject};
}

function csvCreateReport({taskList}){
  try {
    const myName = getMyname();
    var bodyItem = csvCreateBody({myName,taskList});
    let title = bodyItem.subject;
    var output = HtmlService.createTemplateFromFile('csvIndex');
    output.bodyItem = bodyItem;
    output.inputLib = HtmlService.createHtmlOutputFromFile('cdn').getContent();
    output.taskListString = JSON.stringify(taskList);  // taskListをJSON文字列に変換
    // console.log("taskList",taskList)
    var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(890)
    .setHeight(660);
    SpreadsheetApp.getUi().showModelessDialog(html, title);
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
};