// function formatNumberToFixed(value) {
//     if (typeof value === 'number') {
//     return (value == 100) ? value : value.toFixed(1);
//   } 
//   return value ?? "0.00";
// }

// function formatDate(date) {
//   return date ? Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd") : 'なし';
// }

// //正規表現（/\n/g）を用いて文字列内のすべての改行コード（\n）を検索
// function formatMemo(memo) {
//   return memo ? memo.replace(/\n/g, "\n　") + "\n" : "";
// }


function taskBodyOpsion({activeSheet,taskRow,taskCol}){

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
    // const startDay                   = formatDate(selectAllPlanVlales[1][1]);                             // 開始日
    // const completeDay                = formatDate(selectAllPlanVlales[1][2]);                             // 完了日
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

    // const todayPlanUsingItemProgress      = formatNumberToFixed(selectAllPlanVlales[9][todayAchievementNo]* 100);//進捗率　(計画)
    // const todayActualItemProgress         = formatNumberToFixed(selectAllPlanVlales[10][todayAchievementNo]* 100);//進捗率　(実績)
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
　予定工数進捗   ：${todayPlanUsingTimeProgress}%[${todayTotalPlanUsingTime}/${planTotalTime}h]
　予定実施項目数 ：${todayPlanUsingItem}項目[${todayPlanUsingTime}h]
`;

    const todayActual = `
・${taskName}
　工数進捗       ：${todayActualTimeProgress}%[${todayTotalActualTime}h/${planTotalTime}h]／${todayPlanUsingTimeProgress}%[${todayTotalPlanUsingTime}/${planTotalTime}h]
　今日の実績     ：${todayActualItem}項目[${todayActualTime}h]／${todayPlanUsingItem}項目[${todayPlanUsingTime}h]
　${todayMemo}`;
    const tomorrowPlan = `
・${taskName}
　予定工数進捗   ：${tomorrowActualUsingTimeProgress}%[${tomorrowTotalActualTime}h/${planTotalTime}h]
　予定実施項目数 ：${tomorrowActualItem}項目[${tomorrowActualTime}h]
　${tomorrowMemo}`;

    const bodyItem = [todayPlan,todayActual,tomorrowPlan];
    return bodyItem;
}
