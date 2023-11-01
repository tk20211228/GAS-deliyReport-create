function getDay(activeSheet) {
  const selectValue = activeSheet.getActiveRange().getValue();
  try {
    // 期待する日付が取得できるかをチェック
    const date = new Date(selectValue);
    if (isNaN(date.getDate())) throw new Error('Invalid date.');
    return Utilities.formatDate(selectValue, "Asia/Tokyo", "yyyy/MM/dd");
  } catch (error) {
    console.error(error);
    // ユーザーが処理を続行するかどうかを選択
    const userContinues = Browser.msgBox(
      '選択した日付が正しく取得できませんでした。\n処理を継続しますか？',
      Browser.Buttons.YES_NO
    );
    if (userContinues === 'no') {
      throw {
          customError: `選択した日付が正しく取得できませんでした`,
          systemError: error
      };
    }
    return "yyyy/MM/dd"; // 日付フォーマットを返す
  }
}

//全作業計画取得
function getSelectAllPlanVlales(activeSheet){
    const activeRange = activeSheet.getActiveRange();
    const selectRow = activeRange.getRow();
    return activeSheet.getRange(selectRow,2,14,415).getValues();
}
function getAchievementNos(selectAllPlanVlales, selectColumn){
    const todayAchievementNo = selectColumn - 2;
    const selectNexstColumnNo = getSelectNexstColumnNo(selectAllPlanVlales, todayAchievementNo);
    const nexstdayAchievementNo = selectNexstColumnNo + todayAchievementNo + 1;
    return [todayAchievementNo, nexstdayAchievementNo];
}
function getSelectNexstColumnNo(selectAllPlanVlales, todayAchievementNo){
    let selectNexstColumnNo = getNextDayIndex(selectAllPlanVlales[4], todayAchievementNo);
    if(selectNexstColumnNo === -1){
        selectNexstColumnNo = getNextDayIndex(selectAllPlanVlales[2], todayAchievementNo);
    }
    return selectNexstColumnNo === -1 ? 0 : selectNexstColumnNo;
}
function getNextDayIndex(taskList, todayAchievementNo){
    const tasklistNest = [...taskList].slice(todayAchievementNo + 1);
    return tasklistNest.findIndex(currentValue => currentValue > 0);
}
function formatDayDate(selectAllPlanVlales){
    return selectAllPlanVlales[1].slice(1,3).map(date => date ? Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM-dd") : 'なし');
}
function createBodyItemObject(selectAllPlanVlales, dayDete, destination, subject, myName, todayAchievementNo, nexstdayAchievementNo){
    //... bodyItem object creation logic
  let bodyItemObject = {
    destination                :[destination,'宛先'],
    subject                    :[subject,'件名'],
    familyName                 :[myName[0],'担当者'],//
    taskName                   :[selectAllPlanVlales[1][0],'タスク名'],// 

    startDay                   :[dayDete[0],'開始日'],//
    completeDay                :[dayDete[1],'完了日'],// 
    totalItems                 :[Number(selectAllPlanVlales[1][7]),'総項目数'],// 
    planTotalTime              :[Number(selectAllPlanVlales[2][7]),'予定総工数'],// 

    today                      :['',' 本日の作業実績 [ 実績 ] / [ 目標 ] '],

    todayActualItem            :[Number(selectAllPlanVlales[3][todayAchievementNo]),'消化項目'],// 実績(本日)
    todayPlanUsingItem         :[Number(selectAllPlanVlales[1][todayAchievementNo]),'予定項目数'],// 計画(本日)
    todayActualTime            :[Number(selectAllPlanVlales[4][todayAchievementNo]),'実工数'],// 実績(本日)
    todayPlanUsingTime         :[Number(selectAllPlanVlales[2][todayAchievementNo]),'予定工数'],// 計画(本日)

    todayTotalActualItem       :[Number(selectAllPlanVlales[6][todayAchievementNo]),'累積項目数'],// 実績(本日)
    todayTotalPlanUsingItem    :[Number(selectAllPlanVlales[5][todayAchievementNo]),'累積項目数'],// 計画(本日)
    todayTotalActualTime       :[Number(selectAllPlanVlales[8][todayAchievementNo]),'累積時間'],// 実績(本日)
    todayTotalPlanUsingTime    :[Number(selectAllPlanVlales[7][todayAchievementNo]),'累積時間'],// 計画(本日)
    todayMemo                  :[selectAllPlanVlales[13][todayAchievementNo],'メモ'],// 

    tomorrow                   :['',' 明日の作業予定 [ 実績 ] / [ 目標 ] '],

    tomorrowActualItem         :[Number(selectAllPlanVlales[3][nexstdayAchievementNo]),'消化項目'],// 実績(明日)
    tomorrowPlanUsingItem      :[Number(selectAllPlanVlales[1][nexstdayAchievementNo]),'予定項目数'],// 計画(明日)
    tomorrowActualTime         :[Number(selectAllPlanVlales[4][nexstdayAchievementNo]),'実工数'],// 実績(明日)
    tomorrowPlanUsingTime      :[Number(selectAllPlanVlales[2][nexstdayAchievementNo]),'予定工数'],// 計画(明日)

    tomorrowTotalActualItem    :[Number(selectAllPlanVlales[6][nexstdayAchievementNo]),'累積項目数'],// 実績(明日)
    tomorrowTotalPlanUsingItem :[Number(selectAllPlanVlales[5][nexstdayAchievementNo]),'累積項目数'],//計画 (明日)
    tomorrowTotalActualTime    :[Number(selectAllPlanVlales[8][nexstdayAchievementNo]),'累積時間'],// 実績(明日)
    tomorrowTotalPlanUsingTime :[Number(selectAllPlanVlales[7][nexstdayAchievementNo]),'累積時間'],// 計画(明日)
    tomorrowMemo               :[selectAllPlanVlales[13][nexstdayAchievementNo],'メモ'],// 
    };
    return bodyItemObject;
}
function createAddBody(bodyItemObject){
    let addBody = '';
    for(let key in bodyItemObject) {
      if(bodyItemObject[key][1] === 'メモ') continue;
      addBody += createBodyEntry(key, bodyItemObject[key]);
    }
    return addBody;
}
function createBodyEntry(key, arrayItem){
  let body;
    if(arrayItem[1] === '開始日'||arrayItem[1] === '完了日'){
      body = `<label for="${key}" class="col-sm-2 col-form-label">${arrayItem[1]}</label><div class="col-sm-4 mb-2"><input type="date" class="form-control text-end" id="${key}" v-model="bodyItem.${key}[0]"></div>`;

    }else if(arrayItem[1] === '宛先'||arrayItem[1] === '件名'||arrayItem[1] === '担当者'||arrayItem[1] === 'タスク名'){
      body = `<label for="${key}" class="col-sm-2 col-form-label">${arrayItem[1]}</label><div class="col-sm-10 mb-2"><input type="text" class="form-control" id="${key}" v-model="bodyItem.${key}[0]"></div>`;

    }else if(arrayItem[1] === ' 本日の作業実績 [ 実績 ] / [ 目標 ] '|| arrayItem[1] === ' 明日の作業予定 [ 実績 ] / [ 目標 ] '){
      body = `<div class="d-flex justify-content-center mb-2">-----------------${arrayItem[1]}-----------------</div>`;

    }
    else if(arrayItem[1] === '累積項目数'||arrayItem[1] === '累積時間'){
      body = `<label for="${key}" class="col-sm-3 col-form-label">${arrayItem[1]}</label><div class="col-sm-3 mb-2"><input type="number" class="form-control text-end" id="${key}" v-model="bodyItem.${key}[0]" :value="${key}" min="0"></div>`;

    }
    else{
      body = `<label for="${key}" class="col-sm-3 col-form-label">${arrayItem[1]}</label><div class="col-sm-3 mb-2"><input type="number" class="form-control text-end" id="${key}" v-model="bodyItem.${key}[0]" min="0"></div>`;

      };
    return body;

}
function createBodyRenew(myName){
  const activeSheet = SpreadsheetApp.getActiveSheet();
  //日報出力する日付を取得
  let day = getDay(activeSheet);
  if(!day) return;

  const selectAllPlanVlales = getSelectAllPlanVlales(activeSheet);
  const selectColumn = activeSheet.getActiveRange().getColumn();
  //本日の作業実績のインデックス値,翌日の作業計画のインデックス値を取得
  const [todayAchievementNo, nexstdayAchievementNo] = getAchievementNos(selectAllPlanVlales, selectColumn);
  //題名を作成する。のちほど、メールの件名として扱う
  const subject = '[MDM]【日報】'+ myName[1] + '\ ' + day;
  // 日報に必要な日付データをループ処理でフォーマット変換させる。
  // 開始予定,完了予定
  const dayDete = formatDayDate(selectAllPlanVlales);

  const destination = getProp('destination-sdm');

  let bodyItemObject = createBodyItemObject(selectAllPlanVlales, dayDete, destination, subject, myName, todayAchievementNo, nexstdayAchievementNo);
  bodyItemObject['addBody'] = createAddBody(bodyItemObject);
  return bodyItemObject;
}
function createReportRenew(){
  ///メールの内容を作成
  try{
      const myName = getMyname(); // Errorがあるとnot mypropが返る
      var bodyItem = createBodyRenew(myName);
      if(!bodyItem) return;

      let title = bodyItem.subject[0];
      var output = HtmlService.createTemplateFromFile('index');
      output.bodyItemJSON = JSON.stringify(bodyItem);
      output.bodyItem = bodyItem;
      output.inputsub = title;
      output.inputCss = HtmlService.createHtmlOutputFromFile('css').getContent();
      output.inputJs = HtmlService.createHtmlOutputFromFile('js').getContent();

      var html = output.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(1100)
      .setHeight(790);
      SpreadsheetApp.getUi().showModelessDialog(html, title);
    }catch(e){
      if(e.systemError === "not myprop"){
        Browser.msgBox('ユーザー名が設定されていません。\\nプロパティ設定で設定後、再度実行してください。', Browser.Buttons.YES);
        inputMyprop();
        return;
      };
      const customErrorMessage = e.customError || '';
      const systemErrorMessage = e.systemError || e.message || '';
      createError(customErrorMessage, systemErrorMessage);
    }

};
