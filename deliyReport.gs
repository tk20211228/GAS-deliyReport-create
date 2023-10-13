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
    return "yyyy/MM/dd"; // デフォルトの日付フォーマットを返す
  }
}
function createBody(myName){
  const activeSheet = SpreadsheetApp.getActiveSheet();
  //日報出力する日付を取得
  let day = getDay(activeSheet);
  if(!day) return;

  //題名を作成する。のちほど、メールの件名として扱う
  const subject = '[MDM]【日報】'+ myName[1] + '\ ' + day;

  //進捗表から検索対象の値を取得する。
  const activeRange = activeSheet.getActiveRange();
  const selectRow = activeRange.getRow();
  const selectColumn = activeRange.getColumn();

  //全作業計画取得
  const selectAllPlanVlales = activeSheet.getRange(selectRow,2,14,415).getValues();

  //本日の作業実績のインデックス値を取得
  const todayAchievementNo = selectColumn - 2;

  //翌日の作業予定のインデックス値を取得
  const tasklistNest = [...selectAllPlanVlales[4]];
  tasklistNest.splice(0,todayAchievementNo+1);

  var selectNexstColumnNo = tasklistNest.findIndex(currentValue => currentValue > 0);

  if(selectNexstColumnNo == -1){
    const taslListPlanNest = [...selectAllPlanVlales[2]];
    taslListPlanNest.splice(0,todayAchievementNo+1);
    console.log(taslListPlanNest);
     var selectNexstColumnNo = taslListPlanNest.findIndex(currentValue => currentValue > 0);
  }
  if(selectNexstColumnNo == -1){
     var selectNexstColumnNo = 0;
  }

  const nexstdayAchievementNo = selectNexstColumnNo+todayAchievementNo+1;

  // 日報に必要な日付データをループ処理でフォーマット変換させる。
  // 開始予定,完了予定
  let dayDete = [selectAllPlanVlales[1][1],selectAllPlanVlales[1][2]];
  console.log(dayDete);
  console.log(dayDete.length);
  for(b=0;b<dayDete.length;++b){
    if(dayDete[b]){
      var deta = Utilities.formatDate(dayDete[b], "Asia/Tokyo", "yyyy-MM-dd");
      dayDete[b] = deta;
    }else{
      var deta = 'なし';
      dayDete[b] = deta;
    }
  }
  const destination = getProp('destination-sdm');

  let bodyItem = {
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

  var addBody = '';
  for( key in bodyItem ) {
    if(bodyItem[key][1] === 'メモ') continue;
    if(bodyItem[key][1] === '開始日'||bodyItem[key][1] === '完了日'){
      let body = `<label for="${key}" class="col-sm-2 col-form-label">${bodyItem[key][1]}</label><div class="col-sm-4 mb-2"><input type="date" class="form-control text-end" id="${key}" v-model="bodyItem.${key}[0]"></div>`;
      addBody += body;
    }else if(bodyItem[key][1] === '宛先'||bodyItem[key][1] === '件名'||bodyItem[key][1] === '担当者'||bodyItem[key][1] === 'タスク名'){
      let body = `<label for="${key}" class="col-sm-2 col-form-label">${bodyItem[key][1]}</label><div class="col-sm-10 mb-2"><input type="text" class="form-control" id="${key}" v-model="bodyItem.${key}[0]"></div>`;
      addBody += body;
    }else if(bodyItem[key][1] === ' 本日の作業実績 [ 実績 ] / [ 目標 ] '|| bodyItem[key][1] === ' 明日の作業予定 [ 実績 ] / [ 目標 ] '){
      let body = `<div class="d-flex justify-content-center mb-2">-----------------${bodyItem[key][1]}-----------------</div>`;
      addBody += body;
    }
    else if(bodyItem[key][1] === '累積項目数'||bodyItem[key][1] === '累積時間'){
      let body = `<label for="${key}" class="col-sm-3 col-form-label">${bodyItem[key][1]}</label><div class="col-sm-3 mb-2"><input type="number" class="form-control text-end" id="${key}" v-model="bodyItem.${key}[0]" :value="${key}" min="0"></div>`;
      addBody += body;
    }
    else{
      let body = `<label for="${key}" class="col-sm-3 col-form-label">${bodyItem[key][1]}</label><div class="col-sm-3 mb-2"><input type="number" class="form-control text-end" id="${key}" v-model="bodyItem.${key}[0]" min="0"></div>`;
      addBody += body;
      };
    };
  bodyItem['addBody'] = addBody;
  return bodyItem;
}

function createReport(){
  ///メールの内容を作成
  try{
      const myName = getMyname(); // Errorがあるとnot mypropが返る
      var bodyItem = createBody(myName);
      // console.log(bodyItem);
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