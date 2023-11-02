function onOpen(){
    var meinUI = SpreadsheetApp.getUi();
      meinUI
        .createMenu('日報')
          //  .addItem('日報作成', 'createReport')
           .addItem('日報作成', 'createReportRenew')
        .addToUi();
      meinUI
        .createMenu('進捗入力')
        .addItem('全体把握から進捗を作成', 'inputPlanCellsNexst')//lib.gsで管理
        .addToUi();
      meinUI
        .createMenu('CSV入力/出力')
          .addItem('CSV入力/出力 Ver1.0', 'csvInput')//csv.gsで管理
          .addToUi();
      meinUI
        .createMenu('メンテナンス') 
        .addItem('ユーザー名確認', 'myName')//メンテナンス.gsで管理
        .addItem('選択範囲の位置を取得', 'mygetRowcolumnActiveRange')//メンテナンス.gsで管理
        .addItem('getRangeで使用できる選択範囲の位置','mygetRowcolumnActiveRange0530')//メンテナンス.gsで管理//
        .addItem('プロパティ確認','openCheck')//propで管理
        .addItem('ユーザー設定','inputMyprop')//propで管理
        .addToUi();
};