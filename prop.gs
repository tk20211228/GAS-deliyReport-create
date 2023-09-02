//参考URL：https://officeforest.org/wp/2021/01/12/google-apps-script%E3%81%AE%E9%96%8B%E7%99%BA%E7%94%BB%E9%9D%A2%E3%81%8C%E6%96%B0%E3%81%97%E3%81%8F%E3%81%AA%E3%82%8A%E3%81%BE%E3%81%97%E3%81%9F/#i-22
//https://vuetifyjs.com/ja/components/simple-tables/
//スクリプトプロパティを一括取得してダイアログで返す
function openCheck() {
  var html = HtmlService.createHtmlOutputFromFile('proplist')
    // .setSandboxMode(HtmlService.SandboxMode.IFRAME)　※現在は、不要　https://developers.google.com/apps-script/reference/html/html-output#setsandboxmodemode
    .setWidth(600)
    .setHeight(360);// createHtmlOutputFromFile 静的なHTMLを出力する
  SpreadsheetApp.getUi().showModelessDialog(html, '設定済みプロパティ一覧'); // showModalDialogでは非活性となる
}
function inputMyprop() {
  const userEmail = Session.getActiveUser().getEmail();
  // const myName = getMyname();
  const myFirstName = getProp(userEmail);
  // console.log(myFirstName)
  let myLastName = getProp(myFirstName);
  if(!myLastName){
    myLastName = "";
  }else{
    myLastName = getProp(myFirstName).replace(myFirstName,"")
  }
  const myProp ={myFirstName,myLastName,userEmail};
  var output = HtmlService.createTemplateFromFile('propmydata')
    output.myProp = myProp // createTemplateFromFileは動的な内容（変数や条件文などサーバーサイドのJavaScriptコード）をHTMLに埋め込むことができる
  var html = output.evaluate()
    .setWidth(600)
    .setHeight(360);
  SpreadsheetApp.getUi().showModelessDialog(html, 'プロパティ 設定'); // showModalDialogでは非活性となる
}



//プロパティをすべて削除する
function clearprop(){
  try{
    //スクリプトプロパティを削除する
    const prop = PropertiesService.getScriptProperties();
    prop.deleteAllProperties();

    //ユーザプロパティを削除する
    const prop2 = PropertiesService.getUserProperties();
    prop2.deleteAllProperties();

    return "OK";
  }catch(e){
    return "NG";
  }
}

//現在設定されてるスクリプトプロパティ一覧を取得する
function onProp(){
  //プロパティ値を一括取得
  const prop = PropertiesService.getScriptProperties();
  const data = prop.getProperties();
  console.log(data);

  //vue用にデータを加工する
  var propdata = [];
  for (var key in data) {
    //入れ物用意
    let temprop = {};

    //値をtempropに入れる
    temprop.propname = key;
    temprop.propvalue = data[key];
    temprop.proptype = "スクリプト";

    //配列にpushする
    propdata.push(temprop);

  }

  const prop2 = PropertiesService.getUserProperties();
  const user = prop2.getProperties();

  for (var key in user) {
    //入れ物用意
    let temprop = {};

    //値をtempropに入れる
    temprop.propname = key;
    temprop.propvalue = user[key];
    temprop.proptype = "ユーザ";

    //配列にpushする
    propdata.push(temprop);

  }
  
  //HTML側へ返す
  return JSON.stringify(propdata);

}

//プロパティを一個削除する
function deleteman(item){
  //proptypeで処理を分岐
  var prop;
  switch(item.proptype){
    case "スクリプト":
      prop = PropertiesService.getScriptProperties();
      break;
    case "ユーザ":
      prop = PropertiesService.getUserProperties();
      break;
  }
    
  //プロパティを削除する
  var key = item.propname;
  var ret = prop.deleteProperty(key);

  return item.propname + "を削除しました。";
}

//プロパティの新規追加
function oninsert(recman){
  //プロパティタイプを取得
  var ptype = recman.proptype;
  var prop;

  //プロパティタイプによって処理を分岐
  if(ptype == 0){
    prop = PropertiesService.getScriptProperties();
  }else{
    prop = PropertiesService.getUserProperties();
  }

  //プロパティをセットする
  try{
    prop.setProperty(recman.propname,recman.propvalue);
    return "OK";
  }catch(e){
    return "NG";
  }
}

// function setMyprop(){
function setMyprop({firstName,lastName,email}){
  console.log(firstName,lastName,email);
  //プロパティをセットする
  try{
    const fullName = firstName+lastName;
    const propType =  PropertiesService.getScriptProperties();
    propType.setProperty(firstName,fullName);
    propType.setProperty(email,firstName);
    return true;
  }catch(e){
    return e.message;
  }

}



