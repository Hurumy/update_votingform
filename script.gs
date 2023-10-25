//悲しいが画像は手動で追加してください(GASでフォームに画像を添付する方法が用意されていない)

//コンテストの開催年月
const contestnumber = "20XX年XX月";

//投稿フォームのURLを更新してください
const votingformurl = "https://forms.gle/xxxxx";

//投稿フォームと連携されたスプレッドシートのIDを更新してください
const spreadsheetid = "xxx";

//GoogleドライブでアイテムがしまわれるフォルダのID
const folderid = "xxx";

//フォームをしまうフォルダのID
const contestfolder = "xxxx";

//作成した投票フォームのID
var votingformid = "xxxx";



//フォームに投稿の情報を追加する関数
function addEntry(form, title, description, comment, etc, i) {
  var dummy = form.addParagraphTextItem();
  dummy.setTitle(title[i]);
  dummy.setHelpText("概要: " + description[i] + "\n\n説明: " + comment[i] + "\n\nその他の動画リンクなど: " + etc[i]);
}

//最初にまずイベントの投票フォームを作る
function createEventForm() {
  //フォームを作る
  const form = FormApp.create('【投票フォーム】電子工作部作品コンテスト（' + contestnumber + '）');
  form.setDescription(contestnumber + "のコンテストの投票フォームです。どなたでも投票が可能です。また、投票期間中であっても投稿が可能です。投稿フォームはこちら→" + votingformurl);
  
  //ファイルをコンテスト用のフォルダに移動する
  const formfolder = DriveApp.getFolderById(contestfolder);
  const formfile = DriveApp.getFileById(form.getId());
  formfile.moveTo(formfolder);

  votingformid = form.getId();

  const spreadsheet = SpreadsheetApp.openById(spreadsheetid);
  const sheet = spreadsheet.getSheetByName("Form Responses 1");
  const sheetlen = sheet.getLastRow();
  const date = sheet.getRange("A2:A" + sheetlen);
  const title = sheet.getRange("C2:C" + sheetlen).getValues();
  const description = sheet.getRange("D2:D" + sheetlen).getValues();
  const file = sheet.getRange("E2:E" + sheetlen).getValues();
  const comment = sheet.getRange("F2:F" + sheetlen).getValues();
  const etc = sheet.getRange("G2:G" + sheetlen).getValues();

  var validation = form.addTextItem().setTitle("Intra名").setRequired(true);
  var folder = DriveApp.getFolderById(folderid);

  for (var i = 0; i < sheetlen - 1; i++) {
    addEntry(form, title, description, comment, etc, i);
  }

  var item = form.addMultipleChoiceItem().setTitle("最も良いと思う作品を1つ選び、投票してください。");
  item.setChoiceValues(title);
  item.setRequired(true);
}

//定期的にスプシの情報を拾ってきてフォームを更新する
function updateEventForm() {
  const form = FormApp.openById(votingformid);
  const spreadsheet = SpreadsheetApp.openById(spreadsheetid);

  const sheet = spreadsheet.getSheetByName("Form Responses 1");
  const sheetlen = sheet.getLastRow();
  const date = sheet.getRange("A2:A" + sheetlen);
  const title = sheet.getRange("C2:C" + sheetlen).getValues();
  const description = sheet.getRange("D2:D" + sheetlen).getValues();
  const file = sheet.getRange("E2:E" + sheetlen).getValues();
  const comment = sheet.getRange("F2:F" + sheetlen).getValues();
  const etc = sheet.getRange("G2:G" + sheetlen).getValues();

  var folder = DriveApp.getFolderById(folderid);

  const items = form.getItems();
  const elemnum = items.length - 2; //すでにフォームに載っている応募の個数
  //現在シートに載っている応募の個数はsheetlen - 1個

  //Logger.log(elemnum);

  //すでにある質問の数よりもスプレッドシートの要素の行が増えていたら応募を追加する
  if (elemnum < sheetlen - 1) {
    //シートからとってきた応募の配列のインデックスは、[0]~[sheetlen - 2]まで
    //すでにフォームに載っているのはelemnum個なので、[0]~[elemnum - 1]まではすでに追加されている
    for (var i = elemnum; i <= sheetlen - 2; i++) {
      addEntry(form, title, description, comment, etc, i);
    }

    //Logger.log(sheetlen - 1);

    //投票に選択肢を追加する
    items[elemnum + 1].asMultipleChoiceItem().setChoiceValues(title);

    //最終投票フォームを一番後ろに移動する
    //フォームにある質問は、[0]がintra名、[1]~[elemnum]が応募、[elemnum + 1]が投票
    //応募を追加した後での質問の個数は2 + sheetlen - 1個なので、一番後ろのインデックスは[sheetlen]
    form.moveItem(elemnum + 1, sheetlen);
  }
}
