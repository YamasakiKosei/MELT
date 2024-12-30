//******************************************************  注意  **************************************************************
// データベースのシートが順番に並んでいるか確認
// データベースのIDを変更したとき、QUERY関数で一度「アクセスを許可」をする
// . 空白 : の全角/半角に注意 
let maintenance_flag = false; // メンテナンス中 = true

// Google Documentでないとエラー
// フォルダ内が空だとエラー
// パソコンでは、表示されない
// 1回で送れるのは、5つまで
// フォルダ名・ファイル名は、変更禁止
// 画像の拡張子はjpg
// 問題図：3600101.jpg、3610102.jpg
// 解説図：3600111.jpg、3610112.jpg
// docファイルに挿入するには、pngに変換する
// GIFは静止画になる
//****************************************************************************************************************************



/////////////////////////////////////////////////////////　　変数　　///////////////////////////////////////////////////////////

// データベースのIDを変更したとき、QUERY関数で一度「アクセスを許可」をする
let database_ss = SpreadsheetApp.openById('*****');        // データベース
let active_ss = SpreadsheetApp.openById('*****');          // NiAS-NECE 臨床工学技士国家試験
let message_sheet = active_ss.getSheetByName("Message");   // Messageシート 取得
let account_sheet = active_ss.getSheetByName("Account");   // Accountシート 取得
let q_imgFolder = DriveApp.getFolderById('*****');         // 臨床工学技士国家試験＞問題図
let a_imgFolder = DriveApp.getFolderById('*****');         // 臨床工学技士国家試験＞解説図
let outputFolder = DriveApp.getFolderById('*****');        // 臨床工学技士国家試験＞出力ファイル
let home_doc1 = DocumentApp.openById('*****');             // ホーム1のdocファイルのID
let home_doc2 = DocumentApp.openById('*****');             // ホーム2のdocファイルのID
let info_doc = DocumentApp.openById('*****');              // お知らせのdocファイルのID
let original_sheet = active_ss.getSheetByName('Original'); // ユーザー別シートのオリジナル

// LINE Developers ＞ Messaging API > Channel access token
const LINE_TOKEN = '*****';
const LINE_URL = 'https://api.line.me/v2/bot/message/reply';
const url_richmenu = 'https://api.line.me/v2/bot/richmenu';
const url_richmenu_data = 'https://api-data.line.me/v2/bot/richmenu/';
const url_user = 'https://api.line.me/v2/bot/user';
let event = null;          // イベント
let event_type = null;     // イベントタイプ　取得
let replyToken = null;     // 応答用Tokenを取得
let user_id = null;        // ユーザーID　取得
let user_name = null;      // ユーザー名　取得
let messageId = null;      // メッセージID 取得
let LINE_END_POINT = null; // LINE_END_POINT 取得
let userMessage = null;    // メッセージ 取得
let user_sheet = null;     // ユーザー別シート
// let db_url = database_ss.getUrl();

const messages = [];
const errorMessages = [];

let field_dic = {
  "医学概論": 1,
  "臨床医学総論": 2,
  "生体計測装置学": 3,
  "医用治療機器学": 4,
  "医用機器安全管理学": 5,
  "医用電気電子工学": 6,
  "生体機能代行装置学": 7,
  "医用機械工学": 8,
  "生体物性材料工学": 9,
};
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////










////////////////////////////////////////////////////////　　メイン　　//////////////////////////////////////////////////////////

// メインプログラム
function doPost(e) {
  event = JSON.parse(e.postData.contents).events[0];                                    // イベント
  event_type = event.type;                                                              // イベントタイプ　取得
  replyToken = event.replyToken;                                                        // 応答用Tokenを取得
  user_id = event.source.userId;                                                        // ユーザーID　取得
  user_name = getUserProfile(user_id, LINE_TOKEN);                                      // ユーザー名　取得
  messageId = event.message.id;                                                         // メッセージID 取得
  LINE_END_POINT = 'https://api-data.line.me/v2/bot/message/' + messageId + '/content'; // LINE_END_POINT 取得

  try {

    // フォローイベント　のとき
    // if (event_type == "follow") {
    //   const account_sheet = active_ss.getSheetByName("Account"); // Accountシート 取得
    //   account_sheet.appendRow([user_name, user_id]); // スプレッドシート 入力
    //   account_sheet.getDataRange().removeDuplicates(); // 重複 削除
    // }

    // メッセージイベント　のとき
    if (event_type == "message") {
      userMessage = event.message.text; // メッセージを取得

      // メンテナンス
      if (maintenance_flag) sendText("ただいまメンテナンス中です。\n時間をおいてから\n再度お試しください。");

      // ユーザー別シートを持っているかどうか
      let user_sheetName = user_name + "（ID：" + user_id + "）";
      user_sheet = active_ss.getSheetByName(user_sheetName); // ユーザー別シート 取得
      if (!user_sheet) user_sheet = original_sheet.copyTo(active_ss).setName(user_sheetName); // コピー
      

      // メイン
      let file_list;
      if (userMessage == "ホーム") file_list = home(); // ホームメッセージ
      else if (userMessage.includes("ランダム ") && userMessage.replace("ランダム ", "") in field_dic) file_list = random(userMessage); // ランダム出題
      else if (userMessage == "sudo test") file_list = test(); // テスト
      else file_list = keyword_search(userMessage); // キーワード検索
      

      // ファイル　すべて送信
      for (let i = 0; i < file_list.length; ++i) {
        messages.push(file_list[i]); // 送信
      }
    }
  } catch (e) {
    sendText("エラー");
    for (let i = 0; i < errorMessages.length; ++i) {
      sendText(errorMessages[i]);
    }
    sendText(e.stack);
  }

  //lineで返信
  UrlFetchApp.fetch(LINE_URL, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': `Bearer ${LINE_TOKEN}`
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': replyToken,
      'messages': messages
    }),
  });
  ContentService.createTextOutput(JSON.stringify({ 'content': 'post ok' })).setMimeType(ContentService.MimeType.JSON);
  
  if (user_name != "*****") {
    // スプレッドシートにメッセージを保存
    message_sheet.appendRow([getDate(), user_name, userMessage]); // スプレッドシート 入力
    // スプレッドシートにIDを保存
    account_sheet.appendRow([user_name, user_id]); // スプレッドシート 入力
    account_sheet.getDataRange().removeDuplicates(); // 重複 削除
  }

  lock(); // 共有リンク 制限
}
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////





/////////////////////////////////////////////////////　　　ライブラリ　　　///////////////////////////////////////////////////////

// デバッグ
function debug(value='デバッグテスト') {
  active_ss.getSheetByName('Debug').appendRow([new Date(), value]);
}
// スプレッドシートに出力する



// テキスト 送信
function sendText(text) {
  const msg = {
    'type': 'text',
    'text': text
  }
  messages.push(msg);
}
// 引数のテキストを最後にまとめてLINEに送信する



// json 変換
function textListToJson(array) {
  for (i = 0; i < array.length; i++) {
    const msg = {
      'type': 'text',
      'text': array[i]
    }
    array[i] = msg;
  }
  return array;
}
// 引数の配列の各テキストをjsonに変換して配列で返す



// 単発
function sendMessages(messages) {
  //lineで返答する
  UrlFetchApp.fetch(LINE_URL, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': `Bearer ${LINE_TOKEN}`
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': replyToken,
      'messages': messages
    }),
  });
  ContentService.createTextOutput(JSON.stringify({ 'content': 'post ok' })).setMimeType(ContentService.MimeType.JSON);
}
// 引数のテキストをすぐにLINEへ送信し、処理を終了させる



// エラーコード 行数 取得
function getRowNum() {
  let e = new Error();
  e = e.stack.split("\n")[2].split(":");
  e.pop()
  return e.pop();
}
// このスクリプト内で何行目にあるかを取得する



// 問題図 共有リンク 制限
function lock() {
  let imgFolders = q_imgFolder.getFolders(); // 第1回、第2回
  const imgIdList = [];
  while (imgFolders.hasNext()) {
    const imgFolder = imgFolders.next().getFiles(); // 100101.jpg、100201.jpg
    while (imgFolder.hasNext()) {
      const imgId = imgFolder.next().getId();
      imgIdList.push(imgId);
    }
  }
  for (let i = 0; i < imgIdList.length; i++) {
    DriveApp.getFileById(imgIdList[i]).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT); // 共有リンク 制限
  }
}
//臨床工学技士国家試験＞問題図 内のすべての画像の共有を制限する



// ユーザーのプロフィール　取得
function getUserProfile(user_id, LINE_TOKEN) {
  const url = 'https://api.line.me/v2/bot/profile/' + user_id;
  const userProfile = UrlFetchApp.fetch(url, {
    'headers': {
      'Authorization': 'Bearer ' + LINE_TOKEN,
    },
  })
  return JSON.parse(userProfile).displayName;
}


// 日付 取得
function getDate(format = 'YYYY/MM/DD hh:mm:ss') {
  const date = new Date();
  format = format.replace(/YYYY/g, date.getFullYear());
  format = format.replace(/MM/g, ('0' + (date.getMonth() + 1)).slice(-2));
  format = format.replace(/DD/g, ('0' + date.getDate()).slice(-2));
  format = format.replace(/hh/g, ('0' + date.getHours()).slice(-2));
  format = format.replace(/mm/g, ('0' + date.getMinutes()).slice(-2));
  format = format.replace(/ss/g, ('0' + date.getSeconds()).slice(-2));
  return format;
}
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////





//////////////////////////////////////////////////　　　ランダム出題　　　　//////////////////////////////////////////////////////

// ランダム出題
function random(userMessage) {
  search(userMessage); // 検索

  // 結果
  let random = Math.floor(Math.random() * user_sheet.getLastRow()) + 2; // ランダム
  let result = user_sheet.getRange(random, 1, 1, 26).getValues(); // 1行 取得

  return resultToList(result)[0];
}
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////





//////////////////////////////////////////////　　　　　キーワード検索　　　　　　//////////////////////////////////////////////////

// キーワード 検索
function keyword_search(keyword) {

  // 検索
  search(keyword) // 検索

  // 結果
  let result = user_sheet.getRange(2, 1, user_sheet.getLastRow(), 26).getValues();
  result.pop();

  let files;
  if (result.length == 0) files = textListToJson(['「' + keyword + '」' + 'に一致する問題は見つかりませんでした。']);
  else if (keyword.includes('解答')) files = resultToList_a(result)[0];
  else if (result.length == 1) files = resultToList(result)[0];
  else if (1000 < result.length) files = textListToJson(["エラー：検索結果が多すぎます。\n" + "検索結果：" + result.length + "件 / 1000件"]);
  else files = textListToJson(createDoc(keyword, resultToList(result, true)));
  
  return files;
}

// 検索
function search(keyword) {
  user_sheet.clear(); // Userシート 初期化
  let db_sheets_len = database_ss.getSheets().length - 1; // データベースのシート数
  let database_ss_url = database_ss.getUrl(); // データベースのURL
  let time; // 回数
  let importRange; // QUERY関数で読み込むシートの範囲
  let col_num; // 列
  let operator; // 検索オプション

  // ランダム出題ならば
  if (keyword.includes('ランダム')) {
    let random = Math.floor(Math.random() * db_sheets_len) + 1; // ランダム
    time = String(random).padStart(2, '0'); // 0埋め
    col_num = 5; // E列
    operator = '='; // 一致
    keyword = "'" + keyword.replace('ランダム ', '') + "'";
    importRange = `IMPORTRANGE("${database_ss_url}", "第${time}回!A:Z")`;
  }

  // キーワード検索ならば
  else {
    keyword = keyword.includes('解答') ? keyword.replace('解答', '') : keyword; // "解答"を含むなら"解答"を消す

    // ID検索なら変換、メッセージの1行目参照　　36001,36A1,36AM1,1001,1A1,1AM1 → ID
    const re_text1 = '^([1-9]|[1-9][0-9])([01])(0[1-9]|[1-8][0-9]|90)$';
    const re_text2 = '^([1-9]|[1-9][0-9])([AP]M?)(0?[1-9]|[1-8][0-9]|90)$';
    const regexp1 = new RegExp(`${re_text1}`);
    const regexp2 = new RegExp(`${re_text2}`);
    const regexp = new RegExp(`${re_text1}|${re_text2}`);
    let id = keyword.match(regexp);

    // ID検索ならば
    if (id) {
      id = id[0].replace(regexp1, '$1-$2-$3').replace(regexp2, '$1-$2-$3');
      let l = id.split('-');
      time = l[0].padStart(2, '0'); // 0埋め
      col_num = 1; // A列
      operator = '='; // 一致
      keyword = l[0] + l[1].replace(/AM?/, '0').replace(/PM?/, '1') + l[2].padStart(2, '0');
      importRange = `IMPORTRANGE("${database_ss_url}", "第${time}回!A:Z")`;
    }

    // 文字検索ならば
    else {
      let sheetName_list = [];
      for (i = 0; i < db_sheets_len; i++) {
        time = String(i + 1).padStart(2, '0');
        sheetName_list.push(`IMPORTRANGE("${database_ss_url}", "第${time}回!A:Z")`);
      }
      col_num = 26; // Z列
      operator = 'contains'; // 含む
      keyword = "'" + keyword + "'";
      importRange = sheetName_list.join(';');
    }
  }

  let fx = `=QUERY({${importRange}}, "WHERE Col${col_num} ${operator} ${keyword}", true)`;
  user_sheet.getRange(1, 1).setValue(fx); // スプレッドシートに入力
}


// 検索結果をにリスト変換（問題）
function resultToList(result, doc = false) {
  let list = [];

  for (i = 0; i < result.length; i++) {
    let row = result[i];

    // 問題文
    let text = row[1] + '-' + row[2] + '-' + row[3] + '\n\n' + row[7] + '\n\n';
    if (row[8]) {
      text += 'a. ' + row[8] + '\n'
            + 'b. ' + row[9] + '\n'
            + 'c. ' + row[10] + '\n'
            + 'd. ' + row[11] + '\n'
            + 'e. ' + row[12] + '\n\n'
            + '1.' + row[13] + '　2.' + row[14] + '　3.' + row[15] + '　4.' + row[16] + '　5.' + row[17]
    }
    else if (row[13]) {
      text += '1. ' + row[13] + '\n'
            + '2. ' + row[14] + '\n'
            + '3. ' + row[15] + '\n'
            + '4. ' + row[16] + '\n'
            + '5. ' + row[17]
    }
    if (result.length >= 2) text += '\n\n' + '答え：' + row[21]; // 検索結果が2つ以上なら、解答を表示
    let msg = (doc) ? text : {'type': 'text', 'text': text}; // doc:text、LNE:json
    
    // クイックリプライボタン 追加
    if (result.length == 1) {
      msg = {
        'type': 'text', 'text': text,
        "quickReply": {
          "items": [
            {
              "type": "action",
              // "imageUrl": "",
              "action": {
                "type": "message",
                "label": row[0] + "解答",
                "text": row[0] + "解答"
              }
            }
          ]
        }
      }
    } 

    // 問題図
    let imgList = [];
    let id = row[0];
    let terms = "title = '" + id +  "01.jpg' or title = '" + id + "02.jpg' or title = '" + id + "03.jpg'";
    let imgFiles = q_imgFolder.getFoldersByName("第" + String(row[1]).padStart(2,'0') + "回").next().searchFiles(terms);
    while (imgFiles.hasNext()) {
      let imgFile = imgFiles.next();
      let img;

      // json
      if (!doc) {
        let imgId = imgFile.getId();
        DriveApp.getFileById(imgId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // 共有リンク 解除
        img = {
          'type': 'image',
          'originalContentUrl': 'https://drive.google.com/uc?id=' + imgId,
          'previewImageUrl': 'https://drive.google.com/uc?id=' + imgId,
        }
      }
      // doc
      else img = imgFile.getAs('image/png');

      imgList.push(img);
    }

    // リスト 追加
    list.push([msg].concat(imgList)); // [[問題,図1,図2,図3], [問題,図1,図2,図3]]
  }
  return list;
}

// 検索結果をにリスト変換（解答）
function resultToList_a(result) {
  let row = result[0];

  // 問題番号・解答
  let text = row[1] + '-' + row[2] + '-' + row[3] + '　' + '答え：' + row[21];
  let msg = {'type': 'text', 'text': text};

  // 解説図
  let imgList = [];
  let id = row[0];
  let terms = "title = '" + id +  "11.jpg' or title = '" + id + "12.jpg'";
  let imgFiles = a_imgFolder.getFoldersByName("第" + String(row[1]).padStart(2,'0') + "回").next().searchFiles(terms);
  while (imgFiles.hasNext()) {
    let imgFile = imgFiles.next();
    let imgId = imgFile.getId();
    DriveApp.getFileById(imgId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // 共有リンク 解除
    let img = {
      'type': 'image',
      'originalContentUrl': 'https://drive.google.com/uc?id=' + imgId,
      'previewImageUrl': 'https://drive.google.com/uc?id=' + imgId,
    }
    imgList.push(img);
  }
  return [[msg].concat(imgList)]; // [[問題番号,答え,解説図1,解説図2]]
}


// docファイル 作成
function createDoc(keyword, list) {
  let doc = DocumentApp.create(keyword + "_" + getDate('YYYYMMDDhhmmss') + "_" + user_name + user_id); // docファイル 作成
  let body = doc.getBody();

  // 問題・図 結合
  for (i = 0; i < list.length; i++) {
    // 問題文
    let text = (i == 0) ? '' : '\n\n';
    text += list[i][0];
    (i == 0) ? body.setText(text) : body.appendParagraph(text); // テキスト 追記

    // 図 挿入
    if (1 < list[i].length) for (l = 1; l < list[i].length; l++) body.insertImage(body.getNumChildren(), list[i][l]);
  }
  // 図 大きさ 変更
  for (const image of body.getImages()) {
    if (370 < image.getWidth()) {
      image.setHeight(Math.floor(image.getHeight()/image.getWidth()*370))
      image.setWidth(370);
    }
  }

  doc.saveAndClose(); // 保存・閉じる
  let doc_file = DriveApp.getFileById(doc.getId());
  doc_file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // docファイル 共有リンク 解除
  doc_file.moveTo(outputFolder); // docファイル→出力ファイル 移動
  doc_file.setTrashed(true); // docファイル 削除

  return ["ファイルを作成しました。\nロードに時間がかかる場合は、\n時間をおいて開いてください。", doc_file.getUrl()];
}
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////





/////////////////////////////////////////////////　　　　　ホーム　　　　　　/////////////////////////////////////////////////////

// ホームボタン
function home() {
  return textListToJson([home_doc1.getBody().getText(), home_doc2.getBody().getText(), info_doc.getBody().getText()]);
}
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
