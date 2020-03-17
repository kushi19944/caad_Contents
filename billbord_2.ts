import RPA from 'ts-rpa';
import { WebDriver, By } from 'selenium-webdriver';
var fs = require('fs');

// ＊＊＊ Slack通知 ＊＊＊
const Slack_Token = process.env.AbemaTV_hubot_Token;
//const Slack_Channel = process.env.RPA_Test_Channel;
const Slack_Channel = process.env.Contents_Slack_Channel;
const Slack_Text = [`ビルボード AdxID転記 問題なく完了しました`];

// RPAトリガーシートID と シート名 の記載
const Trigger_SheetID = process.env.Billbord_Sheet_ID;
const Trigger_SheetName = process.env.Billbord_Sheet_Name;
// 作業行うスプレッドシートから読み込む行数を記載する
const Sheet_StartRow = 10;
const Sheet_LastRow = 1000;

async function Start() {
  // 実行前にダウンロードフォルダを全て削除する
  await RPA.File.rimraf({ dirPath: `${process.env.WORKSPACE_DIR}` });
  // デバッグログを最小限にする
  RPA.Logger.level = 'INFO';
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
  });
  try {
    await Work();
  } catch (error) {
    // エラー発生時の処理
    //const DOM = await RPA.WebBrowser.driver.getPageSource();
    //await RPA.Logger.info(DOM);
    Slack_Text[0] = `ビルボード AdxID転記 エラー発生です\n@kushi_makoto 確認してください`;
    await RPA.SystemLogger.info(error);
    await RPA.WebBrowser.takeScreenshot();
    await RPA.Logger.info(
      'エラー出現.スクリーンショット撮ってブラウザ終了します'
    );
  } finally {
    await SlackPost(Slack_Text[0]);
    await RPA.WebBrowser.quit();
  }
}

Start();

async function Work() {
  // シリーズのデータを配列で保持しておく
  const Series_AllData = [
    [], //'01_シリーズ（任意指定なし）'
    [], //'02_シリーズ（キャプション・画像・遷移先URL指定あり）'
    [], //'03_シリーズ（画像・遷移先URL指定あり）'
    [], //'04_シリーズ（キャプション・遷移先URL指定あり）'
    [], //'05_シリーズ（キャプション・画像指定あり）'
    [], //'06_シリーズ（遷移先URL指定あり）'
    [], //'07_シリーズ（画像指定あり）'
    [], //'08_シリーズ（キャプション指定あり）'
    [], //'09_LP'
    [], //'10_スロット（任意指定なし）'
    [], //'11_スロット（キャプション・画像指定あり）'
    [], //'12_スロット（キャプション指定あり）'
    [] //'13_スロット（画像指定あり）'
  ];
  // 作業開始前の Slack通知
  await SlackPost(`ビルボード AdxID転記 RPA開始します`);
  const FirstSheetData = [];
  for (let NewNumber = 1; NewNumber < Series_AllData.length + 1; NewNumber++) {
    // 各シートからデータを取得し、元シートに AdxID を転記する
    await GetSheetData(NewNumber);
  }
}

// 作業するデータを取得する
async function GetSheetData(NewNumber) {
  // スプレッドシートからデータを取得
  if (NewNumber >= 10) {
    const FirstData = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${Trigger_SheetID}`,
      range: `${NewNumber}!B${Sheet_StartRow}:AF${Sheet_LastRow}`
    });
    for (let i in FirstData) {
      if (FirstData[i][29] == undefined) {
        continue;
      }
      if (FirstData[i][25] == '') {
        await RPA.Logger.info('AdxID未発行です.スキップします');
        continue;
      }
      if (FirstData[i][30] == undefined) {
        await RPA.Logger.info(`シート : ${NewNumber} 転記します`);
        await RPA.Logger.info(
          `転記先 : ${FirstData[i][29]}  ${FirstData[i][0]} 行目`
        );
        // バルクアップ用シートに 転記済み と記載する
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: Trigger_SheetID,
          range: `${NewNumber}!AF${Number(i) + 10}:AF${Number(i) + 10}`,
          values: [[`転記済み`]],
          parseValues: true
        });
        // AdxID を元シートに転記する
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${FirstData[i][28]}`,
          range: `${FirstData[i][29]}!AA${FirstData[i][0]}:AA${FirstData[i][0]}`,
          values: [[`${FirstData[i][25]}`]],
          parseValues: true
        });
      }
    }
  }
  if (NewNumber <= 9) {
    const FirstData = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${Trigger_SheetID}`,
      range: `0${NewNumber}!B${Sheet_StartRow}:AF${Sheet_LastRow}`
    });
    for (let i in FirstData) {
      if (FirstData[i][29] == undefined) {
        continue;
      }
      if (FirstData[i][25] == '') {
        await RPA.Logger.info('AdxID未発行です.スキップします');
        continue;
      }
      if (FirstData[i][30] == undefined) {
        await RPA.Logger.info(`シート : 0${NewNumber} 転記します`);
        await RPA.Logger.info(
          `転記先 : ${FirstData[i][29]}  ${FirstData[i][0]} 行目`
        );
        // バルクアップ用シートに 転記済み と記載する
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: Trigger_SheetID,
          range: `0${NewNumber}!AF${Number(i) + 10}:AF${Number(i) + 10}`,
          values: [[`転記済み`]],
          parseValues: true
        });
        // AdxID を元シートに転記する
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${FirstData[i][28]}`,
          range: `${FirstData[i][29]}!AA${FirstData[i][0]}:AA${FirstData[i][0]}`,
          values: [[`${FirstData[i][25]}`]],
          parseValues: true
        });
      }
    }
  }
}

// Slack通知の関数
async function SlackPost(Text) {
  // 作業開始時にSlackへ通知する
  await RPA.Slack.chat.postMessage({
    channel: Slack_Channel,
    token: Slack_Token,
    text: `${Text}`,
    icon_emoji: ':snowman:',
    username: 'p1'
  });
}
