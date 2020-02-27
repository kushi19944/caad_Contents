import RPA from 'ts-rpa';
import { WebDriver, By } from 'selenium-webdriver';
var fs = require('fs');
// ADX にログインする際の ID/PW
const ADX_ID = process.env.Billbord_CyNumber;
const ADX_PW = process.env.Billbord_CyPass;
const ADX_URL = process.env.Billbord_ADX_URL;
const ADX_BulkUP_URL = process.env.Billbord_ADX_BulkUp_URL;
// 読み込みする スプレッドシートID と シート名 の記載
const SSID = process.env.Billbord_Sheet_ID;
const SSName1 = process.env.Billbord_Sheet_Name;
// スプレッドシートから読み込む行数を記載する
const Sheet_StartRow = 10;
const Sheet_LastRow = 500;

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
    const DOM = await RPA.WebBrowser.driver.getPageSource();
    await RPA.Logger.info(DOM);
    await RPA.SystemLogger.info(error);
    await RPA.WebBrowser.takeScreenshot();
    await RPA.Logger.info(
      'エラー出現.スクリーンショット撮ってブラウザ終了します'
    );
  } finally {
    await RPA.WebBrowser.quit();
  }
}

Start();

async function Work() {
  const firstSheetData = [];
  const SheetWorkingRow = [];
  const LoopFlag = ['true'];
  const firstSheetID_SheetName = [];
  // ADX ログイン
  await ADXLogin();
  while (0 == 0) {
    LoopFlag[0] = 'false';
    // 作業する行のデータを取得
    await ReadSheet(
      firstSheetData,
      LoopFlag,
      SheetWorkingRow,
      firstSheetID_SheetName
    );
    const SheetData = firstSheetData[0];
    const SheetID_SheetName = firstSheetID_SheetName[0];
    // 作業する行がある場合のみ、処理を行う
    if (LoopFlag[0] == 'true') {
      // 画像容量の判定
      await ImageSizeJudge(SheetData, SheetWorkingRow);
      // バルクアップを行う
      await BulkUp(SheetData, SheetWorkingRow, SheetID_SheetName);
    }
    if (LoopFlag[0] == 'false') {
      await RPA.Logger.info('全ての作業完了しました.RPA終了します');
      break;
    }
  }
}

async function ADXLogin() {
  //Adxにログインする
  await RPA.WebBrowser.get(ADX_URL);
  const LoginBUtton1 = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      xpath: '//*[@id="root"]/section/section/div/div/a'
    }),
    5000
  );
  RPA.WebBrowser.mouseClick(LoginBUtton1);
  const UserID = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ xpath: '//*[@id="username"]' }),
    5000
  );
  RPA.WebBrowser.sendKeys(UserID, [ADX_ID]);
  const UserPass = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="password"]'
  );
  RPA.WebBrowser.sendKeys(UserPass, [ADX_PW]);
  const LoginButton2 = await RPA.WebBrowser.findElementByXPath(
    '/html/body/div/div[2]/div/form/div[6]/a'
  );
  await RPA.WebBrowser.mouseClick(LoginButton2);
  await RPA.sleep(1000);
}

async function ReadSheet(
  SheetData,
  LoopFlag,
  SheetWorkingRow,
  firstSheetID_SheetName
) {
  // スプレッドシートのID・シート名 を取得
  firstSheetID_SheetName[0] = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!Q5:R5`
  });
  // スプレッドシートからデータを取得
  const FirstData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!M${String(Sheet_StartRow)}:AD${String(Sheet_LastRow)}`
  });
  for (let i in FirstData) {
    if (FirstData[i][4] == '') {
      continue;
    }
    // Q列 が 入稿チェック済みかつ、 R列が空白 の行のみ取得
    if (FirstData[i][4] == '入稿チェック済み（沖縄）') {
      if (FirstData[i][17] == '') {
        SheetData[0] = FirstData[i];
        SheetWorkingRow[0] = Number(i) + Number(Sheet_StartRow);
        LoopFlag[0] = 'true';
        await SheetPasteData(
          SSName1,
          `AD${SheetWorkingRow[0]}`,
          `AD${SheetWorkingRow[0]}`,
          [['作業中']]
        );
        break;
      }
      if (FirstData[i][17] == undefined) {
        SheetData[0] = FirstData[i];
        SheetWorkingRow[0] = Number(i) + Number(Sheet_StartRow);
        LoopFlag[0] = 'true';
        await SheetPasteData(
          SSName1,
          `AD${SheetWorkingRow[0]}`,
          `AD${SheetWorkingRow[0]}`,
          [['作業中']]
        );
        break;
      }
    }
  }
}

async function ImageSizeJudge(SheetData, SheetWorkingRow) {
  if (SheetData[9] == '') {
    await RPA.Logger.info('GoogleDrive 無し');
    return;
  }
  if (SheetData[9] != '') {
    await RPA.Logger.info('GoogleDrive 有り. 画像サイズ確認します');
    await SheetPasteData('画像サイズ判定', 'B2', 'B2', [[`${SheetData[9]}`]]);
    while (0 == 0) {
      await RPA.sleep(3000);
      const SizeData = await RPA.Google.Spreadsheet.getValues({
        spreadsheetId: `${SSID}`,
        range: `画像サイズ判定!B3:B3`
      });
      await RPA.Logger.info(SizeData);
      if (Number(SizeData[0][0]) > 2000000) {
        await SheetPasteData(SSName1, SheetWorkingRow[0], SheetWorkingRow[0], [
          ['画像サイズオーバー']
        ]);
      }
      // 画像サイズを取得した後はスプレッドシートB2/B3を消す
      if (SizeData != undefined) {
        await SheetPasteData('画像サイズ判定', 'B2', 'B3', [[''], ['']]);
        break;
      }
    }
  }
}

// スプレッドシートにデータを貼り付ける関数
async function SheetPasteData(SheetName, StartRange, EndRange, Data) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SheetName}!${StartRange}:${EndRange}`,
    values: Data
  });
}

// バルクアップの処理
async function BulkUp(SheetData, SheetWorkingRow, SheetID_SheetName) {
  await RPA.WebBrowser.get(ADX_BulkUP_URL);
  await RPA.sleep(2000);
  const UploadButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className: 'btn btn-sm btn-primary'
    }),
    5000
  );
  await RPA.WebBrowser.mouseClick(UploadButton);
  await RPA.sleep(500);
  const SheetIDInput = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      id: 'sheetId'
    }),
    5000
  );
  await RPA.WebBrowser.sendKeys(SheetIDInput, [SheetID_SheetName[0][0]]);
  await RPA.sleep(100);
  const SheetNameInput = await RPA.WebBrowser.findElementById('sheetName');
  await RPA.WebBrowser.sendKeys(SheetNameInput, [SheetID_SheetName[0][1]]);
  await RPA.sleep(100);
  const StartRowInput = await RPA.WebBrowser.findElementById('startRow');
  await StartRowInput.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(StartRowInput, [SheetWorkingRow[0]]);
  await RPA.sleep(100);
  await RPA.WebBrowser.driver.executeScript(
    `document.querySelector('body > div:nth-child(5) > div > div.fade.in.modal > div > div > div.modal-body > div > form > div.form-group.form-inline > div > div:nth-child(2) > div > div > label > input[type=checkbox]').click()`
  );
  await RPA.sleep(200);
  const EndRowInput = await RPA.WebBrowser.findElementById('endRow');
  await RPA.WebBrowser.sendKeys(EndRowInput, [SheetWorkingRow[0]]);
  const SettingOptionList = await RPA.WebBrowser.findElementsByCSSSelector(
    '#settingId > option'
  );
  // 読み込みセッティングがあるかどうか判定
  const SettingFlag = ['false'];
  for (let i in SettingOptionList) {
    const OptionText = await SettingOptionList[i].getText();
    if (SheetData[13] == OptionText) {
      await RPA.Logger.info(
        '読込セッティング一致しました.バルクアップ開始します'
      );
      const OptionNumber = Number(i) + 1;
      const SettingOption = await RPA.WebBrowser.findElementByCSSSelector(
        `#settingId > option:nth-child(${OptionNumber})`
      );
      const ValueText = await SettingOption.getAttribute(`value`);
      await RPA.WebBrowser.driver
        .findElement(
          By.xpath(`//*[@id="settingId"]/option[@value="${ValueText}"]`)
        )
        .click();
      SettingFlag[0] = 'true';
      await SheetPasteData(
        SSName1,
        `AD${SheetWorkingRow[0]}`,
        `AD${SheetWorkingRow[0]}`,
        [['完了']]
      );
      const ApplyButton = await RPA.WebBrowser.findElementByCSSSelector(
        'body > div:nth-child(5) > div > div.fade.in.modal > div > div > div.modal-footer > div > button.btn.btn-primary'
      );
      await RPA.WebBrowser.mouseClick(ApplyButton);
      await RPA.sleep(500);
      // アップロード中です の文字が消えるまでループ待機
      while (0 == 0) {
        await RPA.sleep(300);
        const Alert = await RPA.WebBrowser.findElementsByClassName(
          'alert alert-info'
        );
        if (Alert.length == 0) {
          break;
        }
      }
      break;
    }
  }
  // 読込セッティング がなければスキップ処理する
  if (SettingFlag[0] == 'false') {
    await RPA.Logger.info('読込セッティングで一致するものがありませんでした.');
    await SheetPasteData(
      SSName1,
      `AD${SheetWorkingRow[0]}`,
      `AD${SheetWorkingRow[0]}`,
      [['読込セッティングが一致しません.作成してください']]
    );
  }
}
