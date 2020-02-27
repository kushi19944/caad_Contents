import RPA from 'ts-rpa';
import { WebDriver, By } from 'selenium-webdriver';
var fs = require('fs');
// ADX にログインする際の ID/PW
const ADX_ID = process.env.Billbord_CyNumber;
const ADX_PW = process.env.Billbord_CyPass;
const ADX_URL = process.env.Billbord_ADX_URL;
const ADX_BulkUP_URL = process.env.Billbord_ADX_BulkUp_URL;
// RPAトリガーシートID と シート名 の記載
const Trigger_SheetID = process.env.Billbord_Sheet_ID;
const Trigger_SheetName = process.env.Billbord_Sheet_Name;
const Trigger_Sheet_StartRow = 2;
const Trigger_Sheet_LastRow = 200;
// 作業行うスプレッドシートから読み込む行数を記載する
const Sheet_StartRow = 10;
const Sheet_LastRow = 1000;
// 作業行うシートID・シート名 を格納する変数
const WorkingSheetID = [];
const WorkingSheetName = [];

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
  const TriggerSheetRow = [];
  const WorkingSheetRow = [];
  const ParentLoopFlag = ['false'];
  const LoopFlag = ['false'];
  // ADX ログイン
  await ADXLogin();
  while (0 == 0) {
    ParentLoopFlag[0] = 'false';
    // RPAトリガーシート を読み込んで作業するシートを取得する
    await ReadSheetRPATrigger(ParentLoopFlag, TriggerSheetRow);
    while (0 == 0) {
      LoopFlag[0] = 'false';
      // 作業する行のデータを取得
      await ReadSheet(firstSheetData, LoopFlag, WorkingSheetRow);
      const SheetData = firstSheetData[0];
      // 作業する行がある場合のみ、処理を行う
      if (LoopFlag[0] == 'true') {
        // 画像容量の判定
        await ImageSizeJudge(SheetData, WorkingSheetRow);
        // バルクアップを行う
        await BulkUp(SheetData, WorkingSheetRow);
      }
      if (LoopFlag[0] == 'false') {
        await TriggerSheetPasteData(TriggerSheetRow, [['完了']]);
        await RPA.Logger.info(
          `このシートは完了しました. 【${WorkingSheetName}】`
        );
        break;
      }
    }
    if (ParentLoopFlag[0] == 'false') {
      await RPA.Logger.info('＊＊＊全ての作業完了しました.RPA終了します＊＊＊');
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

async function ReadSheetRPATrigger(ParentLoopFlag, TriggerSheetRow) {
  // RPAトリガーシートから ID・シート名 を取得
  const TriggerData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${Trigger_SheetID}`,
    range: `${Trigger_SheetName}!B${Trigger_Sheet_StartRow}:D${Trigger_Sheet_LastRow}`
  });
  for (let i in TriggerData) {
    if (TriggerData[i][1] == undefined) {
      continue;
    }
    if (TriggerData[i][1] == '') {
      continue;
    }
    if (TriggerData[i][2] == '作業中') {
      continue;
    }
    if (TriggerData[i][2] == '完了') {
      continue;
    }
    if (TriggerData[i][2] == undefined) {
      ParentLoopFlag[0] = 'true';
      const ID = await TriggerData[i][1].split('/');
      WorkingSheetID[0] = ID[5];
      WorkingSheetName[0] = TriggerData[i][0];
      TriggerSheetRow[0] = Number(i) + Number(Trigger_Sheet_StartRow);
      await TriggerSheetPasteData(TriggerSheetRow, [['作業中']]);
      await RPA.Logger.info(
        ` RPAトリガーシート ${TriggerSheetRow[0]} 行目を作業開始します 【${WorkingSheetName}】`
      );
      break;
    }
    if (TriggerData[i][2] == '') {
      ParentLoopFlag[0] = 'true';
      const ID = await TriggerData[i][1].split('/');
      WorkingSheetID[0] = ID[5];
      WorkingSheetName[0] = TriggerData[i][0];
      TriggerSheetRow[0] = Number(i) + Number(Trigger_Sheet_StartRow);
      await TriggerSheetPasteData(TriggerSheetRow, [['作業中']]);
      await RPA.Logger.info(
        ` RPAトリガーシート ${TriggerSheetRow[0]} 行目を作業開始します 【${WorkingSheetName}】`
      );
      break;
    }
  }
}

async function TriggerSheetPasteData(TriggerSheetRow, Data) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${Trigger_SheetID}`,
    range: `${Trigger_SheetName}!D${TriggerSheetRow[0]}:D${TriggerSheetRow[0]}`,
    values: Data
  });
}

async function ReadSheet(SheetData, LoopFlag, WorkingSheetRow) {
  // スプレッドシートからデータを取得
  const FirstData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${WorkingSheetID[0]}`,
    range: `${WorkingSheetName[0]}!M${String(Sheet_StartRow)}:AD${String(
      Sheet_LastRow
    )}`
  });
  for (let i in FirstData) {
    if (FirstData[i][4] == '') {
      continue;
    }
    // Q列 が 入稿チェック済みかつ、 R列が空白 の行のみ取得
    if (FirstData[i][4] == '入稿チェック済み（沖縄）') {
      if (FirstData[i][17] == '') {
        SheetData[0] = FirstData[i];
        WorkingSheetRow[0] = Number(i) + Number(Sheet_StartRow);
        LoopFlag[0] = 'true';
        await SheetPasteData(
          WorkingSheetName[0],
          `AD${WorkingSheetRow[0]}`,
          `AD${WorkingSheetRow[0]}`,
          [['作業中']]
        );
        break;
      }
      if (FirstData[i][17] == undefined) {
        SheetData[0] = FirstData[i];
        WorkingSheetRow[0] = Number(i) + Number(Sheet_StartRow);
        LoopFlag[0] = 'true';
        await SheetPasteData(
          WorkingSheetName[0],
          `AD${WorkingSheetRow[0]}`,
          `AD${WorkingSheetRow[0]}`,
          [['作業中']]
        );
        break;
      }
    }
  }
}

async function ImageSizeJudge(SheetData, WorkingSheetRow) {
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
        spreadsheetId: `${Trigger_SheetID}`,
        range: `画像サイズ判定!B3:B3`
      });
      await RPA.Logger.info(SizeData);
      if (Number(SizeData[0][0]) > 2000000) {
        await SheetPasteData(
          WorkingSheetName[0],
          WorkingSheetRow[0],
          WorkingSheetRow[0],
          [['画像サイズオーバー']]
        );
      }
      // 画像サイズを取得した後はスプレッドシートB2/B3を消す
      if (SizeData != undefined) {
        await SheetPasteData('画像サイズ判定', 'B2', 'B3', [[''], ['']]);
        break;
      }
    }
  }
}

// 作業しているスプレッドシートにデータを貼り付ける関数
async function SheetPasteData(SheetName, StartRange, EndRange, Data) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${WorkingSheetID[0]}`,
    range: `${SheetName}!${StartRange}:${EndRange}`,
    values: Data
  });
}

// バルクアップの処理
async function BulkUp(SheetData, WorkingSheetRow) {
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
  await RPA.WebBrowser.sendKeys(SheetIDInput, [WorkingSheetID[0]]);
  await RPA.sleep(100);
  const SheetNameInput = await RPA.WebBrowser.findElementById('sheetName');
  await RPA.WebBrowser.sendKeys(SheetNameInput, [WorkingSheetName[0]]);
  await RPA.sleep(100);
  const StartRowInput = await RPA.WebBrowser.findElementById('startRow');
  await StartRowInput.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(StartRowInput, [WorkingSheetRow[0]]);
  await RPA.sleep(100);
  await RPA.WebBrowser.driver.executeScript(
    `document.querySelector('body > div:nth-child(5) > div > div.fade.in.modal > div > div > div.modal-body > div > form > div.form-group.form-inline > div > div:nth-child(2) > div > div > label > input[type=checkbox]').click()`
  );
  await RPA.sleep(200);
  const EndRowInput = await RPA.WebBrowser.findElementById('endRow');
  await RPA.WebBrowser.sendKeys(EndRowInput, [WorkingSheetRow[0]]);
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
        WorkingSheetName[0],
        `AD${WorkingSheetRow[0]}`,
        `AD${WorkingSheetRow[0]}`,
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
      WorkingSheetName[0],
      `AD${WorkingSheetRow[0]}`,
      `AD${WorkingSheetRow[0]}`,
      [['読込セッティングが一致しません.作成してください']]
    );
  }
}
