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
  //RPA.Logger.level = 'INFO';
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
  const TriggerSheetRow = [];
  const ParentLoopFlag = ['false'];
  const Series_AllData_Copy = [];
  // バルクアップ転記用シートをクリアする
  await SheetClear();
  while (0 == 0) {
    ParentLoopFlag[0] = 'false';
    // RPAトリガーシート を読み込んで 転記作業するシートID/シート名を取得する
    await GetSheetID_SheetName(ParentLoopFlag, TriggerSheetRow);
    if (ParentLoopFlag[0] == 'false') {
      await RPA.Logger.info(
        '＊＊＊転記作業 完了しました. バルクアップへ移行します＊＊＊'
      );
      break;
    }
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
    // 転記作業する行のデータを取得
    await ReadSheet(Series_AllData);
    // 他の関数でも使えるようにデータをコピー
    Series_AllData_Copy[0] = await Series_AllData;
    for (let i in Series_AllData) {
      if (Series_AllData[i].length > 0) {
        // 対応フラグごとにそれぞれのシートに分けて転記
        await BulkUpSheetPaste(Series_AllData[i], i);
        // 対象のシート転記を終えたら、 転記完了と記載する
        await PasteWorkEnd(TriggerSheetRow);
      }
    }
  }
  // ＊＊＊ バルクアップする処理 ＊＊＊
  // ADX ログイン
  await ADXLogin();
  for (let i in Series_AllData_Copy[0]) {
    if (Series_AllData_Copy[0][i].length > 0) {
      const FirstData = [];
      const LastRow = [];
      // バルクアップを行うシートのデータを取得する
      await GetBulkUpData(FirstData, LastRow, i);
      const SheetData = [];
      SheetData[0] = await FirstData[0];
      await RPA.Logger.info(SheetData);
      await RPA.Logger.info(`最終行 → ` + LastRow);
      // バルクアップ を行う処理
      await BulkUp(SheetData, LastRow, i);
      // バルクアップ 確定を行う処理
      await BulkUpApply(i);
    }
  }
  await RPA.Logger.info(
    '＊＊＊バルクアップ 完了しました. AdxID発行後 フェーズ2を実行してください＊＊＊'
  );
}

async function GetSheetID_SheetName(ParentLoopFlag, TriggerSheetRow) {
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
    if (TriggerData[i][2] == '転記作業中') {
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
      await SetSheetValues(
        Trigger_SheetID,
        Trigger_SheetName,
        `D${TriggerSheetRow[0]}`,
        `D${TriggerSheetRow[0]}`,
        [['転記作業中']]
      );
      break;
    }
    if (TriggerData[i][2] == '') {
      ParentLoopFlag[0] = 'true';
      const ID = await TriggerData[i][1].split('/');
      WorkingSheetID[0] = ID[5];
      WorkingSheetName[0] = TriggerData[i][0];
      TriggerSheetRow[0] = Number(i) + Number(Trigger_Sheet_StartRow);
      await SetSheetValues(
        Trigger_SheetID,
        Trigger_SheetName,
        `D${TriggerSheetRow[0]}`,
        `D${TriggerSheetRow[0]}`,
        [['転記作業中']]
      );
      break;
    }
  }
}

async function ReadSheet(Series_AllData) {
  // スプレッドシートからデータを取得
  const FirstData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${WorkingSheetID[0]}`,
    range: `${WorkingSheetName[0]}!B${Sheet_StartRow}:Z${Sheet_LastRow}`
  });
  for (let i in FirstData) {
    await FirstData[i].push('');
    await FirstData[i].push('');
    await FirstData[i].push('');
    // AD列に シートID. AE列に シート名 を記載するために配列プッシュする
    await FirstData[i].push(`${WorkingSheetID[0]}`);
    await FirstData[i].push(`${WorkingSheetName[0]}`);
    // 入稿ステータスが 空白ならコンティニュー
    if (FirstData[i][15] == '') {
      continue;
    }
    // Q列 が 入稿チェック済みかつ、 画像が規定サイズ(2000kb)内の行のみ取得
    if (FirstData[i][15] == '入稿チェック済み（沖縄）') {
      if (FirstData[i][20] == '') {
        if (FirstData[i][24] == '01_シリーズ（任意指定なし）') {
          Series_AllData[0].push(FirstData[i]);
        }
        if (
          FirstData[i][24] ==
          '02_シリーズ（キャプション・画像・遷移先URL指定あり）'
        ) {
          Series_AllData[1].push(FirstData[i]);
        }
        if (FirstData[i][24] == '03_シリーズ（画像・遷移先URL指定あり）') {
          Series_AllData[2].push(FirstData[i]);
        }
        if (
          FirstData[i][24] == '04_シリーズ（キャプション・遷移先URL指定あり）'
        ) {
          Series_AllData[3].push(FirstData[i]);
        }
        if (FirstData[i][24] == '05_シリーズ（キャプション・画像指定あり）') {
          Series_AllData[4].push(FirstData[i]);
        }
        if (FirstData[i][24] == '06_シリーズ（遷移先URL指定あり）') {
          Series_AllData[5].push(FirstData[i]);
        }
        if (FirstData[i][24] == '07_シリーズ（画像指定あり）') {
          Series_AllData[6].push(FirstData[i]);
        }
        if (FirstData[i][24] == '08_シリーズ（キャプション指定あり）') {
          Series_AllData[7].push(FirstData[i]);
        }
        if (FirstData[i][24] == '09_LP') {
          Series_AllData[8].push(FirstData[i]);
        }
        if (FirstData[i][24] == '10_スロット（任意指定なし）') {
          Series_AllData[9].push(FirstData[i]);
        }
        if (FirstData[i][24] == '11_スロット（キャプション・画像指定あり）') {
          Series_AllData[10].push(FirstData[i]);
        }
        if (FirstData[i][24] == '12_スロット（キャプション指定あり）') {
          Series_AllData[11].push(FirstData[i]);
        }
        if (FirstData[i][24] == '13_スロット（画像指定あり）') {
          Series_AllData[12].push(FirstData[i]);
        }
      }
      // GoogleDrive の画像サイズを判定する
      if (FirstData[i][20] != '') {
        await RPA.Logger.info('画像サイズ判定します');
        await SheetPasteData('ビルボードRPAトリガー', `G10`, `G10`, [
          [`${FirstData[i][20]}`]
        ]);
        while (0 == 0) {
          await RPA.sleep(7000);
          const SizeData = await RPA.Google.Spreadsheet.getValues({
            spreadsheetId: `${Trigger_SheetID}`,
            range: `ビルボードRPAトリガー!G11:G11`
          });
          if (Number(SizeData[0][0]) > 2000000) {
            await SheetPasteData(
              WorkingSheetName[0],
              `AD${FirstData[i][0]}`,
              `AD${FirstData[i][0]}`,
              [['画像サイズオーバー']]
            );
          }
          if (FirstData[i][24] == '01_シリーズ（任意指定なし）') {
            Series_AllData[0].push(FirstData[i]);
          }
          if (
            FirstData[i][24] ==
            '02_シリーズ（キャプション・画像・遷移先URL指定あり）'
          ) {
            Series_AllData[1].push(FirstData[i]);
          }
          if (FirstData[i][24] == '03_シリーズ（画像・遷移先URL指定あり）') {
            Series_AllData[2].push(FirstData[i]);
          }
          if (
            FirstData[i][24] == '04_シリーズ（キャプション・遷移先URL指定あり）'
          ) {
            Series_AllData[3].push(FirstData[i]);
          }
          if (FirstData[i][24] == '05_シリーズ（キャプション・画像指定あり）') {
            Series_AllData[4].push(FirstData[i]);
          }
          if (FirstData[i][24] == '06_シリーズ（遷移先URL指定あり）') {
            Series_AllData[5].push(FirstData[i]);
          }
          if (FirstData[i][24] == '07_シリーズ（画像指定あり）') {
            Series_AllData[6].push(FirstData[i]);
          }
          if (FirstData[i][24] == '08_シリーズ（キャプション指定あり）') {
            Series_AllData[7].push(FirstData[i]);
          }
          if (FirstData[i][24] == '09_LP') {
            Series_AllData[8].push(FirstData[i]);
          }
          if (FirstData[i][24] == '10_スロット（任意指定なし）') {
            Series_AllData[9].push(FirstData[i]);
          }
          if (FirstData[i][24] == '11_スロット（キャプション・画像指定あり）') {
            Series_AllData[10].push(FirstData[i]);
          }
          if (FirstData[i][24] == '12_スロット（キャプション指定あり）') {
            Series_AllData[11].push(FirstData[i]);
          }
          if (FirstData[i][24] == '13_スロット（画像指定あり）') {
            Series_AllData[12].push(FirstData[i]);
          }
          // 画像サイズを取得した後はスプレッドシートB2/B3を消す
          if (SizeData != undefined) {
            await SheetPasteData('ビルボードRPAトリガー', 'G10', 'G11', [
              [''],
              ['']
            ]);
            break;
          }
        }
      }
    }
  }
}

// 作業しているスプレッドシートにデータを貼り付ける関数
async function SheetPasteData(SheetName, StartRange, EndRange, Data) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${WorkingSheetID[0]}`,
    range: `${SheetName}!${StartRange}:${EndRange}`,
    values: Data,
    parseValues: true
  });
}

// スプレッドシートにデータを貼り付ける関数
async function SetSheetValues(SheetID, SheetName, StartRange, LastRange, Data) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SheetID}`,
    range: `${SheetName}!${StartRange}:${LastRange}`,
    values: Data,
    parseValues: true
  });
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

// バルクアップの処理
async function BulkUp(SheetData, LastRow, i) {
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
  await RPA.WebBrowser.sendKeys(SheetIDInput, [Trigger_SheetID]);
  await RPA.sleep(100);
  const SheetNameInput = await RPA.WebBrowser.findElementById('sheetName');
  const NewNumber = Number(i) + 1;
  if (NewNumber >= 10) {
    await RPA.WebBrowser.sendKeys(SheetNameInput, [`${NewNumber}`]);
  }
  if (NewNumber <= 9) {
    await RPA.WebBrowser.sendKeys(SheetNameInput, [`0${NewNumber}`]);
  }
  await RPA.sleep(100);
  const StartRowInput = await RPA.WebBrowser.findElementById('startRow');
  await StartRowInput.clear();
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(StartRowInput, ['10']);
  await RPA.sleep(100);
  await RPA.WebBrowser.driver.executeScript(
    `document.querySelector('body > div:nth-child(5) > div > div.fade.in.modal > div > div > div.modal-body > div > form > div.form-group.form-inline > div > div:nth-child(2) > div > div > label > input[type=checkbox]').click()`
  );
  await RPA.sleep(200);
  const EndRowInput = await RPA.WebBrowser.findElementById('endRow');
  await RPA.WebBrowser.sendKeys(EndRowInput, [`${LastRow[0]}`]);
  const SettingOptionList = await RPA.WebBrowser.findElementsByCSSSelector(
    '#settingId > option'
  );
  for (let i in SettingOptionList) {
    const OptionText = await SettingOptionList[i].getText();
    if (SheetData[0] == OptionText) {
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
      const ApplyButton = await RPA.WebBrowser.findElementByCSSSelector(
        'body > div:nth-child(5) > div > div.fade.in.modal > div > div > div.modal-footer > div > button.btn.btn-primary'
      );
      await RPA.WebBrowser.mouseClick(ApplyButton);
      await RPA.sleep(500);
      // アップロード中です の文字が消えるまでループ待機
      while (0 == 0) {
        await RPA.sleep(700);
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
}

// バルクアップのシート貼り付ける前に シートをクリアする
async function SheetClear() {
  // シートクリア と入力して スプレッドシートを削除する
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${Trigger_SheetID}`,
    range: `ビルボードRPAトリガー!G7:G7`,
    values: [['シートクリア']],
    parseValues: true
  });
  await RPA.sleep(3000);
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${Trigger_SheetID}`,
    range: `ビルボードRPAトリガー!G7:G7`,
    values: [['']],
    parseValues: true
  });
}

// バルクアップ用の スプレッドシートにデータを 分けて転記
async function BulkUpSheetPaste(Series_AllData, i) {
  const RowData = [];
  const NewNumber = Number(i) + 1;
  if (NewNumber >= 10) {
    const FirstData = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${Trigger_SheetID}`,
      range: `${NewNumber}!Q10:Q${Sheet_LastRow}`
    });
    // シートにデータが入っていなければ 10行目から記載する
    if (FirstData == undefined) {
      RowData[0] = 10;
    }
    // シートにデータが入っていれば 次の空白行から記載する
    if (FirstData != undefined) {
      RowData[0] = FirstData.length + 10;
    }
    await SetSheetValues(
      Trigger_SheetID,
      `${NewNumber}`,
      `B${RowData[0]}`,
      `AE${Sheet_LastRow}`,
      Series_AllData
    );
    return;
  }
  if (NewNumber <= 9) {
    const FirstData = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${Trigger_SheetID}`,
      range: `0${NewNumber}!Q10:Q${Sheet_LastRow}`
    });
    // シートにデータが入っていなければ 10行目から記載する
    if (FirstData == undefined) {
      RowData[0] = 10;
    }
    // シートにデータが入っていれば 次の空白行から記載する
    if (FirstData != undefined) {
      RowData[0] = FirstData.length + 10;
    }
    await SetSheetValues(
      Trigger_SheetID,
      `0${NewNumber}`,
      `B${RowData[0]}`,
      `AE${Sheet_LastRow}`,
      Series_AllData
    );
  }
}

// RPAトリガーシートに 転記完了と記載する関数
async function PasteWorkEnd(TriggerSheetRow) {
  await SetSheetValues(
    Trigger_SheetID,
    `ビルボードRPAトリガー`,
    `D${TriggerSheetRow}`,
    `D${TriggerSheetRow}`,
    [['転記完了']]
  );
}

// バルクアップするため、スプレッドシートからデータを取得する関数
async function GetBulkUpData(firstdata, LastRow, i) {
  const NewNumber = Number(i) + 1;
  // Z列の 対応フラグ と 最終行数 を取得する
  if (NewNumber >= 10) {
    const FirstData = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${Trigger_SheetID}`,
      range: `${NewNumber}!Z10:Z${Sheet_LastRow}`
    });
    if (FirstData == undefined) {
      await RPA.Logger.info('このシートは データ0です.取得できません');
    }
    if (FirstData != undefined) {
      LastRow[0] = Number(FirstData.length) + 9;
      firstdata[0] = FirstData[0][0];
      return;
    }
  }
  if (NewNumber <= 9) {
    const FirstData = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${Trigger_SheetID}`,
      range: `0${NewNumber}!Z10:Z${Sheet_LastRow}`
    });
    if (FirstData == undefined) {
      await RPA.Logger.info('このシートは データ0です.取得できません');
    }
    if (FirstData != undefined) {
      LastRow[0] = Number(FirstData.length) + 9;
      firstdata[0] = FirstData[0][0];
      return;
    }
  }
}

// バルクアップの 確定ボタン を押す処理
async function BulkUpApply(i) {
  await RPA.sleep(3000);
  await RPA.WebBrowser.refresh();
  await RPA.sleep(1000);
  const LinkList = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({
      className: 'btn btn-sm btn-link'
    }),
    5000
  );
  const NewNumber = Number(i) + 1;
  for (let v in LinkList) {
    const LinkText = await LinkList[v].getText();
    if (NewNumber >= 10) {
      if (`${Trigger_SheetID}(${NewNumber})` == LinkText) {
        await RPA.WebBrowser.mouseClick(LinkList[v]);
        await RPA.Logger.info(LinkText);
        await RPA.Logger.info('バルクアップ一覧 一致しました.確定へ進みます');
        break;
      }
    }
    if (NewNumber <= 9) {
      if (`${Trigger_SheetID}(0${NewNumber})` == LinkText) {
        await RPA.WebBrowser.mouseClick(LinkList[v]);
        await RPA.Logger.info(LinkText);
        await RPA.Logger.info('バルクアップ一覧 一致しました.確定へ進みます');
        break;
      }
    }
  }
  const ErrorList = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({
      className: 'col-sm-1'
    }),
    5000
  );
  for (let i in ErrorList) {
    const ErrorText = await ErrorList[i].getText();
    if (ErrorText == 'あり') {
      await RPA.Logger.info(
        'Adxバルクアップにてエラーメッセージが有りました. スプレッドシートの見直しをお願いします'
      );
      // あえてエラーを出して処理を停止させる
      const errorele = await RPA.WebBrowser.findElementById('ooooo');
    }
  }
  const ApplyButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({
      className: 'btn btn-sm btn-success'
    }),
    5000
  );
  if (ApplyButton.length > 0) {
    await RPA.Logger.info('確定ボタンありました. 実行します');
    await RPA.sleep(1000);
    await RPA.WebBrowser.mouseClick(ApplyButton[0]);
  }
  await RPA.sleep(10000);
}
