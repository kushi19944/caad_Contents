import RPA from 'ts-rpa';
import { url } from 'inspector';
import { Driver } from 'selenium-webdriver/chrome';
import { Session, Key, WebDriver } from 'selenium-webdriver';

// 読み込みする スプレッドシートID と シート名 の記載
const SSID = process.env.Ondemand_SheetID;
const SSName1 = process.env.Ondemand_SheetName;
const Ondemand_UserID = process.env.Ondemand_LoginUserID;
const Ondemand_UserPW = process.env.Ondemand_LoginUserPW;
const Ondemand_PageURL = process.env.Ondemand_URL;
// Slack (p1)Bot 通知トークン・チャンネル
const BotToken = process.env.AbemaTV_hubot_Token;
const BotChannel = process.env.Contents_Slack_Channel;
//const BotChannel = process.env.RPA_Test_Channel;

const SlackText = ['特集設定 完了しました A列のフラグ確認をお願いします'];
const ErrorText = [];

async function Test() {
  // デバッグログを最小限にする
  RPA.Logger.level = 'INFO';
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10),
  });
  let firstdata = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!A1:G1000`,
  });
  const Data1 = [];
  const Data2 = [];
  const Data3 = [];
  const Data4 = [];
  const Data5 = [];
  const Data6 = [];
  const Data7 = [];
  const Data8 = [];
  const Data9 = [];
  const Data10 = [];
  for (let i in firstdata) {
    const SumNunber = Number(i) + 1;
    // スプレッドシートの行数を追加する
    firstdata[i].push(`${SumNunber}`);
    if (firstdata[i][3] == '1段目') {
      Data1.push(firstdata[i]);
    }
    if (firstdata[i][3] == '2段目') {
      Data2.push(firstdata[i]);
    }
    if (firstdata[i][3] == '3段目') {
      Data3.push(firstdata[i]);
    }
    if (firstdata[i][3] == '4段目') {
      Data4.push(firstdata[i]);
    }
    if (firstdata[i][3] == '5段目') {
      Data5.push(firstdata[i]);
    }
    if (firstdata[i][3] == '6段目') {
      Data6.push(firstdata[i]);
    }
    if (firstdata[i][3] == '7段目') {
      Data7.push(firstdata[i]);
    }
    if (firstdata[i][3] == '8段目') {
      Data8.push(firstdata[i]);
    }
    if (firstdata[i][3] == '9段目') {
      Data9.push(firstdata[i]);
    }
    if (firstdata[i][3] == '10段目') {
      Data10.push(firstdata[i]);
    }
  }
  try {
    await SlackPost('特集設定 RPA開始します');
    await AdxLogin();
    await LoginJudge_function();
    await TokusyuuSetting(Data1, '551');
    await TokusyuuSetting(Data2, '552');
    await TokusyuuSetting(Data3, '553');
    await TokusyuuSetting(Data4, '554');
    await TokusyuuSetting(Data5, '555');
    await TokusyuuSetting(Data6, '556');
    await TokusyuuSetting(Data7, '557');
    await TokusyuuSetting(Data8, '558');
    await TokusyuuSetting(Data9, '559');
    await TokusyuuSetting(Data10, '560');
  } catch (error) {
    SlackText[0] = '特集設定 エラー発生です\n@kushi_makoto 確認してください';
    await RPA.WebBrowser.takeScreenshot();
    await RPA.SystemLogger.error(error);
  } finally {
    await SlackPost(SlackText[0]);
    await RPA.WebBrowser.quit();
  }
}

Test();

// 特集◯段目 の設定
async function TokusyuuSetting(Data, URL) {
  await RPA.WebBrowser.get(process.env.Ondemand_URL_1 + URL);
  // アイテムグループ(特集)を消す
  await ItemDelete();
  const NewData = [];
  //アイテム(特集)を設定する
  await ItemSetting(Data, NewData);
  // モジュール／アイテムグループ一覧 にて 編集ボタンをおす
  for (let i in NewData) {
    if (NewData[i][0] != '') {
      continue;
    }
    await PageMoving(i, URL);
    await RPA.Logger.info(NewData[i]);
    // セグメント登録
    await EditSegments(NewData[i]);
  }
}

//Adxにログインする
async function AdxLogin() {
  try {
    await RPA.WebBrowser.get(Ondemand_PageURL);
    await RPA.sleep(5000);
    // ログインページ のボタンをクリック
    try {
      await RPA.Logger.info('ログインボタン検索');
      const Button1 = await RPA.WebBrowser.wait(
        RPA.WebBrowser.Until.elementLocated({ className: 'btn btn-primary' }),
        5000
      );
      await RPA.sleep(1000);
      await RPA.WebBrowser.mouseClick(Button1);
    } catch {
      RPA.Logger.info(`ログインボタン見つかりません`);
    }
    await RPA.sleep(2000);
    const UserID = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({ xpath: '//*[@id="username"]' }),
      5000
    );
    RPA.WebBrowser.sendKeys(UserID, [Ondemand_UserID]);
    const UserPass = await RPA.WebBrowser.findElementByXPath(
      '//*[@id="password"]'
    );
    RPA.WebBrowser.sendKeys(UserPass, [Ondemand_UserPW]);
    const LoginButton2 = await RPA.WebBrowser.findElementByXPath(
      '/html/body/div/div[2]/div/form/div[6]/a'
    );
    await RPA.WebBrowser.mouseClick(LoginButton2);
    await RPA.sleep(4000);
    // 社員情報にアクセス 許可かどうか聞かれることがあるためスクショ取って確認できるように
    await RPA.WebBrowser.takeScreenshot();
  } catch (ErrorMess) {
    ErrorText[0] = ErrorMess;
  }
}

// きちんとログイン完了しているか確認
async function LoginJudge_function() {
  try {
    await RPA.Logger.info('ログインボタン再度検索');
    const Button1 = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({ className: 'btn btn-primary' }),
      5000
    );
    await RPA.sleep(1000);
    await RPA.WebBrowser.mouseClick(Button1);
    await RPA.sleep(5000);
    await RPA.Logger.info(`再度ログイン完了しました`);
  } catch {
    RPA.Logger.info(`再度ログインせずにログイン完了`);
  }
}

// アイテムグループ(特集)を消す関数
async function ItemDelete() {
  await RPA.sleep(4000);
  let PanelHeading = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className: 'panel-heading',
    }),
    15000
  );
  for (let i = 0; i < 100; i++) {
    try {
      await RPA.WebBrowser.scrollTo({
        selector: `#root > div > section > section > div > div.form-contens > form > div:nth-child(8) > div > div > div.panel-heading`,
      });
      // 特集の削除ボタンクリック
      await RPA.WebBrowser.driver.executeScript(
        'document.getElementsByClassName(`list-group`)[0].children[0].children[0].children[1].children[0].click()'
      );
      await RPA.sleep(50);
    } catch {
      break;
    }
  }
}

//アイテム(特集)を設定する
async function ItemSetting(Data, NewData) {
  let ItemInput = await RPA.WebBrowser.findElementByClassName(
    'select2-search__field'
  );
  let AppendButton = await RPA.WebBrowser.findElementsByClassName(
    'btn btn-default'
  );
  //リスト一覧が表示されるまでクリックループする
  while (
    (await RPA.WebBrowser.findElements('#select2-tmpGroups-results')).length < 1
  ) {
    await RPA.WebBrowser.mouseClick(
      RPA.WebBrowser.findElement('.select2-search__field')
    );
    RPA.sleep(100);
  }
  //特集を追加する処理
  for (let i in Data) {
    await RPA.WebBrowser.scrollTo({
      xpath:
        '//*[@id="root"]/div/section/section/div/div[2]/form/div[9]/div/span/span/span[1]/span/ul/li/input',
    });
    await RPA.WebBrowser.driver.executeScript(
      `document.getElementsByClassName('select2-search__field')[0].value = '${Data[i][6]}'`
    );
    await RPA.sleep(50);
    await RPA.WebBrowser.sendKeys(ItemInput, [await RPA.WebBrowser.Key.SPACE]);
    await RPA.sleep(50);
    await RPA.WebBrowser.sendKeys(ItemInput, [
      await RPA.WebBrowser.Key.BACK_SPACE,
    ]);
    await RPA.sleep(300);
    // アイテムインプットにて、No results found の文字が出たら消す
    const NoResultsList = await RPA.WebBrowser.findElementsByClassName(
      'select2-results__option select2-results__message'
    );
    if (NoResultsList.length == 1) {
      await ItemInput.clear();
      await RPA.sleep(200);
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: `${SSID}`,
        range: `${SSName1}!A${Data[i][7]}:A${Data[i][7]}`,
        values: [[`【エラー】特集不一致`]],
      });
      await RPA.sleep(300);
      continue;
    }
    if (NoResultsList.length == 0) {
      // No results found じゃなければ NewDataに格納する
      NewData.push(Data[i]);
      await RPA.sleep(50);
      await RPA.WebBrowser.sendKeys(ItemInput, [
        await RPA.WebBrowser.Key.ENTER,
      ]);
      await RPA.sleep(100);
      await RPA.WebBrowser.mouseClick(AppendButton[0]);
    }
  }
  //特集を更新する処理
  const KoushinButton = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="root"]/div/section/section/div/div[3]/div/button[1]'
  );
  await RPA.WebBrowser.mouseClick(KoushinButton);
  await RPA.sleep(2200);
}

// 編集ページ をおす
async function PageMoving(i, URL) {
  await RPA.WebBrowser.get(
    process.env.Ondemand_URL_2 +
      URL +
      '?sc=eyJjb2RlIjoiIn0=&ti=eyJwIjoxLCJzcHAiOjUwLCJzIjoiIiwic24iOiIiLCJzbyI6IiJ9'
  );
  const EditButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({
      className: 'btn btn-sm btn-warning btn-block',
    }),
    5000
  );
  // 指定したボタン が一番したに来るようにスクロールする(true)で一番上に来る
  await RPA.WebBrowser.driver.executeScript(
    `document.getElementsByClassName("btn btn-sm btn-warning btn-block")[${i}].scrollIntoView(false)`
  );
  await RPA.sleep(200);
  await RPA.WebBrowser.driver.executeScript(
    `document.getElementsByClassName('btn btn-sm btn-warning btn-block')[${i}].click()`
  );
}

// セグメントを編集する関数
async function EditSegments(NewData) {
  await RPA.sleep(300);
  var headerText = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({
      className: 'form-header',
    }),
    5000
  );
  var SegmentDeleteButton = await RPA.WebBrowser.findElementsByXPath(
    '//*[@id="root"]/div/section/section/div/div[2]/form/div[3]/div/div/div[2]/div[1]/span/span[2]/button'
  );
  // セグメントを消す
  if (SegmentDeleteButton.length >= 1) {
    for (let button = 0; button < 100; button++) {
      try {
        await RPA.WebBrowser.mouseClick(SegmentDeleteButton[0]);
      } catch {
        break;
      }
      await RPA.sleep(55);
    }
  }
  // セグメントを入力する
  const SegmemtInput = [];
  const InputList = await RPA.WebBrowser.findElementsByClassName(
    'select2-search__field'
  );
  const SegmemtInputButton = await RPA.WebBrowser.findElementsByClassName(
    'btn btn-default'
  );
  for (let i in InputList) {
    const Text = await InputList[i].getAttribute('placeholder');
    if (Text == 'セグメントコンビネーションはこちらから追加してください。') {
      SegmemtInput[0] = i;
      await RPA.Logger.info('セグメント Input ありました. 入力します');
      break;
    }
  }
  const Datas = NewData[5] + '';
  let wordjudge = Datas.includes(',');
  let SegmentData = Datas;
  if (wordjudge == true) {
    var SegmentDatas = Datas.split(',');
    for (let x in SegmentDatas) {
      await RPA.WebBrowser.scrollTo({
        xpath:
          '//*[@id="root"]/div/section/section/div/div[2]/form/div[4]/div/span/span/span[1]/span/ul/li/input',
      });
      await RPA.WebBrowser.sendKeys(InputList[SegmemtInput[0]], [
        SegmentDatas[x],
      ]);
      await RPA.sleep(50);
      await RPA.WebBrowser.sendKeys(InputList[SegmemtInput[0]], [
        RPA.WebBrowser.Key.ENTER,
      ]);
      await RPA.sleep(90);
      await RPA.WebBrowser.mouseClick(SegmemtInputButton[0]);
    }
  } else {
    await RPA.WebBrowser.scrollTo({
      xpath:
        '//*[@id="root"]/div/section/section/div/div[2]/form/div[4]/div/span/span/span[1]/span/ul/li/input',
    });
    await RPA.WebBrowser.sendKeys(InputList[SegmemtInput[0]], [SegmentData]);
    await RPA.sleep(50);
    await RPA.WebBrowser.sendKeys(InputList[SegmemtInput[0]], [
      RPA.WebBrowser.Key.ENTER,
    ]);
    await RPA.sleep(90);
    await RPA.WebBrowser.mouseClick(SegmemtInputButton[0]);
  }
  // 入稿物 の欄までスクロール
  await RPA.WebBrowser.scrollTo({
    xpath: '//*[@id="select2-operationId-container"]',
  });
  await RPA.sleep(110);
  const text1 = await RPA.WebBrowser.findElementById(
    'select2-operationId-container'
  );
  const text2 = await text1.getText();
  //オペレーションが入稿物以外なら下記処理を実行する
  if (text2 != '入稿物') {
    const LoopFlag = ['true'];
    for (let i = 0; i <= 3; i++) {
      try {
        const OpeClick = await RPA.WebBrowser.findElementsByClassName(
          'select2-selection select2-selection--single'
        );
        await OpeClick[0].click();
        await RPA.sleep(1000);
        const OptionList = await RPA.WebBrowser.wait(
          RPA.WebBrowser.Until.elementsLocated({
            className: `select2-results__option`,
          }),
          15000
        );
        for (let i in OptionList) {
          const Text = await OptionList[i].getText();
          if (Text == '入稿物') {
            await OptionList[i].click();
            LoopFlag[0] = 'false';
            break;
          }
        }
      } catch {}
      if (LoopFlag[0] == 'false') {
        break;
      }
      if (i == 3) {
        await RPA.Logger.info('3回エラーが出現しました.スキップします');
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName1}!A${NewData[7]}:A${NewData[7]}`,
          values: [[`【エラー】入稿物設定エラー`]],
        });
        await RPA.sleep(300);
        return;
      }
    }
    await RPA.sleep(1000);
    await RPA.WebBrowser.scrollTo({ xpath: `//*[@id="queryParams.take_n"]` });
    // 取得件数 の Input要素
    const Jouken = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        id: `queryParams.take_n`,
      }),
      5000
    );
    await Jouken.clear();
    await RPA.sleep(50);
    await RPA.WebBrowser.sendKeys(Jouken, ['20']);
    await RPA.sleep(100);
    await RPA.Logger.info('【取得件数】20');
    // ソート順 のオプションエラーが多く出たため、Elementエラーの際は3回ループ処理させる
    const SortLoopFlag = ['true'];
    for (let i = 0; i <= 3; i++) {
      try {
        // ソート順　をクリック
        const SortList = await RPA.WebBrowser.findElementsByClassName(
          'select2-selection select2-selection--single'
        );
        await RPA.WebBrowser.mouseClick(SortList[1]);
        await RPA.sleep(1000);
        const SortOption = await RPA.WebBrowser.wait(
          RPA.WebBrowser.Until.elementsLocated({
            className: `select2-results__option`,
          }),
          15000
        );
        for (let i in SortOption) {
          const SortText = await SortOption[i].getText();
          if (SortText == 'priority') {
            await SortOption[i].click();
            await RPA.Logger.info(`【ソート順】${SortText}`);
            SortLoopFlag[0] = 'false';
            break;
          }
        }
      } catch {}
      if (i == 3) {
        await RPA.Logger.info('3回エラーが出現しました.スキップします');
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName1}!A${NewData[7]}:A${NewData[7]}`,
          values: [[`【エラー】ソート順設定エラー`]],
        });
        await RPA.sleep(300);
        return;
      }
      if (SortLoopFlag[0] == 'false') {
        break;
      }
    }
    await RPA.sleep(200);
    // 最低保証件数 を入力する
    const HosyouKensuu = await RPA.WebBrowser.findElementById(
      'queryParams.ensure_min'
    );
    await HosyouKensuu.clear();
    await RPA.sleep(50);
    await RPA.WebBrowser.sendKeys(HosyouKensuu, ['4']);
    await RPA.sleep(200);
    await RPA.Logger.info('【最低保証件数】4');
  }
  // 取得件数が20以外なら 20に変更する
  const Jouken = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementById('queryParams.take_n').value`
  );
  if (Jouken != '20') {
    const Jouken_ele = await RPA.WebBrowser.findElementById(
      'queryParams.take_n'
    );
    await Jouken_ele.clear();
    await RPA.sleep(50);
    await RPA.WebBrowser.sendKeys(Jouken_ele, ['20']);
    await RPA.sleep(100);
    await RPA.Logger.info('【取得件数】20');
  }
  //セグメント更新する処理
  const SegmentKoushinButton = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="root"]/div/section/section/div/div[3]/div/button[1]'
  );
  await RPA.WebBrowser.mouseClick(SegmentKoushinButton);
  await RPA.sleep(300);
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!A${NewData[7]}:A${NewData[7]}`,
    values: [[`【完了】`]],
  });
  await RPA.sleep(100);
  try {
    const EditButton = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementsLocated({
        className: `btn btn-sm btn-warning btn-block`,
      }),
      3500
    );
  } catch {}
}

async function SlackPost(Text) {
  // 作業開始時にSlackへ通知する
  await RPA.Slack.chat.postMessage({
    channel: BotChannel,
    token: BotToken,
    text: `${Text}`,
    icon_emoji: ':snowman:',
    username: 'p1',
  });
}
