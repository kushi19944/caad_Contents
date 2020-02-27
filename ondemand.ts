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

(async () => {
  try {
    // デバッグログを最小限にする
    RPA.Logger.level = 'INFO';
    await RPA.Google.authorize({
      //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
      refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
      tokenType: 'Bearer',
      expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
    });
    let firstdate = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName1}!C7:AB350`
    });
    let data1 = [];
    let data2 = [];
    let data3 = [];
    let data4 = [];
    let data5 = [];
    let data6 = [];
    let data7 = [];
    let data8 = [];
    let data9 = [];
    let data10 = [];
    let AllData = [
      data1,
      data2,
      data3,
      data4,
      data5,
      data6,
      data7,
      data8,
      data9,
      data10
    ];
    let moji = 1;
    for (let v in AllData) {
      let word = moji + '段目';
      //指定した文字列がある行だけ抽出する
      for (let i = 1; i < firstdate.length + 1; i++) {
        var date = firstdate[i - 1][0] + '';
        const judge = date.includes(word);
        if (judge == true) {
          AllData[v].push([
            firstdate[i - 1][0],
            firstdate[i - 1][3],
            firstdate[i - 1][2]
          ]);
        }
      }
      moji += 1;
    }
    //RPA.Logger.info(AllData)
    RPA.Logger.info(data1);

    //Adxにログインする
    await RPA.WebBrowser.get(Ondemand_PageURL);
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
    RPA.WebBrowser.sendKeys(UserID, [Ondemand_UserID]);
    const UserPass = await RPA.WebBrowser.findElementByXPath(
      '//*[@id="password"]'
    );
    RPA.WebBrowser.sendKeys(UserPass, [Ondemand_UserPW]);
    const LoginButton2 = await RPA.WebBrowser.findElementByXPath(
      '/html/body/div/div[2]/div/form/div[6]/a'
    );
    await RPA.WebBrowser.mouseClick(LoginButton2);
    await RPA.sleep(1500);

    //アイテムグループ(特集)の各ページへ飛ぶ
    const ItemGroupUrl = [551, 552, 553, 554, 555, 556, 557, 558, 559, 560];
    for (let i in ItemGroupUrl) {
      let ItemUrl = process.env.Ondemand_URL_1 + ItemGroupUrl[i];
      await RPA.WebBrowser.get(ItemUrl);
      //アイテムグループ(特集)を消す
      let DeleteButton = await RPA.WebBrowser.wait(
        RPA.WebBrowser.Until.elementLocated({
          xpath:
            '//*[@id="root"]/div/section/section/div/div[2]/form/div[8]/div/div/div[2]/div[1]/span/span[2]/button'
        }),
        5000
      );
      for (let i = 0; i < 50; i++) {
        try {
          await RPA.WebBrowser.scrollTo({
            xpath:
              '//*[@id="root"]/div/section/section/div/div[2]/form/div[8]/div/div/div[2]/div[1]/span/span[2]/button'
          });
          await RPA.WebBrowser.mouseClick(DeleteButton);
        } catch {
          break;
        }
        await RPA.sleep(50);
      }
      //アイテム(特集)を設定する
      let ItemInput = await RPA.WebBrowser.findElementByClassName(
        'select2-search__field'
      );
      let AppendButton = await RPA.WebBrowser.findElementByXPath(
        '//*[@id="root"]/div/section/section/div/div[2]/form/div[9]/div/button'
      );
      //リスト一覧が表示されるまでクリックループする
      while (
        (await RPA.WebBrowser.findElements('#select2-tmpGroups-results'))
          .length < 1
      ) {
        await RPA.WebBrowser.mouseClick(
          RPA.WebBrowser.findElement('.select2-search__field')
        );
        RPA.sleep(100);
      }
      const tmpListElements = await RPA.WebBrowser.findElements(
        '#select2-tmpGroups-results li'
      );
      const tmpList = await Promise.all(
        tmpListElements.map(async elm => await elm.getText())
      );
      for (let j in AllData[i]) {
        let ItemData = (await AllData[i][j][1]) + '';
        //リストとスプレッドシートの特集が一致しない場合、処理を飛ばす
        if (tmpList.indexOf(ItemData) === -1) {
          continue;
        }
        //特集を追加する処理
        await RPA.WebBrowser.scrollTo({
          xpath:
            '//*[@id="root"]/div/section/section/div/div[2]/form/div[9]/div/span/span/span[1]/span/ul/li/input'
        });
        //await RPA.WebBrowser.sendKeys(ItemInput,[ItemData]);
        await RPA.WebBrowser.driver.executeScript(
          `document.getElementsByClassName('select2-search__field')[0].value = '${ItemData}'`
        );
        await RPA.sleep(40);
        await RPA.WebBrowser.sendKeys(ItemInput, [RPA.WebBrowser.Key.SPACE]);
        await RPA.sleep(40);
        await RPA.WebBrowser.sendKeys(ItemInput, [
          RPA.WebBrowser.Key.BACK_SPACE
        ]);
        await RPA.sleep(40);
        await RPA.WebBrowser.sendKeys(ItemInput, [RPA.WebBrowser.Key.ENTER]);
        await RPA.sleep(60);
        await RPA.WebBrowser.mouseClick(AppendButton);
      }
      //特集を更新する処理
      const KoushinButton = await RPA.WebBrowser.findElementByXPath(
        '//*[@id="root"]/div/section/section/div/div[3]/div/button[1]'
      );
      await RPA.WebBrowser.mouseClick(KoushinButton);
      await RPA.sleep(2500);

      //モジュール一覧ページへ飛び、アイテムグループIDを取得する
      await RPA.WebBrowser.get(
        process.env.Ondemand_URL_2 +
          ItemGroupUrl[i] +
          '?sc=eyJjb2RlIjoiIn0=&ti=eyJwIjoxLCJzcHAiOjUwLCJzIjoiIiwic24iOiIiLCJzbyI6IiJ9'
      );
      await RPA.WebBrowser.wait(
        RPA.WebBrowser.Until.elementLocated({
          className: 'btn btn-sm btn-warning btn-block'
        }),
        5000
      );
      const ModuleCount = await RPA.WebBrowser.findElementsByClassName(
        'btn btn-sm btn-warning btn-block'
      );
      for (let k = 0; k < ModuleCount.length; k++) {
        await RPA.WebBrowser.get(
          process.env.Ondemand_URL_2 +
            ItemGroupUrl[i] +
            '?sc=eyJjb2RlIjoiIn0=&ti=eyJwIjoxLCJzcHAiOjUwLCJzIjoiIiwic24iOiIiLCJzbyI6IiJ9'
        );
        await RPA.WebBrowser.wait(
          RPA.WebBrowser.Until.elementLocated({
            className: 'btn btn-sm btn-warning btn-block'
          }),
          5000
        );
        const MakeXpath = Number(k + 1);
        try {
          /*
          const ErrorArateEle = await RPA.WebBrowser.findElementsByXPath(
            '//*[@id="root"]/div/section/section/div/div[2]'
          );
          const ErrorMessege = await Promise.all(
            ErrorArateEle.map(async elm => await elm.getText())
          );
          */
          const ErrorArateEle = await RPA.WebBrowser.findElementByXPath(
            '//*[@id="root"]/div/section/section/div/div[2]'
          );
          const ErrorMessege = await ErrorArateEle.getText();
          var ErrorJudge = ErrorMessege.includes(
            '紐づくグループのパラメータを編集してください。'
          );
          await RPA.Logger.info(ErrorJudge);
        } catch {
          RPA.Logger.info(ErrorJudge);
        }
        //エラーメッセージありの場合は下記で処理する
        if (ErrorJudge == true) {
          await RPA.Logger.info('エラーアラート有り');
          await RPA.WebBrowser.wait(
            RPA.WebBrowser.Until.elementLocated({
              xpath:
                '//*[@id="root"]/div/section/section/div/div[3]/div/div[2]/div[2]/table/tbody/tr[' +
                MakeXpath +
                ']/td[1]'
            }),
            5000
          );
          //*[@id="root"]/div/section/section/div/div[3]/div/div[2]/div[2]/table/tbody/tr[1]/td[1]
          await RPA.WebBrowser.wait(
            RPA.WebBrowser.Until.elementLocated({
              xpath:
                '//*[@id="root"]/div/section/section/div/div[3]/div/div[2]/div[2]/table/tbody/tr[' +
                MakeXpath +
                ']/td[3]'
            }),
            5000
          );
          var MakeItemID = await RPA.WebBrowser.findElementByXPath(
            '//*[@id="root"]/div/section/section/div/div[3]/div/div[2]/div[2]/table/tbody/tr[' +
              MakeXpath +
              ']/td[1]'
          );
          var MakeItemName = await RPA.WebBrowser.findElementByXPath(
            '//*[@id="root"]/div/section/section/div/div[3]/div/div[2]/div[2]/table/tbody/tr[' +
              MakeXpath +
              ']/td[3]'
          );
          var ItemGroupID = await MakeItemID.getText();
          var ItemName = await MakeItemName.getText();
        }
        //エラーメッセージなしの場合は下記で処理する
        if (ErrorJudge == false) {
          await RPA.Logger.info('エラーアラート無し');
          await RPA.WebBrowser.wait(
            RPA.WebBrowser.Until.elementLocated({
              xpath:
                '//*[@id="root"]/div/section/section/div/div[2]/div/div[2]/div[2]/table/tbody/tr[' +
                MakeXpath +
                ']/td[1]'
            }),
            5000
          );
          await RPA.WebBrowser.wait(
            RPA.WebBrowser.Until.elementLocated({
              xpath:
                '//*[@id="root"]/div/section/section/div/div[2]/div/div[2]/div[2]/table/tbody/tr[' +
                MakeXpath +
                ']/td[3]'
            }),
            5000
          );
          var MakeItemID = await RPA.WebBrowser.findElementByXPath(
            '//*[@id="root"]/div/section/section/div/div[2]/div/div[2]/div[2]/table/tbody/tr[' +
              MakeXpath +
              ']/td[1]'
          );
          var MakeItemName = await RPA.WebBrowser.findElementByXPath(
            '//*[@id="root"]/div/section/section/div/div[2]/div/div[2]/div[2]/table/tbody/tr[' +
              MakeXpath +
              ']/td[3]'
          );
          var ItemGroupID = await MakeItemID.getText();
          var ItemName = await MakeItemName.getText();
        }

        //セグメント編集ページに飛ぶ
        await RPA.WebBrowser.get(
          process.env.Ondemand_URL_3 + ItemGroupUrl[i] + '/' + ItemGroupID
        );
        await RPA.sleep(300);
        var DeleteButtonJudge = false;
        try {
          var SegmentDeleteButton = await RPA.WebBrowser.wait(
            RPA.WebBrowser.Until.elementLocated({
              xpath:
                '//*[@id="root"]/div/section/section/div/div[2]/form/div[3]/div/div/div[2]/div[1]/span/span[2]/button'
            }),
            3000
          );
          DeleteButtonJudge = true;
          RPA.Logger.info('セグメント削除ボタン有り');
        } catch {
          DeleteButtonJudge = false;
          RPA.Logger.info('セグメント削除ボタン無し');
        }
        if (DeleteButtonJudge == true) {
          //セグメントを消す
          for (let button = 0; button < 10; button++) {
            try {
              //await RPA.WebBrowser.scrollTo({xpath:'//*[@id="root"]/div/section/section/div/div[2]/form/div[3]/div/div/div[2]/div[1]/span/span[2]/button'})
              await RPA.WebBrowser.mouseClick(SegmentDeleteButton);
            } catch {
              break;
            }
            await RPA.sleep(55);
          }
        }

        //セグメントを追加する処理
        RPA.Logger.info(AllData[i][k]);
        const SegmemtInput = await RPA.WebBrowser.findElementByXPath(
          '//*[@id="root"]/div/section/section/div/div[2]/form/div[4]/div/span/span/span[1]/span/ul/li/input'
        );
        const SegmemtInputButton = await RPA.WebBrowser.findElementByXPath(
          '//*[@id="root"]/div/section/section/div/div[2]/form/div[4]/div/button'
        );
        const Datas = AllData[i][k][2] + '';
        let wordjudge = Datas.includes(',');
        let SegmentData = Datas;
        if (wordjudge == true) {
          var SegmentDatas = Datas.split(',');
          for (let x in SegmentDatas) {
            await RPA.WebBrowser.scrollTo({
              xpath:
                '//*[@id="root"]/div/section/section/div/div[2]/form/div[4]/div/span/span/span[1]/span/ul/li/input'
            });
            await RPA.WebBrowser.sendKeys(SegmemtInput, [SegmentDatas[x]]);
            await RPA.sleep(50);
            await RPA.WebBrowser.sendKeys(SegmemtInput, [
              RPA.WebBrowser.Key.ENTER
            ]);
            await RPA.sleep(90);
            await RPA.WebBrowser.mouseClick(SegmemtInputButton);
          }
        } else {
          await RPA.WebBrowser.scrollTo({
            xpath:
              '//*[@id="root"]/div/section/section/div/div[2]/form/div[4]/div/span/span/span[1]/span/ul/li/input'
          });
          await RPA.WebBrowser.sendKeys(SegmemtInput, [SegmentData]);
          await RPA.sleep(50);
          await RPA.WebBrowser.sendKeys(SegmemtInput, [
            RPA.WebBrowser.Key.ENTER
          ]);
          await RPA.sleep(90);
          await RPA.WebBrowser.mouseClick(SegmemtInputButton);
        }

        //オペレーションが入稿物以外なら下記処理を実行する
        await RPA.WebBrowser.scrollTo({
          xpath: '//*[@id="select2-operationId-container"]'
        });
        await RPA.sleep(110);
        const text1 = await RPA.WebBrowser.findElementByXPath(
          '//*[@id="select2-operationId-container"]'
        );
        const text2 = await text1.getText();
        RPA.Logger.info(text2);
        if (text2 != '入稿物') {
          //const OpeContena = await RPA.WebBrowser.findElementByXPath('//*[@id="root"]/div/section/section/div/div[2]/form/div[8]/div[1]/div/span[1]/span/span[1]/span')
          const OpeInput = await RPA.WebBrowser.findElementsByClassName(
            'select2-selection select2-selection--single'
          );
          await RPA.WebBrowser.mouseClick(OpeInput[0]);
          await RPA.sleep(300);
          await RPA.WebBrowser.sendKeys(OpeInput[0], ['入稿物']);
          await RPA.sleep(100);
          await RPA.WebBrowser.sendKeys(OpeInput[0], [
            RPA.WebBrowser.Key.ENTER
          ]);
          await RPA.sleep(800);
          // 取得条件 の Input要素
          const Jouken = await RPA.WebBrowser.findElementByXPath(
            '//*[@id="queryParams.take_n"]'
          );
          await Jouken.clear();
          await RPA.sleep(200);
          await RPA.WebBrowser.sendKeys(Jouken, ['20']);
          await RPA.Logger.info('取得条件を 20 に設定しました');
          // ソート順　をクリックして priority を入力する
          const SortList = await RPA.WebBrowser.findElementsByClassName(
            'select2-selection select2-selection--single'
          );
          await RPA.WebBrowser.mouseClick(SortList[1]);
          await RPA.sleep(600);
          const SortInput = await RPA.WebBrowser.findElementsByClassName(
            'select2-search__field'
          );
          await RPA.WebBrowser.sendKeys(SortInput[2], ['priority']);
          await RPA.sleep(100);
          await RPA.WebBrowser.sendKeys(SortInput[2], [
            RPA.WebBrowser.Key.ENTER
          ]);
          await RPA.Logger.info('ソート順を priority に設定しました');
          await RPA.sleep(200);
          // 最低保証件数 を入力する
          //const HosyouJouken = await RPA.WebBrowser.findElementByXPath('//*[@id="queryParams.ensure_min"]');
          const HosyouKensuu = await RPA.WebBrowser.findElementById(
            'queryParams.ensure_min'
          );
          await RPA.WebBrowser.sendKeys(HosyouKensuu, ['4']);
          await RPA.Logger.info('最低保証件数を 4 に設定しました');
        }
        //セグメント更新する処理
        const SegmentKoushinButton = await RPA.WebBrowser.findElementByXPath(
          '//*[@id="root"]/div/section/section/div/div[3]/div/button[1]'
        );
        await RPA.WebBrowser.mouseClick(SegmentKoushinButton);
        await RPA.sleep(1800);
      }
    }
  } catch (error) {
    await RPA.WebBrowser.takeScreenshot();
    RPA.SystemLogger.error(error);
  } finally {
    await RPA.WebBrowser.quit();
  }
})();
