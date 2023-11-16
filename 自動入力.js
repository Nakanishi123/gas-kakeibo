const GET_DAY = 1;//何日前までの物を取得するか

function AAA() {
  var Now = new Date();
  var NowMonth = Now.getMonth() + 1;

  var Yesterday = new Date();
  Yesterday.setDate(Now.getDate() - 1);
  var YesterdayMonth = Yesterday.getMonth() + 1;
  var sheetname = String(YesterdayMonth) + "月"

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetname);

  var JapanNet = Getmail("Ｖｉｓａデビット", Now, "customer2@cc.paypay-bank.co.jp")
  var JapanNet_Hiki = Getmail("お引き落としのご連絡", Now, "customer2@cc.paypay-bank.co.jp")
  var sonybank = Getmail("［Sony Bank WALLET ］Visaデビット", Now, "banking@sonybank.net")
  var rakuten = Getmail("◆デビットカードご利用通知メール", Now, "service@ac.rakuten-bank.co.jp")
  var rakuten_card = Getmail("【速報版】カード利用のお知らせ(本人ご利用分)", Now, "info@mail.rakuten-card.co.jp")
  var rakuten_pay = Getmail("楽天", Now, "no-reply@pay.rakuten.co.jp")

  //最終行取得(900,4)=(D900)から上に検索
  var lastRow = sheet.getRange(900, 4).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

  var Fnoyatsu = lastRow - 4
  if (Fnoyatsu > 0) {
    var IDs = sheet.getRange(5, 7, Fnoyatsu, 1).getValues();
  } else {
    var IDs = [];
  }
  Logger.log(IDs)
  var NIDs = []
  for (var i = 0; i < IDs.length; i++) {//配列で与えられるので整形
    NIDs.push(IDs[i][0] + '');//文字列として認識させるため
  }

  //スプシの形に合うように整形
  JapanNet = Get_JN_VISA(JapanNet, NIDs)
  JapanNet_Hiki = Get_JN_Hiki(JapanNet_Hiki, NIDs)
  sonybank = Get_SonyBankWallet(sonybank, NIDs)
  rakuten = Get_rakuten(rakuten, NIDs)
  rakuten_card = Get_rakuten_card(rakuten_card, NIDs)
  rakuten_pay = Get_rakuten_pay(rakuten_pay, NIDs)

  lastRow++
  lastRow = lastRow + SetMail(JapanNet, sheet, lastRow) //SetMailは配列の長さを出力する
  lastRow = lastRow + SetMail(JapanNet_Hiki, sheet, lastRow)
  lastRow = lastRow + SetMail(sonybank, sheet, lastRow)
  lastRow = lastRow + SetMail(rakuten, sheet, lastRow)
  lastRow = lastRow + SetMail(rakuten_card, sheet, lastRow)
  lastRow = lastRow + SetMail(rakuten_pay, sheet, lastRow)

}

function Get_JN_VISA(val, Juhuku) {
  re = []
  for (var i = 0; i < val.length; i++) {
    if (Juhuku.indexOf(val[i][4]) === -1) {
      var Kingaku = val[i][3].match(/お引落金額：(.{1,16})円/)
      var Kameiten = val[i][3].match(/加盟店名：(.+)/)
      Logger.log([Kingaku, Kameiten, val]);
      try {
        re.push([Kameiten[1], Kingaku[1], val[i][0], "JNB VISAデビット", '', val[i][4]])
      } catch (error) {
        console.log("エラー内容：" + error);
      }
    }
  }
  return re
}

function Get_JN_Hiki(val, Juhuku) {
  re = []
  for (var i = 0; i < val.length; i++) {
    if (Juhuku.indexOf(val[i][4]) === -1) {
      var Kingaku = val[i][3].match(/お引落金額：(.{1,16})円/)
      Logger.log([Kingaku, val]);
      re.push(["JNB 引き落とし", Kingaku[1], val[i][0], "JNB 引き落とし", '', val[i][4]])
    }
  }
  return re
}

function Get_SonyBankWallet(val, Juhuku) {
  re = []
  for (var i = 0; i < val.length; i++) {
    if (Juhuku.indexOf(val[i][4]) === -1) {
      var Kingaku = val[i][3].match(/ご利用金額（※）：(.{1,16})円/)
      var Kameiten = val[i][3].match(/ご利用加盟店\s*：(.+)/)
      Logger.log([Kingaku, Kameiten, val]);
      if (Kingaku != null && Kameiten != null) {
        re.push([Kameiten[1], Kingaku[1], val[i][0], "SonyBankWallet", '', val[i][4]])
      }
    }
  }
  return re
}

function Get_rakuten(val, Juhuku) {
  re = []
  for (var i = 0; i < val.length; i++) {
    if (Juhuku.indexOf(val[i][4]) === -1) {
      var Kingaku = val[i][3].match(/口座引落分：(.{1,16})円/)
      Logger.log([Kingaku, val]);
      re.push(["楽天 デビット", Kingaku[1], val[i][0], "楽天 デビット", '', val[i][4]])
    }
  }
  return re
}

function Get_rakuten_card(val, Juhuku) {
  re = []
  for (var i = 0; i < val.length; i++) {
    if (Juhuku.indexOf(val[i][4]) === -1) {
      var Kingaku = [...val[i][3].matchAll(/利用金額: (.{1,16}) 円/g)]
      var day = [...val[i][3].matchAll(/■利用日: (\d+\/\d+\/\d+)/g)]
      Logger.log([Kingaku, day, val]);
      for (var j = 0; j < Kingaku.length; j++) {
        re.push(["楽天カード", Kingaku[j][1], day[j][1], "楽天カード", '', val[i][4]])
      }
    }
  }
  return re
}

function Get_rakuten_pay(val, Juhuku) {
  re = []
  for (var i = 0; i < val.length; i++) {
    if (Juhuku.indexOf(val[i][4]) === -1) {
      var Kingaku = val[i][3].match(/決済総額\s*(.{1,16})円/)
      var Kameiten = val[i][3].match(/ご利用店舗\s*(.+)/)

      Logger.log([Kingaku, Kameiten, val]);
      if (Kingaku != null && Kameiten != null) {
        re.push([Kameiten[1], Kingaku[1], val[i][0], "楽天ペイ", '', val[i][4]])
      }
    }
  }
  return re
}



function SetMail(val, sheet, lastRow) {
  if (val.length != 0) {
    var colleng = val.length;
    var range = sheet.getRange(lastRow, 2, colleng, 6);
    range.setValues(val);
  }
  return val.length
}

function Getmail(keyword, date, from) {
  SearchQuery = "";
  if (keyword != undefined) {
    SearchQuery = SearchQuery + "subject:" + keyword + " ";
  }
  if (from != undefined) {
    SearchQuery = SearchQuery + "from:" + from + " ";
  }
  Logger.log(SearchQuery)
  if (date != undefined) {
    var adate = new Date();
    adate.setDate(date.getDate() - GET_DAY);
    adate = Utilities.formatDate(adate, 'JST', 'yyyy/M/d');
    var bdate = new Date();
    bdate.setDate(date.getDate());
    bdate = Utilities.formatDate(bdate, 'JST', 'yyyy/M/d');
    mailafter = "after:" + adate + " ";
    mailbefore = "before:" + bdate + " ";
    SearchQuery = SearchQuery + mailafter + mailbefore;
  }

  re = GmailApp.search(SearchQuery);
  var reM = GmailApp.getMessagesForThreads(re);

  var JNM = []
  for (var i = 0; i < reM.length; i++) {
    for (var j = 0; j < reM[i].length; j++) {
      var date = reM[i][j].getDate();
      var reply = reM[i][j].getReplyTo();
      var subj = reM[i][j].getSubject();
      var body = reM[i][j].getPlainBody().slice(0, 500);
      var GettedId = reM[i][j].getId();
      JNM.push([date, reply, subj, body, GettedId])
    }
  }
  Logger.log(JNM)
  return JNM
}
