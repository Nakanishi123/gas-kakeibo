const BEFORE_DAY = 5; //何日前までの物を取得するか

class Data {
  constructor(name, price, date, category, object, id) {
    this.name = name;
    this.price = price;
    this.date = date;
    this.category = category;
    this.object = object;
    this.id = id;
  }

  toRow() {
    return [this.name, this.price, this.date, this.category, this.object, this.id];
  }
}

class Parser {
  constructor(category, subject, from, nameRegex, priceRegex, objectRegex, dateRegex, nameDefault) {
    this.category = category;
    this.subject = subject;
    this.from = from;
    this.nameRegex = nameRegex;
    this.priceRegex = priceRegex;
    this.objectRegex = objectRegex;
    this.dateRegex = dateRegex;
    this.nameDefault = nameDefault;
  }
}

function autofill() {
  const alreadyIds = getAlreadyId();
  const parsers = [
    new Parser( //d払い
      'd払い',
      '【d払い】決済完了のお知らせ(自動配信)',
      'docomo',
      /【加盟店名】\s*(.*)/,
      /【ご利用代金】\D*((\d|.|,)+)/,
      '',
      /【決済日時】\s*(.*)/,
      'd払い'
    ),
    new Parser( //楽天ペイ
      '楽天ペイ',
      '楽天ペイアプリご利用内容確認メール',
      'no-reply@pay.rakuten.co.jp',
      /ご利用店舗\s*(.+)/,
      /決済総額\s*(.{1,16})円/,
      '',
      /ご利用日時\s*(.*)/,
      '楽天ペイ'
    ),
    new Parser( //楽天カード
      '楽天カード',
      '【速報版】カード利用のお知らせ(本人ご利用分)',
      'info@mail.rakuten-card.co.jp',
      '',
      /利用金額: (.{1,16}) 円/,
      '',
      /■利用日: (\d+\/\d+\/\d+)/,
      '楽天カード'
    ),
    new Parser( // 楽天 デビット
      '楽天 デビット',
      '◆デビットカードご利用通知メール',
      'service@ac.rakuten-bank.co.jp',
      '',
      /口座引落分：(.{1,16})円/,
      '',
      '',
      '楽天 デビット'
    ),
    new Parser( // Sony Bank Wallet
      'SonyBankWallet',
      '［Sony Bank WALLET ］Visaデビット',
      'banking@sonybank.net',
      /ご利用加盟店\s*：(.+)/,
      /ご利用金額（※）：(.{1,16})円/,
      '',
      '',
      'SonyBankWallet'
    ),
    new Parser( // JapanNet Visa デビット
      'JNB VISAデビット',
      'Ｖｉｓａデビット',
      'customer2@cc.paypay-bank.co.jp',
      /加盟店名：(.+)/,
      /お引落金額：(.{1,16})円/,
      '',
      '',
      'JNB VISAデビット'
    ),
    new Parser( // JapanNet 引き落とし
      'JNB 引き落とし',
      'お引き落としのご連絡',
      'customer2@cc.paypay-bank.co.jp',
      '',
      /お引落金額：(.{1,16})円/,
      '',
      '',
      'JNB 引き落とし'
    ),
    new Parser( // メルカード
      'メルカード',
      'メルカードのご利用がありました',
      'no-reply@mercari.jp',
      /店舗名\s*：\s*(.*)/,
      /決済金額\s*：\s*￥(.*)/,
      '',
      /決済日時\s*：\s*(.*)/,
      'メルカード'
    ),
    new Parser( // メルペイ
      'メルペイ',
      'コード決済でお支払いがありました',
      'no-reply@mercari.jp',
      /店舗名\s*：\s*(.*)/,
      /取引金額合計\s*：\s*￥(.*)/,
      '',
      /取引日時\s*：\s*(.*)/,
      'メルペイ'
    ),
  ];
  const data = parsers
    .map((parser) => getData(parser, alreadyIds))
    .flat()
    .sort((a, b) => a.date - b.date);
  paste(data);
}

function paste(data) {
  const todayMonth = new Date().getMonth() + 1;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  for (let i = 1; i <= todayMonth; i++) {
    try {
      const thisMonthData = data.filter((row) => row.date.getMonth() + 1 === i);
      if (thisMonthData.length === 0) continue;

      const sheet = spreadsheet.getSheetByName(`${i}月`);
      const lastRow = sheet.getRange(900, 4).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
      sheet
        .getRange(lastRow + 1, 2, thisMonthData.length, 6)
        .setValues(thisMonthData.map((row) => row.toRow()));
    } catch (error) {
      Logger.log(`${i}月 シートに入力できなかった`, error);
    }
  }
}

function getData(parser, alreadyIds) {
  const mails = getMail(parser.subject, parser.from, alreadyIds);
  const rows = mails.map((mail) => {
    const body = mail.getPlainBody();
    const receiveDate = mail.getDate();
    const id = mail.getId();

    const name = (body.match(parser.nameRegex) || [])[1] || parser.nameDefault;
    const dateRaw = (body.match(parser.dateRegex) || [])[1] || receiveDate.toUTCString();
    const date = new Date(dateRaw.replace(/\(.\)/,"")); //(火)などの曜日を削除して日付に変換
    try {
      const price = body.match(parser.priceRegex)[1];
      return new Data(name, price, date, parser.category, '', id);
    } catch (error) {
      Logger.log(error);
      return null;
    }
  });
  return rows.filter((row) => row != null);
}

function getMail(querySubject, from, alreadyIds) {
  const queryParts = [];
  queryParts.push(`subject:${querySubject}`);
  queryParts.push(`from:${from}`);

  const todayStr = Utilities.formatDate(new Date(Date.now() + 86400000), 'JST', 'yyyy/M/d');
  const afterDate = new Date(Date.now() - BEFORE_DAY * 24 * 60 * 60 * 1000);
  const afterDateStr = Utilities.formatDate(afterDate, 'JST', 'yyyy/M/d');
  queryParts.push(`after:${afterDateStr}`);
  queryParts.push(`before:${todayStr}`);

  const SearchQuery = queryParts.join(' ');
  const threads = GmailApp.search(SearchQuery);
  const messages = GmailApp.getMessagesForThreads(threads);
  return messages.flat().filter((message) => !alreadyIds.includes(message.getId()));
}

function getAlreadyId() {
  const todayMonth = new Date().getMonth() + 1;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const alreadyIds = [];
  for (let i = 1; i <= todayMonth; i++) {
    const sheet = spreadsheet.getSheetByName(`${i}月`);
    const lastRow = sheet.getRange(900, 4).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    const numRow = lastRow - 4;
    if (numRow > 0) {
      const ids = sheet.getRange(5, 7, numRow).getValues();
      alreadyIds.push(...ids.flat());
    }
  }
  return alreadyIds;
}
