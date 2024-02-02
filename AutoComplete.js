function insertLastUpdated() {
  var ss = SpreadsheetApp.getActiveSheet(); //現在触っているシートを取得
  var currentRow = ss.getActiveCell().getRow(); //アクティブなセルの行番号を取得
  var currentCell = ss.getActiveCell(); //アクティブなセルの入力値を取得
  var cellSomething = currentCell.getValue();
  Logger.log(updateRange);
  //更新日の記入
  if (currentRow > 4) {
    //4行目を除くため
    if (currentCell.getColumn() == 2) {
      if (cellSomething != '') {
        var updateRange = ss.getRange(currentRow, String(currentCell.getColumn() + 2)); //どの列に更新日時を挿入したいか
        var genkinRange = ss.getRange(currentRow, String(currentCell.getColumn() + 3));
        var idRange = ss.getRange(currentRow, String(currentCell.getColumn() + 5));
        if (updateRange.getValue() == '') {
          updateRange.setValue(new Date());
          genkinRange.setValue('現金');
          idRange.setValue('ThisIsMoney');
        }
      }
    }
  }
}
