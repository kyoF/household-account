// 今日の日付を取得
const getTodayDate = () => {
  const sheet = SpreadsheetApp.getActiveSheet();
  const todayDate = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  sheet.getRange('c2').setValue(todayDate);
}

// 明日の日付を取得
const getTomorrowDate = () => {
  const sheet = SpreadsheetApp.getActiveSheet();
  const newDate = new Date();
  const tmpTomorrowDate = new Date(newDate);
  tmpTomorrowDate.setDate(tmpTomorrowDate.getDate() + 1)
  const tomorrowDate = Utilities.formatDate(tmpTomorrowDate, 'Asia/Tokyo', 'yyyy/MM/dd');
  sheet.getRange('c2').setValue(tomorrowDate);
}

// 昨日の日付を取得
const getYesterdayDate = () => {
  const sheet = SpreadsheetApp.getActiveSheet();
  const newDate = new Date();
  const tmpYesterdayDate = new Date(newDate);
  tmpYesterdayDate.setDate(tmpYesterdayDate.getDate() - 1)
  const yesterdayDate = Utilities.formatDate(tmpYesterdayDate, 'Asia/Tokyo', 'yyyy/MM/dd');
  sheet.getRange('c2').setValue(yesterdayDate);
}

// 金額と項目と備考を追加
const inputHouseholdAccount = () => {
  const sheet = SpreadsheetApp.getActiveSheet();

  // 入力値を取得
  const date = Utilities.formatDate(sheet.getRange('c2').getValue(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const category = sheet.getRange('c3').getValue();
  const remarks = sheet.getRange('c4').getValue();
  const price = sheet.getRange('c5').getValue();

  // 日付とカテゴリと金額が入っていない場合はエラー
  if (!(date && category && price)) {
    Browser.msgBox('日付、カテゴリ、金額を全て入力してください');
    return;
  }
  
  // 入力する日の日付が最後に入力してある場所を取得
  const textfinder = sheet.getRange('A12:A').createTextFinder(date);
  const todayCellsList = textfinder.findAll();
  const lastRange = todayCellsList[todayCellsList.length-1].getA1Notation();

  // 日付・曜日・カテゴリ・値段・備考のセル位置を取得
  let lastCellOfDate = sheet.getRange(lastRange);
  let lastCellOfDaysOfWeek = sheet.getRange(lastRange).offset(0,1);
  let lastCellOfPrice = sheet.getRange(lastRange).offset(0,2);
  let lastCellOfCategory = sheet.getRange(lastRange).offset(0, 3);
  let lastCellOfRemarks = sheet.getRange(lastRange).offset(0, 4);

  // 曜日を取得
  const day_num = new Date(lastCellOfDate.getValue()).getDay();
  const daysOfWeek = judgeDaysOfWeek(day_num);

  // 入力するためのセル位置を確保する
  if (lastCellOfPrice.getValue() != '') {
    // 下に1行挿入して、入力先のセルとする
    sheet.insertRowAfter(lastCellOfPrice.getRowIndex());
    lastCellOfDate = lastCellOfDate.offset(1, 0);
    lastCellOfDaysOfWeek = lastCellOfDaysOfWeek.offset(1, 0);
    lastCellOfPrice = lastCellOfPrice.offset(1, 0);
    lastCellOfCategory = lastCellOfCategory.offset(1, 0);
    lastCellOfRemarks = lastCellOfRemarks.offset(1, 0);

    // 挿入した1行に日付と曜日を設定
    lastCellOfDate.setValue(date);
    lastCellOfDaysOfWeek.setValue(daysOfWeek);

    // 土日だった場合、背景色を該当色に設定
    if (day_num == 0) {
      lastCellOfDate.setBackground('#f4cccc');
      lastCellOfDaysOfWeek.setBackground('#f4cccc');
      lastCellOfPrice.setBackground('#f4cccc');
      lastCellOfCategory.setBackground('#f4cccc');
      lastCellOfRemarks.setBackground('#f4cccc');
    } else if (day_num == 6) {
      lastCellOfDate.setBackground('#cfe2f3');
      lastCellOfDaysOfWeek.setBackground('#cfe2f3');
      lastCellOfPrice.setBackground('#cfe2f3');
      lastCellOfCategory.setBackground('#cfe2f3');
      lastCellOfRemarks.setBackground('#cfe2f3');
    }

  }

  // 該当セルに金額・カテゴリ・詳細を設定
  lastCellOfPrice.setValue(price);
  lastCellOfCategory.setValue(category);
  lastCellOfRemarks.setValue(remarks);
}

// 番号から曜日を取得
const judgeDaysOfWeek = (day_num) => {
  if (day_num == 0) {
    return '日';
  } else if (day_num == 1) {
    return '月';
  } else if (day_num == 2) {
    return '火';
  } else if (day_num == 3) {
    return '水';
  } else if (day_num == 4) {
    return '木';
  } else if (day_num == 5) {
    return '金';
  } else if (day_num == 6) {
    return '土';
  }
}