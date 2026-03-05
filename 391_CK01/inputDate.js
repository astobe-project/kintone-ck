/*
(function () {
  'use strict';

  // 月末集計日のフィールドコード
  const FIELD_YM = '月末集計日'; // DATEフィールド

  
  // Dateオブジェクトから "YYYY-MM-DD" 文字列にする

  function formatDateToYMD(date) {
    const y = date.getFullYear();
    const m = ('0' + (date.getMonth() + 1)).slice(-2);
    const d = ('0' + date.getDate()).slice(-2);
    return `${y}-${m}-${d}`;
  }

    // 基準日（today）から前月末日を求める

  function getPrevMonthEndFromToday() {
    const today = new Date();

    let year = today.getFullYear();
    let month = today.getMonth(); // 0-11（0=1月）

    // 前月へ
    month -= 1;
    if (month < 0) {
      month = 11;   // 12月
      year -= 1;
    }

    // 指定した月の末日： new Date(year, month+1, 0)
    const lastDay = new Date(year, month + 1, 0);
    return formatDateToYMD(lastDay);
  }

  // ★ 新規作成画面表示時 ★
  kintone.events.on('app.record.create.show', function (event) {
    const record = event.record;

    // すでに値が入っていたら上書きしたくない場合はこのifを残す
    if (record[FIELD_YM] && record[FIELD_YM].value) {
      return event;
    }

    const prevMonthEnd = getPrevMonthEndFromToday();
    record[FIELD_YM].value = prevMonthEnd;
    console.log('月末集計日を自動設定(新規作成):', prevMonthEnd);

    return event;
  });

})();
*/