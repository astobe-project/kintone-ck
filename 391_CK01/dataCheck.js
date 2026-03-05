(function () {
  'use strict';

  // ★ 対象フィールドコード
  const TARGET_FIELD = '販売価格取得エラー数';

  // ★ 強調スタイル（オレンジ系）
  const highlightStyle = `
    background-color: #ffb347 !important;   /* 柔らかいオレンジ */
    color: #000 !important;                 /* 黒文字で読みやすい */
    font-weight: bold !important;
    border-radius: 4px !important;
    padding: 3px 6px !important;
  `;

  // ================================
  //  ① 詳細画面：色を付ける
  // ================================
  kintone.events.on('app.record.detail.show', function (event) {
    const record = event.record;
    const value = Number(record[TARGET_FIELD].value || 0);

    const el = kintone.app.record.getFieldElement(TARGET_FIELD);
    if (!el) return event;

    if (value !== 0) {
      el.style.cssText += highlightStyle;
    } else {
      el.style.cssText = ''; // ノーマルに戻す
    }

    return event;
  });

  // ================================
  //  ② 一覧画面：セルに色を付ける
  // ================================
  kintone.events.on('app.record.index.show', function (event) {
    const records = event.records;
    if (!records) return event;

    // フィールド要素を取得
    const elements = kintone.app.getFieldElements(TARGET_FIELD);
    if (!elements) return event;

    // レコードごとに強調
    for (let i = 0; i < records.length; i++) {
      const rec = records[i];
      const val = Number(rec[TARGET_FIELD].value || 0);
      const el = elements[i];

      if (!el) continue;

      const td = el.closest('td');
      if (!td) continue;

      if (val !== 0) {
        td.style.cssText += highlightStyle;
      } else {
        td.style.cssText = ''; // 標準に戻す
      }
    }

    return event;
  });

})();
