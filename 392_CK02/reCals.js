//------------------------------------------------------
// 保存時に「販売額_8」「販売額_10」を再計算
//------------------------------------------------------
kintone.events.on(['app.record.edit.submit', 'app.record.create.submit'], function (event) {
  const record = event.record;
  const subRows = record['明細'].value || [];

  let total8 = 0;
  let total10 = 0;

  // --- 明細ループ ---
  subRows.forEach(row => {
    const value = Number(row.value['販売額']?.value || 0);
    const rate  = Number(row.value['消費税率']?.value || 0); // ← フィールドコードを実際のものに変更！

    // 税率で振り分け
    if (rate === 8) {
      total8 += value;
    } else if (rate === 10) {
      total10 += value;
    }
  });

  // --- 現在値（数値化）---
  const current8 = Number(record['販売額_8']?.value || 0);
  const current10 = Number(record['販売額_10']?.value || 0);

  // --- 差分チェック ---
  const changed8 = Math.abs(current8 - total8) > 0.001;
  const changed10 = Math.abs(current10 - total10) > 0.001;
  const changed = changed8 || changed10;

  if (changed) {
    const ok = confirm(
      '合計額が変更されました。\n再計算結果を反映しますか？'
    );

    if (ok) {
      record['販売額_8'].value = total8;
      record['販売額_10'].value = total10;
      alert(
        `再計算結果を反映しました。\n` +
        `販売額(8%): ${total8.toLocaleString()}円\n販売額(10%): ${total10.toLocaleString()}円`
      );
    } else {
      // キャンセルしても保存は続行
      alert('再計算をキャンセルしました（元の値を保持します）。');
    }
  }

  return event;
});
