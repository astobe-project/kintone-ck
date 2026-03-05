(function () {
  'use strict';

  // ===== 設定 =====
  const SPACE_ID      = 'distribute_btn';   // ボタンを置くスペース
  const TARGET_APP_ID = 392;                // 振り分け先アプリID

  // ▼元アプリ（391側）サブテーブル構成
  const SRC_SUBTABLE_CONFIGS = [
    { code: '明細',   suffix: ''   },
    { code: '明細_0', suffix: '_0' },
    { code: '明細_1', suffix: '_1' },
    { code: '明細_2', suffix: '_2' },
    { code: '明細_3', suffix: '_3' },
    { code: '明細_4', suffix: '_4' },
    { code: '明細_5', suffix: '_5' }
  ];

  // ▼元アプリフィールド
  const SRC_FIELD_YM          = '月末集計日';
  const SRC_FIELD_STORE       = '店名';
  const SRC_FIELD_DATE        = '日付';
  const SRC_FIELD_DENPYO      = '伝票No';
  const SRC_FIELD_PRODUCT     = '製品名';
  const SRC_FIELD_SNAME       = '商品名';
  const SRC_FIELD_QTY         = '数量';
  const SRC_FIELD_UNIT        = '単位';
  const SRC_FIELD_COST_PRICE  = '仕入単価';
  const SRC_FIELD_COST_AMOUNT = '仕入額';
  const SRC_FIELD_SALES_PRICE = '販売単価';
  const SRC_FIELD_SALES_AMOUNT= '販売額';
  const SRC_FIELD_TAXRATE     = '消費税率';
  const SRC_FIELD_NOTE        = '備考';
  const SRC_FIELD_REFLECT     = '反映';
  const SRC_FIELD_CUSTNO      = '顧客No'; // ← 顧客Noを明細から拾う想定

  // ▼ターゲットアプリ（392）ヘッダ
  const T_FIELD_YM       = '月末集計日';
  const T_FIELD_STORE    = '店名';
  const T_FIELD_BILLTO   = '請求先名';
  const T_FIELD_DUE      = '支払期日';
  const T_FIELD_POST     = '郵便番号';
  const T_FIELD_ADDR     = '住所';
  const T_FIELD_CUSTNO   = '顧客No';
  const T_FIELD_SEND     = '送付方法';
  const T_FIELD_BANK     = '振込先';
  const T_FIELD_ITEM_COUNT     = '商品点数';
  const T_FIELD_DELIVERY_COUNT = '宅急便数';
  const T_FIELD_SALES_8        = '販売額_8';
  const T_FIELD_SALES_10       = '販売額_10';

  // ▼ターゲットアプリ明細サブテーブル
  const T_SUBTABLE          = '明細';
  const T_FIELD_DATE        = '日付';
  const T_FIELD_DENPYO      = '伝票No';
  const T_FIELD_PRODUCT     = '製品名';
  const T_FIELD_ITEM_NAME   = '商品名';
  const T_FIELD_QTY         = '数量';
  const T_FIELD_UNIT        = '単位';
  const T_FIELD_COST_PRICE  = '仕入単価';
  const T_FIELD_COST_AMOUNT = '仕入額';
  const T_FIELD_SALES_PRICE = '販売単価';
  const T_FIELD_SALES_AMOUNT= '販売額';
  const T_FIELD_TAXRATE     = '消費税率';
  const T_FIELD_NOTE        = '備考';

  // ▼伝票チェックサブテーブル
  const T_CHECK_SUBTABLE = '伝票チェック';
  const T_CHECK_DATE     = '伝票日付';
  const T_CHECK_DENPYO   = '伝票No_';
  const T_CHECK_TOTAL    = '伝票合計金額';
  const T_CHECK_SUM      = '集計金額';
  const T_CHECK_FLAG     = 'チェック欄';

  // ▼顧客マスター(appId=383)
  const CUSTOMER_APP_ID = 383;
  const C_FIELD_NO      = '顧客No';
  const C_FIELD_DUE     = '支払期日';
  const C_FIELD_POST    = '郵便番号';
  const C_FIELD_ADDR    = '請求先住所';
  const C_FIELD_SEND    = '送付方法';
  const C_FIELD_BILLTO  = '請求先名';
  const C_FIELD_BANK    = '振込先';


  // ===== ユーティリティ =====
  function normalizeNumber(str) {
    if (str === null || str === undefined) return '';
    return String(str).replace(/,/g, '');
  }
  function toNumberSafe(str) {
    const n = Number(normalizeNumber(str));
    return isNaN(n) ? 0 : n;
  }
  function escapeForQuery(str) {
    return String(str || '').replace(/"/g, '\\"');
  }
  function cleanStr(s) {
    return String(s || '').replace(/\s+/g, ' ').trim();
  }

  // ===== データ構築 =====
  function buildExportData(record) {
    const ym = record[SRC_FIELD_YM]?.value;
    const slipTotals = {}, productSums = {}, storeDetails = {};
    let hasDetail = false;

    SRC_SUBTABLE_CONFIGS.forEach(cfg => {
      const sub = record[cfg.code]?.value;
      if (!sub?.length) return;

      sub.forEach(row => {
        const v = row.value || {};
        const store = cleanStr(v[SRC_FIELD_STORE + cfg.suffix]?.value);
        const product = cleanStr(v[SRC_FIELD_PRODUCT + cfg.suffix]?.value);
        let custNo = cleanStr(v[SRC_FIELD_CUSTNO + cfg.suffix]?.value);
        if (!store || !product) return;
        hasDetail = true;

        const date        = v[SRC_FIELD_DATE  + cfg.suffix]?.value;
        const denpyo      = v[SRC_FIELD_DENPYO+ cfg.suffix]?.value;
        const itemName    = cleanStr(v[SRC_FIELD_SNAME   + cfg.suffix]?.value);
        const qty         = v[SRC_FIELD_QTY         + cfg.suffix]?.value;
        const unit        = v[SRC_FIELD_UNIT        + cfg.suffix]?.value;
        const costPrice   = v[SRC_FIELD_COST_PRICE  + cfg.suffix]?.value;
        const costAmount  = v[SRC_FIELD_COST_AMOUNT + cfg.suffix]?.value;
        const salesPrice  = v[SRC_FIELD_SALES_PRICE + cfg.suffix]?.value;
        const salesAmount = v[SRC_FIELD_SALES_AMOUNT+ cfg.suffix]?.value;
        const taxRate     = v[SRC_FIELD_TAXRATE     + cfg.suffix]?.value;
        const note        = v[SRC_FIELD_NOTE        + cfg.suffix]?.value;

        if (product === '【伝票合計】' && date && denpyo) {
          const key = [store, date, denpyo].join('|');
          slipTotals[key] = toNumberSafe(costAmount);
          return;
        }
        if (/^【/.test(product)) return;


        // ===== 宅急便代の特別処理 =====
        const isDelivery =
          product.includes('宅急便代') ||
          itemName.includes('宅急便代');
  
        const salesAmtNum = toNumberSafe(salesAmount);
  
        // 宅急便代が0円なら完全スキップ
        if (salesAmtNum === 0) {
          return;
        }
  
        // 宅急便代で顧客No未設定なら共通顧客Noを使用
        if (isDelivery && !custNo) {
          custNo = '091';
        }

        const detail = { store, custNo, date, denpyo, product, itemName, qty, unit, costPrice, costAmount, salesPrice, salesAmount, taxRate, note };
        if (!storeDetails[store]) storeDetails[store] = [];
        storeDetails[store].push(detail);

        if (date && denpyo) {
          const key = [store, date, denpyo].join('|');
          productSums[key] = (productSums[key] || 0) + toNumberSafe(costAmount);
        }
      });
    });

    if (!hasDetail) throw new Error('明細サブテーブルにデータがありません。');

    const storeChecks = {};
    Object.keys(productSums).forEach(key => {
      const [store, date, denpyo] = key.split('|');
      const sumAmount = productSums[key];
      const slipTot = slipTotals[key];
      const isMatch = (typeof slipTot === 'number') && slipTot === sumAmount;
      const row = {
        value: {
          [T_CHECK_DATE]:   { value: date || '' },
          [T_CHECK_DENPYO]: { value: denpyo || '' },
          [T_CHECK_TOTAL]:  { value: slipTot != null ? String(slipTot) : '' },
          [T_CHECK_SUM]:    { value: String(sumAmount) },
          [T_CHECK_FLAG]:   { value: isMatch ? '' : 'NG' }
        }
      };
      (storeChecks[store] ||= []).push(row);
    });

    return { ym, storeDetails, storeChecks };
  }

  // ===== 顧客マスター検索（顧客Noベース・文字列対応）=====
  async function fetchCustomerMap(customerNos) {
    const result = {};
    if (!customerNos.length) return result;

    const uniqueNos = [...new Set(customerNos.filter(v => v && v.trim() !== ''))];
    if (uniqueNos.length === 0) return result;

    const chunks = [];
    while (uniqueNos.length) chunks.push(uniqueNos.splice(0, 100));

    for (const chunk of chunks) {
      const query = `${C_FIELD_NO} in ("${chunk.map(no => escapeForQuery(no)).join('","')}")`;
      const resp = await kintone.api(kintone.api.url('/k/v1/records.json', true), 'GET', {
        app: CUSTOMER_APP_ID,
        query,
        fields: [C_FIELD_NO, C_FIELD_DUE, C_FIELD_POST, C_FIELD_ADDR, C_FIELD_SEND, C_FIELD_BILLTO, C_FIELD_BANK],
        totalCount: false
      });

      (resp.records || []).forEach(r => {
        const no = r[C_FIELD_NO]?.value || '';
        if (!no) return;
        result[no] = {
          customerNo: no,
          dueDate:    r[C_FIELD_DUE]?.value || '',
          post:       r[C_FIELD_POST]?.value || '',
          address:    r[C_FIELD_ADDR]?.value || '',
          sendMethod: r[C_FIELD_SEND]?.value || '',
          billTo:     r[C_FIELD_BILLTO]?.value || '',
          bank:       r[C_FIELD_BANK]?.value || ''
        };
      });
    }

    return result;
  }

  // ===== 既存削除 =====
  async function deleteExistingTargetRecords(ym, store) {
    if (!ym || !store) return;
    const query = `${T_FIELD_YM} = "${escapeForQuery(ym)}" and ${T_FIELD_STORE} = "${escapeForQuery(store)}"`;
    const res = await kintone.api(kintone.api.url('/k/v1/records.json', true), 'GET', {
      app: TARGET_APP_ID, query, fields: ['$id'], totalCount: false
    });
    const ids = (res.records || []).map(r => r.$id.value);
    if (ids.length)
      await kintone.api(kintone.api.url('/k/v1/records.json', true), 'DELETE', { app: TARGET_APP_ID, ids });
  }

  // ===== エクスポート処理 =====
  async function exportToTargetApp(record, progressCallback) {
    const { ym, storeDetails, storeChecks } = buildExportData(record);
    const stores = Object.keys(storeDetails);
    const T_SUBTABLE_EXTRA = '明細追加分';

    // 顧客Noリスト作成
    const customerNos = stores.map(store => {
      const first = storeDetails[store]?.[0];
      return first ? first.custNo || '' : '';
    }).filter(v => v);

    const customerMap = await fetchCustomerMap(customerNos);

    for (let i = 0; i < stores.length; i++) {
      const store = stores[i];
      const plainDetails = storeDetails[store] || [];
      const checkRows = storeChecks[store] || [];

      const custNo = plainDetails[0]?.custNo || '';
      const c = customerMap[custNo];
      if (!c) {
        console.warn(`⚠ 顧客マスター未ヒット: 顧客No=${custNo}, 店名=${store}`);
        continue;
      }

      const { customerNo, dueDate, post, address, sendMethod, billTo, bank } = c;

      // 集計
      let itemCount8 = 0, deliveryCount = 0, salesTotal8 = 0, salesTotal10 = 0;
      for (const d of plainDetails) {
        const tax = toNumberSafe(d.taxRate);
        const amt = toNumberSafe(d.salesAmount);
        if (tax === 8) { itemCount8++; salesTotal8 += amt; }
        if (tax === 10) salesTotal10 += amt;
        if (d.product.includes('宅急') || d.itemName.includes('宅急')) deliveryCount++;
      }

      await deleteExistingTargetRecords(ym, store);

      const detailRows = plainDetails.map(d => ({
        value: {
          [T_FIELD_DATE]:        { value: d.date || '' },
          [T_FIELD_DENPYO]:      { value: d.denpyo || '' },
          [T_FIELD_PRODUCT]:     { value: d.product || '' },
          [T_FIELD_ITEM_NAME]:   { value: d.itemName || '' },
          [T_FIELD_QTY]:         { value: normalizeNumber(d.qty) },
          [T_FIELD_UNIT]:        { value: d.unit || '' },
          [T_FIELD_COST_PRICE]:  { value: normalizeNumber(d.costPrice) },
          [T_FIELD_COST_AMOUNT]: { value: normalizeNumber(d.costAmount) },
          [T_FIELD_SALES_PRICE]: { value: normalizeNumber(d.salesPrice) },
          [T_FIELD_SALES_AMOUNT]:{ value: normalizeNumber(d.salesAmount) },
          [T_FIELD_TAXRATE]:     { value: normalizeNumber(d.taxRate) },
          [T_FIELD_NOTE]:        { value: d.note || '' }
        }
      }));

      const body = {
        app: TARGET_APP_ID,
        record: {
          [T_FIELD_YM]: { value: ym },
          [T_FIELD_STORE]: { value: store },
          [T_FIELD_BILLTO]: { value: billTo || '' },
          [T_FIELD_CUSTNO]: { value: customerNo || '' },
          [T_FIELD_DUE]: { value: dueDate || '' },
          [T_FIELD_POST]: { value: post || '' },
          [T_FIELD_ADDR]: { value: address || '' },
          [T_FIELD_SEND]: { value: sendMethod || '' },
          [T_FIELD_BANK]: { value: bank || '' },
          [T_FIELD_ITEM_COUNT]: { value: String(itemCount8) },
          [T_FIELD_DELIVERY_COUNT]: { value: String(deliveryCount) },
          [T_FIELD_SALES_8]: { value: String(salesTotal8) },
          [T_FIELD_SALES_10]: { value: String(salesTotal10) },
          [T_SUBTABLE]: { value: detailRows },
          [T_CHECK_SUBTABLE]: { value: checkRows },
          [T_SUBTABLE_EXTRA]: {
            value: [
              {
                value: {
                  '備考_追加分': { value: '' }
                }
              }
            ]
          }
        }
      };

      await kintone.api(kintone.api.url('/k/v1/record.json', true), 'POST', body);

      if (typeof progressCallback === 'function')
        progressCallback(i + 1, stores.length, store);
    }
    return stores.length;
  }

  // ===== 反映フラグ更新 =====
  async function setReflectDone(recordId) {
    await kintone.api(kintone.api.url('/k/v1/record', true), 'PUT', {
      app: kintone.app.getId(),
      id: recordId,
      record: { [SRC_FIELD_REFLECT]: { value: '済' } }
    });
  }

  // ===== ボタン設置 =====
  kintone.events.on('app.record.detail.show', event => {
    const space = kintone.app.record.getSpaceElement(SPACE_ID);
    if (!space || space.querySelector('button[data-invoice-export]')) return;

    const btn = document.createElement('button');
    btn.textContent = '明細→請求集計アプリ反映';
    btn.className = 'invoice-export-btn';
    btn.dataset.invoiceExport = '1';
    btn.style.marginLeft = '8px';
    space.appendChild(btn);

    btn.onclick = async () => {
      const record = event.record;
      const reflectVal = record[SRC_FIELD_REFLECT]?.value || '';
      if (reflectVal === '済' && !confirm('既に反映済みです。再実行で削除→再作成します。続行しますか？')) return;

      const errCount = Number(record['販売価格取得エラー数']?.value || 0);
      if (errCount > 0 && !confirm(`⚠ 販売価格取得エラーが ${errCount} 件あります。続行しますか？`)) return;

      let progressWrapper = space.querySelector('#invoice-export-progress');
      if (!progressWrapper) {
        progressWrapper = document.createElement('div');
        progressWrapper.id = 'invoice-export-progress';
        progressWrapper.style.marginTop = '8px';
        progressWrapper.style.fontSize = '12px';
        space.appendChild(progressWrapper);
      } else progressWrapper.innerHTML = '';

      const progressLabel = document.createElement('div');
      const barOuter = document.createElement('div');
      const barInner = document.createElement('div');
      barOuter.style = 'position:relative;width:300px;height:12px;border-radius:6px;background:#eee;overflow:hidden;margin-top:4px;';
      barInner.style = 'height:100%;width:0%;background:#4CAF50;transition:width 0.2s ease;';
      barOuter.appendChild(barInner);
      progressWrapper.append(progressLabel, barOuter);

      btn.disabled = true;
      const oldText = btn.textContent;
      btn.textContent = '反映中...';

      try {
        const count = await exportToTargetApp(record, (done, total, store) => {
          const percent = Math.round((done / total) * 100);
          barInner.style.width = percent + '%';
          progressLabel.textContent = `処理中: ${done}/${total} 店舗 (${store})`;
        });
        await setReflectDone(record.$id.value);
        barInner.style.width = '100%';
        progressLabel.textContent = `完了: ${count} 店舗のレコードを反映しました。`;
        alert(`請求集計アプリ(appId=${TARGET_APP_ID})に ${count} 店舗分を反映しました。`);
      } catch (e) {
        console.error(e);
        progressLabel.textContent = 'エラー発生。コンソールを確認してください。';
        alert('振り分け処理でエラーが発生しました。\n' + (e.message || e));
      } finally {
        btn.disabled = false;
        btn.textContent = oldText;
      }
    };
  });

})();
