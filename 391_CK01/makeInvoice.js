(function () {
  'use strict';

  // ===== 自アプリ(391)側 設定 =====
  const SPACE_ID      = 'import_invoice_btn'; // ボタンを置くスペースフィールド
  const FIELD_TEXT    = 'テキスト';          // 元テキストを貼るフィールド

  // ★ 1サブテーブルあたりの最大行数（ここを 3000 / 1000 などに変えればOK）
  const MAX_ROWS_PER_TABLE = 1500;

  // ★ サブテーブル構成（フィールドコードと、行内フィールドに付くサフィックス）
  const SUBTABLE_CONFIG = [
    { code: '明細',   suffix: ''   },  // メイン：顧客No, 店名, ...
    { code: '明細_0', suffix: '_0' },  // 明細_0：顧客No_0, 店名_0, ...
    { code: '明細_1', suffix: '_1' },  // 明細_1：顧客No_1, 店名_1, ...
    { code: '明細_2', suffix: '_2' },
    { code: '明細_3', suffix: '_3' },
    { code: '明細_4', suffix: '_4' },
    { code: '明細_5', suffix: '_5' },
    { code: 'エラーリスト', suffix: '_ERR' }
  ];

  const FIELD_CUSTNO  = '顧客No';    // ベースフィールド名（suffix無し）
  const FIELD_STORE   = '店名';
  const FIELD_DATE    = '日付';      // DATE
  const FIELD_DENPYO  = '伝票No';
  const FIELD_ITEM    = '製品名';    // PDFから取った「製品名」（仕入側名称）
  const FIELD_QTY     = '数量';
  const FIELD_UNIT    = '単位';
  const FIELD_BP      = '仕入単価';
  const FIELD_BA      = '仕入額';
  const FIELD_SP      = '販売単価';
  const FIELD_SA      = '販売額';
  const FIELD_SNAME   = '商品名';    // 販売用の名称
  const FIELD_NOTE    = '備考';
  const FIELD_CHECK   = 'CHECK';     // チェック用
  const FIELD_TAXRATE = '消費税率';

  const FIELD_ERROR_COUNT = '販売価格取得エラー数'; // レコード本体の数値フィールド
  const FIELD_STORE_COUNT = '店名数';              // 顧客No件数
  const FIELD_YM          = '月末集計日';          // テキストから取得してセット

  // レコード本体の集計フィールド
  const FIELD_ROW_COUNT = '明細件数';
  const FIELD_BUY_TOTAL = '仕入合計額';
  const FIELD_SUM_8     = '合計額_8';
  const FIELD_SUM_10    = '合計額_10';

  // ===== 外部アプリ 設定 =====
  const CUSTOMER_APP_ID = 383;  // 顧客マスター
  const PRICE_APP_ID    = 393;  // 価格表

  // ▼顧客マスター(appid=383)
  const C_FIELD_STORE_NAME  = '店名';   // 顧客マスターの「店名」
  const C_FIELD_CUSTOMER_NO = '顧客No'; // 顧客マスターの「顧客No」

  // ▼価格表(appid=393)
  const P_FIELD_CUSTOMER_NO   = '顧客No';
  const P_FIELD_PRODUCT_NAME  = '製品名';   // 仕入側名称（PDF上の製品名と一致）
  const P_FIELD_ITEM_NAME     = '商品名';   // 販売名称
  const P_FIELD_SALE_PRICE    = '販売単価';
  const P_FIELD_TAXRATE       = '消費税率';

  // ===== ログ =====
  function dbg() {
    console.log.apply(console, arguments);
  }

  // 数値文字列の正規化（カンマ除去だけ）
  function normalizeNumber(str) {
    if (!str) return '';
    return String(str).replace(/,/g, '');
  }

  function toNumberSafe(str) {
    const n = Number(normalizeNumber(str));
    return isNaN(n) ? 0 : n;
  }

  // 文字列をキー用に正規化（全角/半角スペースを単一半角スペースに、前後トリム）
  function normalizeKey(str) {
    if (!str) return '';
    let s = String(str);
    // 全角スペース → 半角スペース
    s = s.replace(/\u3000/g, ' ');
    // 連続スペース → 1個
    s = s.replace(/\s+/g, ' ');
    return s.trim();
  }

  // yy/mm/dd → YYYY-MM-DD に変換
  function convertYYMMDDToISO(yyMMdd) {
    const m = yyMMdd.match(/^(\d{2})\/(\d{2})\/(\d{2})$/);
    if (!m) return '';
    const yy = parseInt(m[1], 10);
    const mm = m[2];
    const dd = m[3];
    const century = (yy >= 70) ? '19' : '20';
    const yyyy = century + m[1];
    return `${yyyy}-${mm}-${dd}`;
  }

  // ===== (5416) の2行下から "YYYY MM DD" を拾って DATE 形式(YYYY-MM-DD)にする =====
  function extractYmFromText(rawText) {
    const lines = rawText
      .split(/\r?\n/)
      .map(l => l.trim())
      .filter(l => l.length > 0);

    let targetLine = '';

    for (let i = 0; i < lines.length; i++) {
      if (/\(5416\)/.test(lines[i])) {
        if (i + 2 < lines.length) {
          targetLine = lines[i + 2];
        }
        break;
      }
    }

    if (!targetLine) {
      for (const line of lines) {
        if (/\d{4}\s+\d{2}\s+\d{2}/.test(line)) {
          targetLine = line;
          break;
        }
      }
    }

    if (!targetLine) {
      dbg('extractYmFromText: 対象行が見つかりませんでした');
      return '';
    }

    const m = targetLine.match(/(\d{4})\s+(\d{2})\s+(\d{2})/);
    if (!m) {
      dbg('extractYmFromText: 日付パターンマッチ失敗:', targetLine);
      return '';
    }

    const yyyy = m[1];
    const mm   = m[2];
    const dd   = m[3];
    const iso  = `${yyyy}-${mm}-${dd}`;

    dbg('extractYmFromText: 取得した日付:', iso, '元行:', targetLine);
    return iso;
  }

  // サブテーブル1行（顧客No/商品名/販売単価は後で付与）
  function makeDetailRow(store, date, denpyo, itemName, qty, unit, buyPrice, buyAmount, note) {
    const row = { value: {} };
    row.value[FIELD_CUSTNO]  = { value: '' };
    row.value[FIELD_STORE]   = { value: store  || '' };
    row.value[FIELD_DATE]    = { value: date   || '' };
    row.value[FIELD_DENPYO]  = { value: denpyo || '' };
    row.value[FIELD_ITEM]    = { value: itemName || '' };
    row.value[FIELD_QTY]     = { value: normalizeNumber(qty)    || '' };
    row.value[FIELD_UNIT]    = { value: unit   || '' };
    row.value[FIELD_BP]      = { value: normalizeNumber(buyPrice)  || '' };
    row.value[FIELD_BA]      = { value: normalizeNumber(buyAmount) || '' };
    row.value[FIELD_SP]      = { value: '' };
    row.value[FIELD_SA]      = { value: '' };
    row.value[FIELD_SNAME]   = { value: '' };
    row.value[FIELD_NOTE]    = { value: note   || '' };
    row.value[FIELD_CHECK]   = { value: '' }; // 初期値は空
    row.value[FIELD_TAXRATE] = { value: '' };

    return row;
  }

  // 行の最後の数値
  function extractLastNumber(line) {
    const m = line.match(/([\d,]+(?:\.\d+)?)(?!.*[\d,])/);
    if (!m) return '';
    return normalizeNumber(m[1]);
  }

  /**
   * 商品行パース
   */
  function parseProductLine(text) {
    const tokens = text.split(/\s+/).filter(t => t.length > 0);
    if (tokens.length === 0) {
      return { itemName: text, qty: '', unit: '', price: '', amount: '' };
    }

    const numericIdx = [];
    const numRe = /^[\d.,]+$/;
    for (let i = 0; i < tokens.length; i++) {
      if (numRe.test(tokens[i])) {
        numericIdx.push(i);
      }
    }

    let qty = '';
    let price = '';
    let amount = '';
    let unit = '';

    if (numericIdx.length >= 2) {
      const amountIdx = numericIdx[numericIdx.length - 1];
      const priceIdx  = numericIdx[numericIdx.length - 2];
      const qtyIdx    = (numericIdx.length >= 3)
        ? numericIdx[numericIdx.length - 3]
        : numericIdx[0];

      qty    = normalizeNumber(tokens[qtyIdx]);
      price  = normalizeNumber(tokens[priceIdx]);
      amount = normalizeNumber(tokens[amountIdx]);

      let unitIdx = -1;
      if (qtyIdx + 1 < tokens.length && !numRe.test(tokens[qtyIdx + 1])) {
        unitIdx = qtyIdx + 1;
      } else if (qtyIdx - 1 >= 0 && !numRe.test(tokens[qtyIdx - 1])) {
        unitIdx = qtyIdx - 1;
      }
      if (unitIdx >= 0) unit = tokens[unitIdx];

      let cutIdx = qtyIdx;
      if (unitIdx >= 0) cutIdx = Math.min(qtyIdx, unitIdx);
      const itemTokens = tokens.slice(0, cutIdx);
      const itemName = itemTokens.join(' ').trim();

      return {
        itemName: itemName || text,
        qty,
        unit,
        price,
        amount
      };
    }

    // 数値がほぼ見つからない場合
    return {
      itemName: text,
      qty: '',
      unit: '',
      price: '',
      amount: ''
    };
  }

  /**
   * テキスト → 明細行配列
   */
  function parseInvoiceText(rawText) {
    const lines = rawText
      .split(/\r?\n/)
      .map(l => l.trim())
      .filter(l => l.length > 0);

    const rows = [];
    let currentStore  = '';
    let currentDate   = '';
    let currentDenpyo = '';

    for (let i = 0; i < lines.length; i++) {
      let line = lines[i];

      if (line.indexOf('※は軽減税率対象品目') !== -1) {
        continue;
      }

      // 【〜】行
      if (/^【[^】]+】/.test(line)) {
        const headMatch = line.match(/^【([^】]+)】/);
        const headLabel = headMatch ? headMatch[1].trim() : '';

        // 【伝票合計】系
        if (/^伝票(金額|合計)/.test(headLabel)) {
          const amount = extractLastNumber(line);
          rows.push(
            makeDetailRow(
              currentStore,
              currentDate,
              currentDenpyo,
              '【' + headLabel + '】',
              '',
              '',
              '',
              amount,
              ''
            )
          );
          continue;
        }

        // 【納品先合計】系
        if (/^納品先(合計|金額)/.test(headLabel)) {
          const amount = extractLastNumber(line);
          rows.push(
            makeDetailRow(
              currentStore,
              '',
              '',
              '【' + headLabel + '】',
              '',
              '',
              '',
              amount,
              ''
            )
          );
          continue;
        }

        // ★ 店名など（= currentStore を更新するパターン）
        let label = headLabel;

        // 行全体が「【〜】」で終わっている場合、最外側の【～】を店名とみなす
        if (line.endsWith('】')) {
          const firstIdx = line.indexOf('【');
          const lastIdx  = line.lastIndexOf('】');
          if (firstIdx === 0 && lastIdx > firstIdx) {
            label = line.slice(firstIdx + 1, lastIdx).trim();
          }
        }

        currentStore = label;
        continue;
      }

      // 日付 + 伝票No
      let rest = line;
      const dateDenpyoMatch = line.match(/^(\d{2}\/\d{2}\/\d{2})\s+(\d+)\s+(.*)$/);
      if (dateDenpyoMatch) {
        const yymmdd = dateDenpyoMatch[1];
        currentDate   = convertYYMMDDToISO(yymmdd);
        currentDenpyo = dateDenpyoMatch[2];
        rest          = dateDenpyoMatch[3];
      }

      // 商品行（※〜）
      if (rest.indexOf('※') !== -1) {
        const idx = rest.indexOf('※');
        const afterMark = rest.slice(idx + 1).trim();

        if (afterMark.indexOf('※は軽減税率対象品目') !== -1) {
          continue;
        }

        const parsed = parseProductLine(afterMark);

        rows.push(
          makeDetailRow(
            currentStore,
            currentDate,
            currentDenpyo,
            parsed.itemName,
            parsed.qty,
            parsed.unit,
            parsed.price,   // 仕入単価
            parsed.amount,  // 仕入額
            ''
          )
        );
      }
    }

    return rows;
  }

  /**
   * 日付ごと（店名＋日付単位）の最後に「宅急便代」の行を追加する
   */
  function appendDeliveryChargeRows(rows) {
    if (!rows || !rows.length) return rows;

    const result = [];
    let lastKey = null;
    let hadProductInGroup = false;
    let lastStore = '';
    let lastDate  = '';

    const pushDeliveryRow = () => {
      if (!lastKey || !hadProductInGroup) return;
      result.push(
        makeDetailRow(
          lastStore,
          lastDate,
          '',
          '宅急便代',  // 価格マスターの「製品名」と合わせる
          '1',
          '式',        // 単位「式」
          '',
          '',
          ''
        )
      );
    };

    rows.forEach(row => {
      const v = row.value || {};
      const store = (v[FIELD_STORE] && v[FIELD_STORE].value) || '';
      const date  = (v[FIELD_DATE]  && v[FIELD_DATE].value)  || '';
      const item  = (v[FIELD_ITEM]  && v[FIELD_ITEM].value)  || '';

      const key = `${store}|${date}`;

      if (key !== lastKey) {
        // グループが変わったら前グループ末尾に宅急便代
        pushDeliveryRow();
        lastKey = key;
        lastStore = store;
        lastDate  = date;
        hadProductInGroup = false;
      }

      result.push(row);

      // 通常商品行っぽいものだけを「商品あり」とカウント
      if (item && !/^【/.test(item)) {
        hadProductInGroup = true;
      }
    });

    // 最後のグループについても追加
    pushDeliveryRow();

    return result;
  }

  // ========== 顧客マスター & 価格表 連携 ==========

  function escapeForQuery(str) {
    return String(str || '').replace(/"/g, '\\"');
  }

  async function buildCustomerMapFromRows(rows) {
    const storeSet = new Set();
    rows.forEach(r => {
      const store = r.value[FIELD_STORE] && r.value[FIELD_STORE].value;
      if (store) storeSet.add(store);
    });

    const customerMap = {};
    for (const store of storeSet) {
      const query = `${C_FIELD_STORE_NAME} = "${escapeForQuery(store)}"`;
      const resp = await kintone.api(
        kintone.api.url('/k/v1/records.json', true),
        'GET',
        {
          app: CUSTOMER_APP_ID,
          query: query,
          fields: [C_FIELD_CUSTOMER_NO],
          totalCount: false
        }
      );
//      dbg('顧客マスター query:', store, query, resp);

      if (resp.records && resp.records.length > 0) {
        customerMap[store] = resp.records[0][C_FIELD_CUSTOMER_NO].value;
      } else {
        customerMap[store] = '';
      }
    }

//    dbg('customerMap:', customerMap);
    return customerMap;
  }

  async function buildPriceMap(custNoList) {
    const priceMap = {};
    const uniqueCustNos = Array.from(new Set(custNoList.filter(v => v)));

    if (!uniqueCustNos.length) {
      dbg('buildPriceMap: 顧客Noがないためスキップ');
      return priceMap;
    }

    const chunkSize = 40;
    const limit = 500;

    for (let i = 0; i < uniqueCustNos.length; i += chunkSize) {
      const chunk = uniqueCustNos.slice(i, i + chunkSize);
      if (!chunk.length) continue;

      const inList = chunk.map(c => `"${escapeForQuery(c)}"`).join(',');
      const baseQuery = `${P_FIELD_CUSTOMER_NO} in (${inList})`;

      let offset = 0;
      while (true) {
        const query = `${baseQuery} limit ${limit} offset ${offset}`;
        const resp = await kintone.api(
          kintone.api.url('/k/v1/records.json', true),
          'GET',
          {
            app: PRICE_APP_ID,
            query: query,
            fields: [P_FIELD_CUSTOMER_NO, P_FIELD_PRODUCT_NAME, P_FIELD_ITEM_NAME, P_FIELD_SALE_PRICE, P_FIELD_TAXRATE],
            totalCount: false
          }
        );
        const recs = resp.records || [];
//        dbg('価格表 resp (chunk):', { query, hits: recs.length });

        recs.forEach(r => {
          const custNo = r[P_FIELD_CUSTOMER_NO]  && r[P_FIELD_CUSTOMER_NO].value;
          const prod   = r[P_FIELD_PRODUCT_NAME] && r[P_FIELD_PRODUCT_NAME].value;
          if (!custNo || !prod) return;

          const key = `${custNo}|${normalizeKey(prod)}`;
          priceMap[key] = {
            name:    r[P_FIELD_ITEM_NAME]  ? r[P_FIELD_ITEM_NAME].value  : '',
            price:   r[P_FIELD_SALE_PRICE] ? r[P_FIELD_SALE_PRICE].value : '',
            taxRate: r[P_FIELD_TAXRATE]    ? r[P_FIELD_TAXRATE].value    : ''
          };
        });

        if (recs.length < limit) break;
        offset += limit;
      }
    }

//    dbg('priceMap built:', priceMap);
    return priceMap;
  }

  /**
   * 顧客No・販売単価・商品名・税率を付与する処理
   * 「店名＝共通」の価格は全店で共通利用する
   */
  async function enrichRowsWithCustomerAndPrice(rows) {
    // === 顧客マスターから store → 顧客No マッピング ===
    const customerMap = await buildCustomerMapFromRows(rows);
  
    // ★ 個別顧客No ＋ 宅急便用共通顧客No(091) を必ず含める
    const custNoList = Object.values(customerMap)
      .filter(v => v)
      .concat(['091']);
  
    // === 各顧客Noごとの価格表を取得 ===
    const priceMap = await buildPriceMap(custNoList);
  
    // === ▼ 店名=共通 / 顧客No空 の価格も priceMap に追加 ===
    const commonResp = await kintone.api(
      kintone.api.url('/k/v1/records.json', true),
      'GET',
      {
        app: PRICE_APP_ID,
        query: `${P_FIELD_CUSTOMER_NO} = "" or 店名 = "共通" limit 500`,
        fields: [
          P_FIELD_PRODUCT_NAME,
          P_FIELD_ITEM_NAME,
          P_FIELD_SALE_PRICE,
          P_FIELD_TAXRATE
        ],
        totalCount: false
      }
    );
  
    (commonResp.records || []).forEach(r => {
      const prod = r[P_FIELD_PRODUCT_NAME]?.value || '';
      if (!prod) return;
  
      const key = `共通|${normalizeKey(prod)}`;
      priceMap[key] = {
        name:    r[P_FIELD_ITEM_NAME]?.value  || '',
        price:   r[P_FIELD_SALE_PRICE]?.value || '',
        taxRate: r[P_FIELD_TAXRATE]?.value    || ''
      };
    });
  
    // === ▼ 行ごとに 顧客No・販売情報 を付与 ===
    let errorCount = 0;
    const usedCustNoSet = new Set();
  
    rows.forEach(row => {
      const v = row.value || {};
      const store  = v[FIELD_STORE]?.value || '';
      const item   = v[FIELD_ITEM]?.value  || '';
      const qtyStr = v[FIELD_QTY]?.value   || '';
  
      // --- 顧客No設定（店名ベース） ---
      const custNo = customerMap[store] || '';
      v[FIELD_CUSTNO].value = custNo;
      if (custNo) usedCustNoSet.add(custNo);
  
      // --- 【〜】行・商品名なし行はスキップ ---
      if (!item || /^【/.test(item)) {
        v[FIELD_CHECK].value = 'SKIP';
        return;
      }
  
      // --- 宅急便代の補正 ---
      let keyItem = item;
      const isDelivery = item.includes('宅急便代');
  
      if (isDelivery) {
        keyItem = '宅急便代';
        v[FIELD_UNIT].value = '式';
      }
  
      const normItem = normalizeKey(keyItem);
  
      // === 価格検索（優先順位つき）===
      // ① 個別顧客No
      // ② 宅急便用 共通顧客No(091)
      // ③ 店名=共通
      const keyPersonal = `${custNo}|${normItem}`;
      const key091      = `091|${normItem}`;
      const keyCommon   = `共通|${normItem}`;
  
      const info =
        priceMap[keyPersonal] ||
        priceMap[key091] ||
        priceMap[keyCommon];
  
      // --- ヒットしなかった場合 ---
      if (!info) {
        v[FIELD_SNAME].value   = '';
        v[FIELD_TAXRATE].value = '';
        v[FIELD_SP].value      = '';
        v[FIELD_SA].value      = '';
        v[FIELD_CHECK].value   = 'NG';
        errorCount++;
  
        console.warn('❌ 販売価格未ヒット:', {
          store,
          custNo,
          item,
          keyPersonal,
          key091,
          keyCommon
        });
        return;
      }
  
      // --- 情報反映 ---
      const spriceNorm = normalizeNumber(info.price);
      const snameVal   = info.name    || '';
      const taxRateVal = info.taxRate || '';
  
      v[FIELD_SNAME].value   = snameVal;
      v[FIELD_TAXRATE].value = taxRateVal;
      v[FIELD_SP].value      = spriceNorm;
  
      // --- 金額計算 ---
      if (snameVal && spriceNorm) {
        const qtyNum = toNumberSafe(qtyStr);
        const spNum  = toNumberSafe(spriceNorm);
        v[FIELD_SA].value =
          (qtyNum && spNum) ? String(qtyNum * spNum) : '';
        v[FIELD_CHECK].value = 'OK';
      } else {
        v[FIELD_SA].value = '';
        v[FIELD_CHECK].value = 'NG';
        errorCount++;
  
        console.warn('⚠ 商品名または販売単価不足:', {
          store,
          custNo,
          item,
          snameVal,
          spriceNorm
        });
      }
    });
  
    // === 集計 ===
    const storeCount = usedCustNoSet.size;
    dbg('rows after enrich:', rows);
    dbg('errorCount:', errorCount);
    dbg('storeCount:', storeCount);
  
    return { rows, errorCount, storeCount };
  }





  /**
   * サブテーブル1行を「suffix付きフィールド名」にコピーしたクローンを返す
   */
  function cloneRowWithSuffix(row, suffix) {
    const v = row.value || {};
    const newV = {};

    function c(fieldBase) {
      const src = v[fieldBase];
      if (!src) return;
      newV[fieldBase + suffix] = { value: src.value };
    }

    c(FIELD_CUSTNO);
    c(FIELD_STORE);
    c(FIELD_DATE);
    c(FIELD_DENPYO);
    c(FIELD_ITEM);
    c(FIELD_QTY);
    c(FIELD_UNIT);
    c(FIELD_BP);
    c(FIELD_BA);
    c(FIELD_SP);
    c(FIELD_SA);
    c(FIELD_SNAME);
    c(FIELD_NOTE);
    c(FIELD_CHECK);
    c(FIELD_TAXRATE);

    return { value: newV };
  }



  function extractErrorRows(rows) {
    const errorRows = [];
  
    console.group('🔍 extractErrorRows');
    console.log('rows.length:', rows.length);
  
    rows.forEach((r, idx) => {
      const v = r.value || {};
      const checkVal = v[FIELD_CHECK]?.value;
      if (checkVal !== 'NG') return; // NGのみ
  
      console.log(`NG行 ${idx}:`, {
        店名: v[FIELD_STORE]?.value,
        製品名: v[FIELD_ITEM]?.value,
        CHECK: checkVal
      });
  
      const errRow = { value: {} };
  
      function c(fieldBase) {
        if (v[fieldBase]) {
          errRow.value[fieldBase + '_ERR'] = { value: v[fieldBase].value };
        }
      }
  
      c(FIELD_CUSTNO);
      c(FIELD_STORE);
      c(FIELD_DATE);
      c(FIELD_DENPYO);
      c(FIELD_ITEM);
      c(FIELD_QTY);
      c(FIELD_UNIT);
      c(FIELD_BP);
      c(FIELD_BA);
      c(FIELD_SP);
      c(FIELD_SA);
      c(FIELD_SNAME);
      c(FIELD_TAXRATE);
      c(FIELD_NOTE);
      c(FIELD_CHECK);
  
      errorRows.push(errRow);
    });
  
    console.log('抽出された NG 件数:', errorRows.length);
    console.groupEnd();
  
    return errorRows;
  }





  // ===== メイン処理 =====
  async function handleImport(record, progressCallback) {
    const rawText = record[FIELD_TEXT] && record[FIELD_TEXT].value;
    if (!rawText) {
      alert('テキストフィールドにデータがありません。PDFからコピーして貼り付けてください。');
      return;
    }

    if (typeof progressCallback === 'function') {
      progressCallback(5, 'テキスト取得中...');
    }

    // 月末集計日
    const ymDate = extractYmFromText(rawText);
    dbg('handleImport: 月末集計日(テキストから取得):', ymDate);

    if (typeof progressCallback === 'function') {
      progressCallback(15, '日付抽出中...');
    }

    const appId = kintone.app.getId();
    const recordId = record.$id?.value;
    if (!recordId) {
      alert('レコードIDが取得できませんでした。詳細画面で実行してください。');
      throw new Error('recordId undefined');
    }


    try {
      // Step 0: サブテーブル全削除（明細・明細_0〜明細_5）
      const clearBody = {
        app: appId,
        id: recordId,
        record: {}
      };
      SUBTABLE_CONFIG.forEach(cfg => {
        clearBody.record[cfg.code] = { value: [] };
      });

      await kintone.api(
        kintone.api.url('/k/v1/record', true),
        'PUT',
        clearBody
      );
      dbg('subtables cleared');

      if (typeof progressCallback === 'function') {
        progressCallback(30, '既存明細クリア中...');
      }

      // Step 1: テキスト解析
      let parsedRows = parseInvoiceText(rawText);
      dbg('rows parsed (before appendDeliveryChargeRows):', parsedRows);

      // 宅急便代行の追加
      parsedRows = appendDeliveryChargeRows(parsedRows);
      dbg('rows after appendDeliveryChargeRows:', parsedRows);

      // ★ 総容量チェック（明細テーブル総数の上限）★
      const totalCapacity = MAX_ROWS_PER_TABLE * SUBTABLE_CONFIG.length;
      if (parsedRows.length > totalCapacity) {
          alert(
              `明細行が ${parsedRows.length} 行ありますが、最大登録可能件数は ${totalCapacity} 行です。\n\n` +
              `明細テーブルの容量を超えたため、処理を中断しました。`
          );
          throw new Error('総容量超過のため取込み中断');
      }

      if (parsedRows.length === 0) {
        alert('明細として取り込める行が見つかりませんでした。テキストの形式を確認してください。');
        if (typeof progressCallback === 'function') {
          progressCallback(0, '明細データが見つかりませんでした');
        }
        return;
      }

      if (typeof progressCallback === 'function') {
        progressCallback(50, '明細解析中...');
      }

      // Step 2: 顧客No & 価格表
      const { rows, errorCount, storeCount } = await enrichRowsWithCustomerAndPrice(parsedRows);

      if (typeof progressCallback === 'function') {
        progressCallback(80, '価格表適用中...');
      }

      const totalRows = rows.length;
      dbg('totalRows:', totalRows);

      // サブテーブルごとの value 配列
      const subtableValues = SUBTABLE_CONFIG.map(() => []);

      let overflowCount = 0;
      rows.forEach((r, idx) => {
        const tableIndex = Math.floor(idx / MAX_ROWS_PER_TABLE);
        if (tableIndex >= SUBTABLE_CONFIG.length) {
          overflowCount++;
          return;
        }
        const cfg = SUBTABLE_CONFIG[tableIndex];
        if (cfg.suffix) {
          subtableValues[tableIndex].push(cloneRowWithSuffix(r, cfg.suffix));
        } else {
          subtableValues[tableIndex].push(r);
        }
      });

      if (overflowCount > 0) {
        dbg('行数がサブテーブル総容量を超えています。overflowCount=', overflowCount);
        // 必要なら alert してもよいが、とりあえずログのみにしておく
      }

      // 仕入額合計（NG含む・【〜】行は除外）
      let buyTotal = 0;
      rows.forEach(row => {
        const v = row.value || {};
        const item  = (v[FIELD_ITEM] && v[FIELD_ITEM].value) || '';

        if (item && /^【/.test(item)) return;

        const buyAmount = toNumberSafe(v[FIELD_BA] && v[FIELD_BA].value);
        if (buyAmount) {
          buyTotal += buyAmount;
        }
      });

      // 8%・10%の販売額合計
      let total8  = 0;
      let total10 = 0;
      rows.forEach(row => {
        const v = row.value || {};
        const tax = toNumberSafe(v[FIELD_TAXRATE] && v[FIELD_TAXRATE].value);
        const sa  = toNumberSafe(v[FIELD_SA] && v[FIELD_SA].value);
        if (!sa) return;
        if (tax === 8)  total8  += sa;
        if (tax === 10) total10 += sa;
      });


      // ★ NG行をエラーリスト用に抽出
      const errorRows = extractErrorRows(rows);
      dbg('errorRows (for エラーリスト):', errorRows);


      // Step 3: サブテーブル＋集計項目＋月末集計日 再登録
      const body = {
        app: appId,
        id: recordId,
        record: {}
      };

      SUBTABLE_CONFIG.forEach((cfg, idx) => {
        if (cfg.code === 'エラーリスト') {
          body.record[cfg.code] = { value: errorRows }; // ★ 追加
        } else {
          body.record[cfg.code] = { value: subtableValues[idx] };
        }
      });


      body.record[FIELD_ROW_COUNT]   = { value: totalRows };
      body.record[FIELD_BUY_TOTAL]   = { value: buyTotal };
      body.record[FIELD_SUM_8]       = { value: total8 };
      body.record[FIELD_SUM_10]      = { value: total10 };
      body.record[FIELD_ERROR_COUNT] = { value: errorCount };
      body.record[FIELD_STORE_COUNT] = { value: storeCount };


      if (ymDate) {
        body.record[FIELD_YM] = { value: ymDate };
      }

      const resp = await kintone.api(
        kintone.api.url('/k/v1/record', true),
        'PUT',
        body
      );
//      dbg('update resp', resp);

      if (typeof progressCallback === 'function') {
        progressCallback(100, '取込み完了');
      }

      alert(
        'テキストを明細に取り込み、顧客Noと商品名・販売単価・消費税率を反映しました。\n' +
        '（既存の明細はクリア済み／販売価格取得エラー数・店名数・明細件数・各種合計も更新しました）'
      );

      location.reload();

    } catch (e) {
      console.error(e);
      if (typeof progressCallback === 'function') {
        progressCallback(0, 'エラーが発生しました');
      }
      alert('サブテーブル更新でエラーが発生しました。\n' + (e.message || ''));
    }
  }

  // ===== ボタン設置（進捗バー付き） =====
  kintone.events.on('app.record.detail.show', function (event) {

    const space = kintone.app.record.getSpaceElement(SPACE_ID);
    if (!space) return;

    if (space.querySelector('button[data-invoice-import]')) return;

    const btn = document.createElement('button');
    btn.textContent = 'テキスト → 明細取込み';
    btn.className = 'invoice-import-btn';
    btn.setAttribute('data-invoice-import', '1');

    btn.onclick = function () {
      const subtableNamesText = SUBTABLE_CONFIG.map(cfg => cfg.code).join('・');
      if (!confirm(`既存の明細（${subtableNamesText}）はすべて削除され、新しく取り込み直します。よろしいですか？`)) {
        return;
      }

      let progressWrapper = space.querySelector('#invoice-import-progress');
      if (!progressWrapper) {
        progressWrapper = document.createElement('div');
        progressWrapper.id = 'invoice-import-progress';
        progressWrapper.style.marginTop = '8px';
        progressWrapper.style.fontSize = '12px';
        space.appendChild(progressWrapper);
      } else {
        progressWrapper.innerHTML = '';
      }

      const progressLabel = document.createElement('div');
      progressLabel.textContent = '処理準備中...';

      const barOuter = document.createElement('div');
      barOuter.style.position = 'relative';
      barOuter.style.width = '300px';
      barOuter.style.height = '12px';
      barOuter.style.borderRadius = '6px';
      barOuter.style.backgroundColor = '#eee';
      barOuter.style.overflow = 'hidden';
      barOuter.style.marginTop = '4px';

      const barInner = document.createElement('div');
      barInner.style.height = '100%';
      barInner.style.width = '0%';
      barInner.style.backgroundColor = '#4a90e2';
      barInner.style.transition = 'width 0.2s ease';

      barOuter.appendChild(barInner);
      progressWrapper.appendChild(progressLabel);
      progressWrapper.appendChild(barOuter);

      btn.disabled = true;
      const oldText = btn.textContent;
      btn.textContent = '解析中...';

      (async () => {
        try {
          await handleImport(event.record, function (percent, message) {
            if (typeof percent === 'number') {
              if (percent < 0) percent = 0;
              if (percent > 100) percent = 100;
              barInner.style.width = percent + '%';
            }
            if (message) {
              progressLabel.textContent = message;
            }
          });
        } finally {
          btn.disabled = false;
          btn.textContent = oldText;
        }
      })();
    };

    space.appendChild(btn);
  });


})();
