(function () {
  "use strict";

  //=====================================================
  // 設定
  //=====================================================
  const TEMPLATE_APP_ID = "398";
  const PRINTED_AT_FIELD_CODE = "請求書作成日時";
  const LAMBDA_URL = "https://ozfy2pux8i.execute-api.ap-southeast-2.amazonaws.com/default/kintone-pdf-generator";

  //=====================================================
  // CSS 注入
  //=====================================================
  (function injectCSS() {
    const style = document.createElement("style");
    style.textContent = `
      .kintone-custom-btn {
        position: relative;
        background: linear-gradient(135deg, #3b82f6, #2563eb);
        color: #fff;
        border: none;
        border-radius: 999px;
        padding: 4px 12px;
        font-size: 12px;
        font-weight: 600;
        cursor: pointer;
        overflow: hidden;
        transition: all 0.25s ease;
      }
      
      .kintone-custom-btn:hover {
        transform: translateY(-1px);
        box-shadow:
          0 0 6px rgba(59,130,246,0.6),
          0 0 14px rgba(59,130,246,0.4);
      }
      
      .kintone-custom-btn::before {
        content: "";
        position: absolute;
        top: 0;
        left: -120%;
        width: 120%;
        height: 100%;
        background: linear-gradient(
          120deg,
          transparent,
          rgba(255,255,255,0.6),
          transparent
        );
        transition: all 0.5s ease;
      }
      
      .kintone-custom-btn:hover::before {
        left: 120%;
      }
      
      .kintone-custom-btn::after {
        content: "★";
        position: absolute;
        top: -6px;
        right: -6px;
        font-size: 10px;
        color: gold;
        opacity: 0;
        transform: scale(0.5);
        transition: all 0.25s ease;
      }
      
      .kintone-custom-btn:hover::after {
        opacity: 1;
        transform: scale(1);
      }
      
      .kintone-custom-btn:active {
        transform: translateY(1px);
      }

      
      .kintone-custom-btn-single {
        background: linear-gradient(135deg, #60a5fa, #3b82f6); /* 一括より少し淡い */
        color: #fff;
        border: 1px solid #3b82f6;
        border-radius: 999px;
        padding: 4px 12px;
        font-size: 12px;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.25s ease;
      }
      
      .kintone-custom-btn-single:hover {
        background: linear-gradient(135deg, #3b82f6, #2563eb);
        box-shadow: 0 2px 6px rgba(59,130,246,0.4);
        transform: translateY(-1px);
      }
      
      .kintone-custom-btn-single:active {
        transform: translateY(0);
        box-shadow: 0 1px 3px rgba(59,130,246,0.3);
      }

    `;
    document.head.appendChild(style);
  })();

  //=====================================================
  // テンプレート取得
  //=====================================================
  async function getTemplateFile(type) {
    const query = `種別 in ("${type}")`;
    const resp = await kintone.api(
      kintone.api.url("/k/v1/records", true),
      "GET",
      { app: TEMPLATE_APP_ID, query }
    );

    if (!resp.records.length) {
      throw new Error("テンプレートが見つかりません");
    }

    const rec = resp.records[0];
    const templateKey = rec["ひな形ファイル"].value[0].fileKey;
    const styleKey = rec["スタイルファイル"]?.value[0]?.fileKey;

    async function fetchFile(key) {
      if (!key) return "";
      const res = await fetch(
        kintone.api.url("/k/v1/file", true) + `?fileKey=${key}`,
        {
          method: "GET",
          headers: { "X-Requested-With": "XMLHttpRequest" }
        }
      );
      return res.text();
    }

    return {
      templateContent: await fetchFile(templateKey),
      styleContent: await fetchFile(styleKey)
    };
  }

  //=====================================================
  // プレースホルダ置換
  //=====================================================
  function replacePlaceholders(template, record) {
    return template.replace(/\${(.*?)}/g, (m, field) => {
      const f = record[field];
      if (!f) return m;

      if (f.type === "NUMBER" || f.type === "CALC") {
        return Number(f.value || 0).toLocaleString();
      }
      return f.value ?? "";
    });
  }

  // -----------------------------
  // 発行日（翌月1日）生成
  // -----------------------------
  function getIssueDateFromRows(rows) {
    if (!rows.length) return "";
  
    const firstDateStr = rows[0].date;
    if (!firstDateStr) return "";
  
    const d = new Date(firstDateStr.replace(/-/g, "/"));
    if (isNaN(d)) return "";
  
    // 翌月1日
    const issue = new Date(d.getFullYear(), d.getMonth() + 1, 1);
  
    const pad = n => String(n).padStart(2, "0");
    return `${issue.getFullYear()}年${pad(issue.getMonth() + 1)}月${pad(issue.getDate())}日`;
  }

  //=====================================================
  // 日時保存用の共通関数
  //=====================================================
  async function updatePrintedAt(appId, recordId) {
    if (!PRINTED_AT_FIELD_CODE) return;
  
    const now = getNowForKintoneDatetime();
  
    await kintone.api(
      kintone.api.url("/k/v1/record", true),
      "PUT",
      {
        app: appId,
        id: recordId,
        record: {
          [PRINTED_AT_FIELD_CODE]: { value: now }
        }
      }
    );
  }

  //=====================================================
  // 完全版HTML生成（ページ分割・明細対応）
  //=====================================================
  function buildCompleteHTML(record, templateHTML, styleCSS) {
  
    const subRows   = record["明細"]?.value || [];
    const extraRows = record["明細追加分"]?.value || [];
  
    // -----------------------------
    // ページ分割
    // -----------------------------
    function paginate(rows) {
      const pages = [];
      pages.push(rows.slice(0, 24));
      for (let i = 24; i < rows.length; i += 48) {
        pages.push(rows.slice(i, i + 48));
      }
      return pages;
    }
  
    // -----------------------------
    // 請求対象期間文字列生成
    // -----------------------------
    function getBillingPeriodText(rows) {
      if (!rows.length) return "";
  
      const firstDateStr = rows[0].date;
      if (!firstDateStr) return "";
  
      const d = new Date(firstDateStr.replace(/-/g, "/"));
      if (isNaN(d)) return "";
  
      const y = d.getFullYear();
      const m = d.getMonth();
  
      const firstDay = new Date(y, m, 1);
      const lastDay  = new Date(y, m + 1, 0);
  
      const fmt = (dt) =>
        `${dt.getFullYear()}年${dt.getMonth() + 1}月${dt.getDate()}日`;
  
      return `請求対象期間：${fmt(firstDay)} ～ ${fmt(lastDay)}`;
    }
  
    // -----------------------------
    // 明細 + 明細追加分 を統合
    // -----------------------------
    function buildMergedRows() {
      const map = {};
  
      // 通常明細
      subRows.forEach(r => {
        const date = r.value["日付"]?.value || "";
        if (!map[date]) map[date] = [];
        map[date].push({
          date,
          item: r.value["商品名"]?.value || "",
          qty: r.value["数量"]?.value || "",
          price: r.value["販売単価"]?.value || 0,
          amount: r.value["販売額"]?.value || 0
        });
      });
  
      // 明細追加分（同日付の下に追加）
      extraRows.forEach(r => {
        const amount = Number(r.value["販売額_追加分"]?.value || 0);
      
        // ★ 販売額_追加分 が 0 の場合は追加しない
        if (amount === 0) return;
      
        const date = r.value["日付_追加分"]?.value || "";
        if (!map[date]) map[date] = [];
      
        map[date].push({
          date,
          item: r.value["商品名_追加分"]?.value || "",
          qty: r.value["数量_追加分"]?.value || "",
          price: r.value["販売単価_追加分"]?.value || 0,
          amount: amount
        });
      });

  
      // 日付昇順でフラット化
      return Object.keys(map)
        .sort()
        .flatMap(date => map[date]);
    }
  
    // ===== 統合明細 =====
    const mergedRows = buildMergedRows();
    const pageList = paginate(mergedRows);
    const totalPages = pageList.length || 1;
  
    const parser = new DOMParser();
    const baseDoc = parser.parseFromString(templateHTML, "text/html");
  
    // -----------------------------
    // 各ブロック取得
    // -----------------------------
    const headerHTML  = baseDoc.querySelector("header")?.outerHTML || "";
    const summaryHTML = baseDoc.querySelector(".summary")?.outerHTML || "";
    const taxHTML     = baseDoc.querySelector(".tax-breakdown")?.outerHTML || "";
    const detailHTML  = baseDoc.querySelector(".detail")?.outerHTML || "";
    const footerHTML  = baseDoc.querySelector("footer")?.outerHTML || "";
  
    let htmlAll = "";
  
    // =============================
    // ページ生成
    // =============================
    pageList.forEach((rows, pageIndex) => {
  
      const billingPeriodText = getBillingPeriodText(rows);
      const issueDateText    = getIssueDateFromRows(rows);
      let pageHTML = "";
  
      // --- 1ページ目のみ ---
      if (pageIndex === 0) {
      
        // =============================
        // 発行日（翌月1日）を計算
        // =============================
        function getIssueDateFromRows(rows) {
          if (!rows.length) return "";
      
          const firstDateStr = rows[0].date;
          if (!firstDateStr) return "";
      
          const d = new Date(firstDateStr.replace(/-/g, "/"));
          if (isNaN(d)) return "";
      
          // 翌月1日
          const issue = new Date(d.getFullYear(), d.getMonth() + 1, 1);
      
          const pad = n => String(n).padStart(2, "0");
          return `${issue.getFullYear()}年${pad(issue.getMonth() + 1)}月${pad(issue.getDate())}日`;
        }
      
        const issueDateText = getIssueDateFromRows(rows);
      
        // =============================
        // ヘッダーをDOMとして操作
        // =============================
        const headerDoc = parser.parseFromString(headerHTML, "text/html");
      
        const issueDateEl = headerDoc.querySelector("#issue-date");
        if (issueDateEl && issueDateText) {
          issueDateEl.textContent = `発行日：${issueDateText}`;
        }
      
        // =============================
        // 消費税内訳をDOMとして操作
        // =============================
        const taxDoc = parser.parseFromString(taxHTML, "text/html");
      
        // 「その他（10%対象）」行を取得
        const otherRow = Array.from(
          taxDoc.querySelectorAll("tr")
        ).find(tr => tr.textContent.includes("その他"));
      
        if (otherRow) {
          const amountCell = otherRow.querySelector(".tax-amount");
          const taxCell    = otherRow.querySelector(".tax-tax");
      
          const amount = Number(
            (amountCell?.textContent || "")
              .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
              .replace(/[,円\s]/g, "")
          );
      
          const tax = Number(
            (taxCell?.textContent || "")
              .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
              .replace(/[,円\s]/g, "")
          );
      
          // 金額・税額ともに 0 のときだけ削除
          if (amount === 0 && tax === 0) {
            otherRow.remove();
          }
        }
      
        // =============================
        // HTML結合
        // =============================
        pageHTML +=
          headerDoc.body.innerHTML +
          summaryHTML +
          taxDoc.body.innerHTML;
      }


  
      // --- 明細 ---
      const detailDoc = parser.parseFromString(detailHTML, "text/html");
  
      // 請求対象期間をタイトル直下に表示
      if (billingPeriodText) {
        const titleEl = detailDoc.querySelector(".detail-title");
        if (titleEl) {
          const periodEl = detailDoc.createElement("div");
          periodEl.className = "billing-period";
          periodEl.textContent = billingPeriodText;
          titleEl.insertAdjacentElement("afterend", periodEl);
        }
      }
  
      const tbody = detailDoc.querySelector(".detail-body");
      if (tbody) {
        tbody.innerHTML = "";
  
        rows.forEach(r => {
        
          const amount = Number(r.amount || 0);
          const price  = Number(r.price  || 0);
          const isNegative = amount < 0;
          // ★ マイナス行なら span で包む（行全体）
          const wrap = (v) =>
            isNegative ? `<span class="neg-row">${v}</span>` : v;
        
          tbody.insertAdjacentHTML("beforeend", `
            <tr>
              <td>${wrap(r.date)}</td>
              <td>${wrap(r.item)}</td>
              <td>${wrap(r.qty)}</td>
              <td>${wrap(price.toLocaleString())}</td>
              <td>${wrap(amount.toLocaleString())}</td>
            </tr>
          `);
        });

      }
  
      pageHTML += detailDoc.body.innerHTML;
  
      // --- フッター ---
      const footerDoc = parser.parseFromString(footerHTML, "text/html");
      const pageNo = footerDoc.querySelector(".page-number");
      
      if (pageNo) {
        const current = pageIndex + 1;
      
        // ★ 表示文言切り替え
        const suffix =
          current === totalPages
            ? "以上"
            : "次ページ";
      
        pageNo.textContent = `${current} / ${totalPages}　${suffix}`;
      }
      
      pageHTML += footerDoc.body.innerHTML;

  
      // --- ページラップ ---
      htmlAll += `
        <div class="a4page">
          <div class="pdf-content">
            ${pageHTML}
          </div>
        </div>
      `;
    });
  
    // =============================
    // 完成HTML
    // =============================
    return `
  <!DOCTYPE html>
  <html lang="ja">
  <head>
    <meta charset="UTF-8">
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP&display=swap" rel="stylesheet">
    <style>
      * { font-family: 'Noto Sans JP', sans-serif !important; }
      ${styleCSS}
    </style>
  </head>
  <body>
    ${htmlAll}
  </body>
  </html>
  `;
  }









  //=====================================================
  // Lambda呼び出し
  //=====================================================
  async function callLambdaPDF(html, fileName) {
    console.log("▶ Lambda呼び出し:", fileName);

    const [body, status] = await kintone.proxy(
      LAMBDA_URL,
      "POST",
      { "Content-Type": "application/json" },
      JSON.stringify({ html, fileName })
    );

    if (status !== 200) {
      console.error("❌ Lambda error:", status, body);
      throw new Error(`Lambda failed with status ${status}`);
    }

    const result = JSON.parse(body);

    if (!result.ok || !result.base64) {
      console.error("❌ Invalid Lambda response:", result);
      throw new Error("Lambda response invalid");
    }

    console.log("✅ Lambda成功:", fileName);
    return result;
  }

  //=====================================================
  // Base64 → Blob変換
  //=====================================================
  function base64ToBlob(base64, mime) {
    const cleanBase64 = base64.replace(/\s/g, '');
    const binary = atob(cleanBase64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return new Blob([bytes], { type: mime });
  }

  //=====================================================
  // 請求サマリ集計
  //=====================================================
  function buildSummaryFromRecords(records) {
    const map = {};
  
    records.forEach(rec => {
      // -----------------------------
      // 基本キー項目
      // -----------------------------
      const ym         = rec["月末集計日"]?.value || "";
      const customerNo = rec["顧客No"]?.value || "";
      const storeName  = rec["店名"]?.value || "";
      const billToName = rec["請求先名"]?.value || "";
  
      // -----------------------------
      // 振込先：スペース区切りの先頭のみ
      // -----------------------------
      const bankRaw = rec["振込先"]?.value || "";
      const bank = bankRaw.split(/\s+/)[0] || "";
  
      // -----------------------------
      // 集計キー（年月＋顧客No）
      // -----------------------------
      const key = `${ym}__${customerNo}`;
  
      // -----------------------------
      // 初期化
      // -----------------------------
      if (!map[key]) {
        map[key] = {
          月末集計日: ym,
          顧客No: customerNo,
          店名: storeName,
          請求先名: billToName,
          振込先: bank,
  
          前回御請求額: 0,
          繰越金額: 0,
          今回御買上額: 0,
          消費税: 0,
          今回御請求額: 0
        };
      }
  
      // -----------------------------
      // 数値集計（すべて Number + fallback）
      // -----------------------------
      map[key].前回御請求額 += Number(rec["前回御請求額"]?.value || 0);
      map[key].繰越金額     += Number(rec["繰越金額"]?.value || 0);
      map[key].今回御買上額 += Number(rec["今回御請求額_税抜"]?.value || 0);
      map[key].消費税       += Number(rec["消費税"]?.value || 0);
      map[key].今回御請求額 += Number(rec["今回御請求額"]?.value || 0);
    });
  
    // -----------------------------
    // 配列で返却
    // -----------------------------
    return Object.values(map);
  }



  //=====================================================
  // 入金管理台帳HTML生成（buildSummaryFromRecords の出力に合わせる版）
  //=====================================================
  function buildSummaryHTML(summaryList, billingMonth) {
  
    // ============================
    // 合計用
    // ============================
    let totalCount   = 0;
    let totalPrev    = 0; // 前回御請求額
    let totalCarry   = 0; // 繰越金額
    let totalSales   = 0; // 今回御買上額
    let totalTax     = 0; // 消費税
    let totalBill    = 0; // 今回御請求額

    // ============================
    // 顧客No 999 を分離
    // ============================
    const normalList = summaryList.filter(s => s.顧客No !== "999");
    const special999 = summaryList.find(s => s.顧客No === "999");

    // billingMonth が無い/壊れてても落ちないように保険
    const yyyymm = String(billingMonth || "")
      .replace(/-/g, "")
      .slice(0, 6) || "YYYYMM";
  
    const fmt = (v) => Number(v || 0).toLocaleString();
  
    // ============================
    // 明細行 + 合計集計
    // ============================
    const rowsHTML = normalList.map(s => {
  
      totalCount += 1;
      totalPrev  += Number(s.前回御請求額 || 0);
      totalCarry += Number(s.繰越金額 || 0);
      totalSales += Number(s.今回御買上額 || 0);
      totalTax   += Number(s.消費税 || 0);
      totalBill  += Number(s.今回御請求額 || 0);

      return `
        <tr>
          <td style="text-align:left;">${s.月末集計日 || ""}</td>
          <td style="text-align:left;">${s.顧客No || ""}</td>
          <td style="text-align:left;">${s.店名 || ""}</td>
          <td style="text-align:left;">${s.請求先名 || ""}</td>
          <td style="text-align:left;">${s.振込先 || ""}</td>
  
          <td>${fmt(s.前回御請求額)}</td>
          <td>${fmt(s.繰越金額)}</td>
          <td>${fmt(s.今回御買上額)}</td>
          <td>${fmt(s.消費税)}</td>
          <td>${fmt(s.今回御請求額)}</td>
  
          <td style="text-align:center;">□</td>
        </tr>
      `;
    }).join("");
  
    // ============================
    // ★ 合計行（③はここ）
    // ============================
    const totalRowHTML = `
      <tr style="font-weight:600; background:#f0f0f0;">
        <td colspan="5" style="text-align:center;">
          合計（${totalCount} 件）
        </td>
        <td style="text-align:right;">${fmt(totalPrev)}</td>
        <td style="text-align:right;">${fmt(totalCarry)}</td>
        <td style="text-align:right;">${fmt(totalSales)}</td>
        <td style="text-align:right;">${fmt(totalTax)}</td>
        <td style="text-align:right;">${fmt(totalBill)}</td>
        <td></td>
      </tr>
    `;

    // ============================
    // ★ 顧客No999 別表
    // ============================
    const specialTableHTML = special999 ? `
      <h2 style="margin:16px 0 6px; font-size:12px;">
        自社分（顧客No 999）
      </h2>
    
      <table>
        <thead>
          <tr>
            <th>月末集計日</th>
            <th>顧客No</th>
            <th>店名</th>
            <th>社名</th>
            <th>今回御買上額</th>
            <th>消費税</th>
            <th>今回御請求額</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>${special999.月末集計日 || ""}</td>
            <td>${special999.顧客No}</td>
            <td>${special999.店名 || ""}</td>
            <td>${special999.請求先名 || ""}</td>
            <td style="text-align:right;">${fmt(special999.今回御買上額)}</td>
            <td style="text-align:right;">${fmt(special999.消費税)}</td>
            <td style="text-align:right;">${fmt(special999.今回御請求額)}</td>
          </tr>
        </tbody>
      </table>
    ` : "";

  
    // ============================
    // HTML出力
    // ============================
    return `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP&display=swap" rel="stylesheet">
        <style>
          * { font-family: 'Noto Sans JP', sans-serif !important; }
          @page { size: A4 portrait; margin: 12mm 10mm; }
          body { margin: 0; }
          h1 { text-align: center; margin: 0 0 12px 0; font-size: 16px; }
          table { width: 100%; border-collapse: collapse; font-size: 9px; }
          th, td { border: 1px solid #333; padding: 3px 4px; white-space: nowrap; }
          thead th { background: #f5f5f5; font-weight: 600; }
  
          /* 左寄せ列（文字） */
          th:nth-child(1), th:nth-child(2), th:nth-child(3), th:nth-child(4), th:nth-child(5),
          td:nth-child(1), td:nth-child(2), td:nth-child(3), td:nth-child(4), td:nth-child(5) {
            text-align: left;
          }
  
          /* 右寄せ列（数値） */
          th:nth-child(n+6):nth-child(-n+10),
          td:nth-child(n+6):nth-child(-n+10) {
            text-align: right;
          }
  
          /* 入金チェック */
          th:nth-child(11), td:nth-child(11) {
            text-align: center;
          }
        </style>
      </head>
      <body>
        <h1>${yyyymm} 入金管理台帳</h1>
        <table>
          <thead>
            <tr>
              <th>月末集計日</th>
              <th>顧客No</th>
              <th>店名</th>
              <th>請求先名</th>
              <th>振込先</th>
              <th>前回御請求額</th>
              <th>繰越金額</th>
              <th>今回御買上額</th>
              <th>消費税</th>
              <th>今回御請求額</th>
              <th>入金</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHTML}
            ${totalRowHTML}
          </tbody>
        </table>
          ${specialTableHTML}
      </body>
      </html>
    `;
  }



async function bindRowPrintButtons(templateHTML, styleCSS) {
  document.querySelectorAll('[data-record-id]').forEach(btn => {
    if (btn.dataset.bound) return;
    btn.dataset.bound = "1";

    btn.addEventListener("click", async (e) => {
      e.stopPropagation();

      const recordId = btn.dataset.recordId;
      btn.disabled = true;
      const original = btn.textContent;
      btn.textContent = "印刷中…";

      try {
        await printRecordById(
          recordId,
          templateHTML,
          styleCSS,
          { skipMissingCheck: false }
        );
      } catch (err) {
        console.error(err);
        alert("請求書印刷に失敗しました");
      } finally {
        btn.disabled = false;
        btn.textContent = original;
      }
    });
  });
}



  //=====================================================
  // ボタン生成
  //=====================================================
  const btn = document.createElement('button');
  btn.id = 'invoice-print-btn';
  btn.textContent = '請求書印刷';
  btn.className = 'kintone-custom-btn';

  // --- ブルー系に上書き ---
  btn.style.backgroundColor = '#0d6efd';   // Bootstrap Blue
  btn.style.fontSize = '15px';
  btn.style.fontWeight = '600';
  btn.style.padding = '8px 18px';
  btn.style.borderRadius = '20px';
  btn.style.marginLeft = '8px';

  //=====================================================
  // ホバー効果（ブルー系）
  //=====================================================
  btn.addEventListener('mouseenter', () => {
    if (btn.disabled) return;
    btn.style.backgroundColor = '#0b5ed7';
  });

  btn.addEventListener('mouseleave', () => {
    if (btn.disabled) return;
    btn.style.backgroundColor = '#0d6efd';
  });

  //=====================================================
  // ボタン押下処理
  //=====================================================
  btn.onclick = async () => {
    btn.disabled = true;
    const originalText = btn.textContent;
    btn.textContent = '印刷中…';

    try {
      const recordId = record.$id.value;

      await printRecordById(
        recordId,
        templateContent,
        styleContent
      );

    } catch (err) {
      console.error('請求書印刷エラー:', err);
      alert(
        '請求書印刷中にエラーが発生しました。\n' +
        '詳細はコンソールを確認してください。'
      );
    } finally {
      btn.disabled = false;
      btn.textContent = originalText;
    }
  };

  //=====================================================
  // 日付生成（kintone DATETIME形式）
  //=====================================================
  function getNowForKintoneDatetime() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    return (
      d.getFullYear() +
      "-" +
      pad(d.getMonth() + 1) +
      "-" +
      pad(d.getDate()) +
      "T" +
      pad(d.getHours()) +
      ":" +
      pad(d.getMinutes()) +
      ":" +
      pad(d.getSeconds()) +
      "+0900"
    );
  }
  
  //=====================================================
  // 指定レコード 1件印刷
  // options.skipMissingCheck: true → 販売額未設定チェックをスキップ（= 一括印刷用YES）
  //=====================================================
  async function printRecordById(recordId, templateHTML, styleCSS, options = {}) {
    const { skipMissingCheck = false } = options;
  
    const appId = kintone.app.getId();
  
    // --- レコード取得 ---
    let result;
    try {
      result = await kintone.api(
        kintone.api.url("/k/v1/record", true),
        "GET",
        { app: appId, id: recordId }
      );
    } catch (e) {
      console.error("レコード取得エラー:", e);
      return { status: "skip" };  // 取得できない → スキップ
    }
  
    const rec = result.record;
    const subRows = rec["明細"]?.value || [];
  
    // --- 販売額未設定チェック（個別印刷時のみ） ---
    if (!skipMissingCheck) {
      const missing = subRows.filter(r => !r.value["販売額"]?.value).length;
      if (missing > 0) {
        const ok = confirm(
          `販売額未設定が ${missing} 件あります。\nそのまま印刷しますか？`
        );
        if (!ok) return { status: "skip" }; // キャンセル → 次に進むが短いWait
      }
    }
  
    // --- プレースホルダ置換 ---
    const filledHTML = replacePlaceholders(templateHTML, rec);

    const completeHTML = buildCompleteHTML(rec, filledHTML, styleCSS);

    // ★ 印刷実行→キャンセル判定付き
    const printResult = await openPrintWindow(rec, completeHTML, styleCSS);
  
    if (printResult.canceled) {
      return { status: "skip", elapsed: printResult.elapsed };
    }
  
    // --- 請求書作成日時の自動保存 ---
    if (PRINTED_AT_FIELD_CODE) {
      try {
        const now = getNowForKintoneDatetime();
        await kintone.api(
          kintone.api.url("/k/v1/record", true),
          "PUT",
          {
            app: appId,
            id: recordId,
            record: {
              [PRINTED_AT_FIELD_CODE]: { value: now }
            }
          }
        );
      } catch (e) {
        console.warn("請求書作成日時保存エラー（処理続行）:", e);
      }
    }
  
    // --- 正常終了 ---
    return { status: "ok", elapsed: printResult.elapsed };
  }

function createProgressLabel(total) {
  let el = document.getElementById("zip-progress");
  if (!el) {
    el = document.createElement("div");
    el.id = "zip-progress";
    el.style.marginLeft = "12px";
    el.style.fontWeight = "600";
    el.style.color = "#0d6efd";
    document.querySelector(".gaia-argoui-app-index-toolbar")
      ?.appendChild(el);
  }
  el.textContent = `0 / ${total} 件`;
  return el;
}


  //=====================================================
  // ZIP生成 & ダウンロード（完全版・進捗表示付き）
  //=====================================================
  async function generateAndDownloadZip(records, templateHTML, styleCSS) {
    const total = records.length;

    const ymDate = records[0]["月末集計日"]?.value; // "2025-12-31"
    const yyyymm = ymDate ? ymDate.replace(/-/g, "").slice(0, 6) : "unknown";
  
    // -----------------------------
    // 進捗表示ラベル作成
    // -----------------------------
    let progressEl = document.getElementById("zip-progress");
    if (!progressEl) {
      progressEl = document.createElement("div");
      progressEl.id = "zip-progress";
      progressEl.style.marginLeft = "12px";
      progressEl.style.fontWeight = "600";
      progressEl.style.color = "#0d6efd";
      document
        .querySelector(".gaia-argoui-app-index-toolbar")
        ?.appendChild(progressEl);
    }
    progressEl.textContent = `0 / ${total} 件`;
  
    // -----------------------------
    // JSZip 動的ロード
    // -----------------------------
    if (typeof JSZip === "undefined") {
      const script = document.createElement("script");
      script.src =
        "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
      document.head.appendChild(script);
  
      await new Promise((resolve, reject) => {
        script.onload = resolve;
        script.onerror = reject;
      });
    }
  
    const zip = new JSZip();
    let done = 0;
  
    // -----------------------------
    // 各レコード処理
    // -----------------------------
    for (const rec of records) {
      const recordId = rec.$id.value;
  
      progressEl.textContent = `${done} / ${total} 件 処理中…`;
  
      // レコード詳細取得
      const result = await kintone.api(
        kintone.api.url("/k/v1/record", true),
        "GET",
        { app: kintone.app.getId(), id: recordId }
      );
      const fullRecord = result.record;
  
      // HTML生成
      const filledHTML = replacePlaceholders(templateHTML, fullRecord);
      const completeHTML = buildCompleteHTML(
        fullRecord,
        filledHTML,
        styleCSS
      );
  
      // ファイル名
      const baseName =
        fullRecord["ファイル名"]?.value ||
        fullRecord["請求書番号"]?.value ||
        `invoice_${recordId}`;
      const pdfFileName = `${baseName}.pdf`;
  
      // LambdaでPDF生成
      const lambdaResult = await callLambdaPDF(
        completeHTML,
        pdfFileName
      );
  
      // Base64 → Blob
      const pdfBlob = base64ToBlob(
        lambdaResult.base64,
        "application/pdf"
      );
  
      // ZIPへ追加
      zip.file(pdfFileName, pdfBlob);
  
      // ★★★ ここを追加 ★★★
      await updatePrintedAt(
        kintone.app.getId(),
        recordId
      );
      
      done++;
      progressEl.textContent = `${done} / ${total} 件 完了`;
    }
  
    // -----------------------------
    // ★ 入金管理台帳PDFを追加
    // -----------------------------
    const summaryList = buildSummaryFromRecords(records);

    if (summaryList.length) {
    
      const summaryHTML = buildSummaryHTML(summaryList, yyyymm);
    
      const ledgerFileName = `入金管理台帳_${yyyymm}.pdf`;
    
      const ledgerResult = await callLambdaPDF(
        summaryHTML,
        ledgerFileName
      );
    
      const ledgerBlob = base64ToBlob(
        ledgerResult.base64,
        "application/pdf"
      );
    
      zip.file(ledgerFileName, ledgerBlob);
    }


    // -----------------------------
    // ZIP圧縮
    // -----------------------------
    progressEl.textContent = "ZIP圧縮中…";
  
    const zipBlob = await zip.generateAsync({ type: "blob" });
  
    // -----------------------------
    // ダウンロード
    // -----------------------------
    const url = URL.createObjectURL(zipBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "invoices.zip";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  
    progressEl.textContent = `完了 (${total} / ${total})`;
  }



  //=====================================================
  // 一覧画面: ZIP一括DLボタン
  //=====================================================
  async function addZipDownloadButton(records, templateHTML, styleCSS) {
    const bar = document.querySelector(".gaia-argoui-app-index-toolbar");
    if (!bar || document.getElementById("invoice-zip-btn")) return;

    const btn = document.createElement("button");
    btn.id = "invoice-zip-btn";
    btn.textContent = "請求書PDF一括DL";
    btn.className = "kintone-custom-btn";
    btn.style.fontSize = "18px";
    btn.style.padding = "10px 10px";
    btn.style.marginLeft = "50px";

    btn.onclick = async () => {
      if (!records || records.length === 0) {
        alert("レコードがありません");
        return;
      }

      const ok = confirm(
        `表示中 ${records.length} 件の請求書PDFと入金管理台帳を生成してZIPでダウンロードします。\n\nよろしいですか?`
      );
      if (!ok) return;

      btn.disabled = true;
      const originalText = btn.textContent;
      btn.textContent = "生成中...";

      try {
        await generateAndDownloadZip(records, templateHTML, styleCSS);
        alert("請求書PDFのダウンロードが完了しました");
      } catch (e) {
        console.error("❌ エラー:", e);
        alert("PDF生成中にエラーが発生しました:\n" + e.message);
      } finally {
        btn.disabled = false;
        btn.textContent = originalText;
      }
    };

    bar.appendChild(btn);
  }

  //=====================================================
  // 小窓で印刷（PDF保存用）
  //=====================================================
  async function openPrintWindow(record, html, styleCSS) {
    const fileName =
      record["ファイル名"]?.value ||
      record["請求書番号"]?.value ||
      `invoice_${record.$id.value}`;
  
    const pw = window.open("", "_blank", "width=900,height=1100");
    if (!pw) {
      alert("ポップアップを許可してください");
      return { canceled: true };
    }
  
    pw.document.open();
pw.document.write(`
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>${fileName}</title>

  <style>
    /* =========================
       通常表示（念のため）
       ========================= */
    body {
      margin: 0;
      padding: 0;
    }

    /* =========================
       印刷専用設定
       ========================= */
    @media print {

      /* ★ プリンタ余白（最重要） */
      @page {
        size: A4;
        margin:0;
      }

      body {
        margin: 0;
        padding: 0;
      }

      /* ★ ページ内の安全領域 */
      .print-root {
        box-sizing: border-box;
        padding: 0;
      }

      /* ★ 1ページ単位のブロック */
      .a4page {
        page-break-after: always;
        box-sizing: border-box;
      }

      /* 最終ページの余計な改ページ防止 */
      .a4page:last-child {
        page-break-after: auto;
      }
    }

    /* =========================
       テンプレCSS注入
       ========================= */
    ${styleCSS}

  </style>
</head>

<body>
  <div class="print-root">
    ${html}
  </div>
</body>
</html>
`);

    pw.document.close();
  
    // 描画待ち
    await new Promise(resolve => setTimeout(resolve, 500));
  
    pw.focus();
    pw.print();
  
    // 自動クローズ
    setTimeout(() => {
      try { pw.close(); } catch (e) {}
    }, 800);
  
    return { canceled: false };
  }


//=====================================================
// 一覧画面イベント（個別PDFボタン + 一括ZIPボタン）
//=====================================================
  kintone.events.on("app.record.index.show", async function (event) {
    console.group("📄 index.show 開始");
  
    const records = event.records || [];
    console.log("records.length =", records.length);
  
    if (!records.length) {
      console.warn("❌ records が空");
      console.groupEnd();
      return event;
    }
  
    // -----------------------------
    // ① テンプレート取得
    // -----------------------------
    let tmpl;
    try {
      console.log("▶ テンプレート取得開始");
      tmpl = await getTemplateFile("請求書");
      console.log("✅ テンプレート取得成功", tmpl);
    } catch (e) {
      console.error("❌ テンプレート取得エラー", e);
      console.groupEnd();
      return event;
    }
  
    // -----------------------------
    // ② 印刷ボタンフィールド取得
    // -----------------------------
    const cells = kintone.app.getFieldElements("印刷ボタン");
    console.log("cells =", cells);
  
    if (!cells) {
      console.error("❌ 印刷ボタン フィールドが一覧に存在しません");
      console.groupEnd();
      return event;
    }
  
    console.log(
      `records.length=${records.length}, cells.length=${cells.length}`
    );
  
    // -----------------------------
    // ③ 各セルにボタン設置
    // -----------------------------
    cells.forEach((cell, i) => {
      console.log(`▶ cell[${i}]`, cell);
  
      if (!records[i]) {
        console.warn(`⚠ records[${i}] が存在しません`);
        return;
      }
  
      const recordId = records[i].$id.value;
      console.log(`recordId=${recordId}`);
  
      cell.innerHTML = "";
      cell.style.display = "flex";
      cell.style.justifyContent = "center";
      cell.style.alignItems = "center";
      cell.style.padding = "0";
      
  
      const btn = document.createElement("button");
      
      btn.className = "kintone-custom-btn-single";
      btn.textContent = "請求書PDF";
      btn.style.fontSize = "12px";
      btn.style.padding = "6px 10px";
      btn.style.margin = "10px 10px";
  
      btn.onclick = async () => {
        console.log(`🖨 PDF作成クリック recordId=${recordId}`);
        try {
          await printRecordById(
            recordId,
            tmpl.templateContent,
            tmpl.styleContent
          );
        } catch (e) {
          console.error("❌ 印刷失敗", e);
          alert("請求書印刷に失敗しました");
        }
      };
  
      cell.appendChild(btn);
      console.log(`✅ ボタン挿入完了 recordId=${recordId}`);
    });
  
    // -----------------------------
    // ④ ZIP一括DLボタン
    // -----------------------------
    console.log("▶ ZIPボタン設置開始");
    await addZipDownloadButton(
      records,
      tmpl.templateContent,
      tmpl.styleContent
    );
    console.log("✅ ZIPボタン設置完了");
  
    console.groupEnd();
    return event;
  });

})();