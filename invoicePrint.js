(function () {
  "use strict";

  //=====================================================
  // 設定
  //=====================================================
  const TEMPLATE_APP_ID       = "398";
  const TEMPLATE_TYPE         = "請求書";
  const PRINT_FIELD_CODE      = "印刷ボタン";
  const PRINTED_AT_FIELD_CODE = "請求書作成日時";

  // ★ 一括印刷の状態フラグ
  let isBulkPrinting = false;       // 一括印刷中かどうか
  let bulkCancelRequested = false;  // 中断リクエストが出ているか
  let bulkBtn = null;             // 一括印刷ボタンの参照


  //=====================================================
  // CSS 注入（オレンジボタン）
  //=====================================================
  //=====================================================
  // CSS 注入（統一デザイン）
  //=====================================================
  (function injectCSS() {
    const style = document.createElement("style");
    style.textContent = `
      /* 共通ボタンデザイン */
      .invoice-print-btn,
      .invoice-bulk-btn {
        background: linear-gradient(135deg, #ff9800, #f57c00);
        color: #fff;
        font-size: 17px;          /* ← サイズ統一 */
        font-weight: 600;
        border: none;
        border-radius: 30px;      /* ← 丸みを強調 */
        padding: 12px 32px;       /* ← 高さ・横幅アップ */
        cursor: pointer;
        box-shadow: 0 3px 8px rgba(0,0,0,0.2);
        transition: all 0.2s ease;
      }
  
      /* ホバー時の共通効果 */
      .invoice-print-btn:hover,
      .invoice-bulk-btn:hover {
        background: linear-gradient(135deg, #ffa726, #ef6c00);
        transform: translateY(-1px);
        box-shadow: 0 6px 12px rgba(0,0,0,0.3);
      }
  
      /* 一覧の一括印刷ボタンは少し間隔を取る */
      .invoice-bulk-btn {
        margin-left: 16px;
      }
  
      @media print {
        .a4page {
          page-break-after: always;
        }
      }
    `;
    document.head.appendChild(style);
  })();



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
    const styleKey    = rec["スタイルファイル"]?.value[0]?.fileKey;

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
      styleContent:    await fetchFile(styleKey)
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

  // （必要ならレイアウトデバッグ用。不要ならコメントアウトしてOK）
  function debugPageHeights(pw) {
    try {
      const A4_HEIGHT_PX = 1122;
//      console.log("===== ページ高さデバッグ開始 =====");
      const pages = pw.document.querySelectorAll(".a4page");
      pages.forEach((page, index) => {
        const h = page.scrollHeight;
//        console.log(`ページ ${index + 1}: 高さ ${h}px`);
        if (h > A4_HEIGHT_PX) {
//          console.warn(
//            `⚠ ページ ${index + 1} が A4 高さを超えています (${h} > ${A4_HEIGHT_PX})`
//          );
        }
      });
//      console.log("===== ページ高さデバッグ終了 =====");
    } catch (e) {
//      console.warn("高さデバッグでエラー:", e);
    }
  }

  //=====================================================
  // 小窓で印刷（自動クローズ / YES前提）
  //=====================================================
  async function openPrintWindow(record, filledHTML, styleCSS) {
    // ★ 印刷開始時間（キャンセル推定に必須）
    const startTime = performance.now();
    // ファイル名
    const fileName =
      record["ファイル名"]?.value ||
      record["請求書番号"]?.value ||
      "invoice_" + record.$id.value;

    const subRows = record["明細"]?.value || [];

    // ページ分割（1ページ目16行、以降48行）
    function paginate(rows) {
      const pages = [];
      pages.push(rows.slice(0, 16));
      for (let i = 16; i < rows.length; i += 48) {
        pages.push(rows.slice(i, i + 48));
      }
      return pages;
    }

    const pageList = paginate(subRows);
    const totalPages = pageList.length;

    const parser = new DOMParser();
    const tmplDoc = parser.parseFromString(filledHTML, "text/html");

    const headerHTML  = tmplDoc.querySelector("header")?.outerHTML || "";
    const summaryHTML = tmplDoc.querySelector(".summary")?.outerHTML || "";
    const taxHTML = tmplDoc.querySelector(".tax-breakdown")?.outerHTML || "";
    const detailHTML0 = tmplDoc.querySelector(".detail")?.outerHTML || "";
    const footerHTML  = tmplDoc.querySelector("footer")?.outerHTML || "";

    let htmlAll = "";

    for (let i = 0; i < totalPages; i++) {
      const rows = pageList[i];
      let pageHTML = "";

      if (i === 0) {
        pageHTML += headerHTML + summaryHTML + taxHTML;
      }

      const detailDoc = parser.parseFromString(detailHTML0, "text/html");
      const tbody = detailDoc.querySelector(".detail-body");
      tbody.innerHTML = "";

      let prevDate = null;

      rows.forEach(r => {
        const d    = r.value["日付"]?.value || "";
        const item = r.value["商品名"]?.value || "";
        const qty  = r.value["数量"]?.value || "";
        const unit = Number(r.value["販売単価"]?.value || 0).toLocaleString();
        const amt  = Number(r.value["販売額"]?.value || 0).toLocaleString();

        tbody.insertAdjacentHTML("beforeend", `
          <tr>
            <td>${prevDate === d ? "" : d}</td>
            <td>${item}</td>
            <td>${qty}</td>
            <td>${unit}</td>
            <td>${amt}</td>
          </tr>
        `);

        prevDate = d;
      });

      const footDoc = parser.parseFromString(footerHTML, "text/html");
      const msg = footDoc.querySelector("#footer-message");
      if (msg) {
        msg.innerHTML = `
          <div style="white-space:nowrap; display:flex; justify-content:flex-end;">
            <span>${i < totalPages - 1 ? "次のページがあります" : "これが最終ページです"}</span>
            <span style="margin-left:4px;">（${i + 1}/${totalPages}）</span>
          </div>
        `;
      }

      pageHTML += detailDoc.body.innerHTML + footDoc.body.innerHTML;

      htmlAll += `
        <div class="a4page">
          <div class="pdf-content">${pageHTML}</div>
        </div>
      `;
    }

    // 小窓生成
    const pw = window.open("", "printWindow", "width=900,height=1100");
    if (!pw) {
      alert("ポップアップを許可してください。");
      return;
    }

    pw.document.open();
    pw.document.write(`
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="UTF-8">
          <title>${fileName}</title>
          <style>${styleCSS}</style>
        </head>
        <body>${htmlAll}</body>
      </html>
    `);
    pw.document.close();

    // 少し待ってから印刷 → 自動クローズ（YES前提）
    await new Promise((resolve) => {
      setTimeout(() => {
        try {
          debugPageHeights(pw);
        } catch (e) {
          console.warn(e);
        }
        try {
          pw.focus();
          pw.print();
        } catch (e) {
          console.warn("print時エラー:", e);
        }
        setTimeout(() => {
          try {
            pw.close();
          } catch (e) {
            console.warn("close時エラー:", e);
          }
          resolve();
        }, 800);
      }, 500);
    });
    //------------------------------------------------------
    // 印刷ダイアログキャンセル推定
    //------------------------------------------------------
    const elapsed = performance.now() - startTime;
    // 小窓が極端に早く閉じたらキャンセル扱い
    const wasCanceled = elapsed < 2500;
  
    // 呼び出し元に返す
    return { canceled: wasCanceled, elapsed };
    
  }

  //=====================================================
  // 指定レコード 1件印刷
  // options.skipMissingCheck: true → 販売額未設定チェックをスキップ（= 一括印刷用YES）
  //=====================================================
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
  
    // ★ 印刷実行→キャンセル判定付き
    const printResult = await openPrintWindow(rec, filledHTML, styleCSS);
  
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

  //=====================================================
  // 一括印刷
  // 1回だけ確認 → 各レコードは自動YES扱い
  // 途中でボタンを再クリックすると中断予約
  //=====================================================
  //=====================================================
  // 一括印刷（中断ボタン対応版）
  //=====================================================
  async function bulkPrint(records, templateHTML, styleCSS) {
  
    if (!records.length) {
      alert("レコードがありません");
      return;
    }
  
    const total = records.length;
  
    const ok = confirm(
      `表示中 ${total} 件を順に印刷します。\n\n` +
      "※各レコードの印刷完了確認は行わず、自動で進みます。\n" +
      "※途中で「中断」もできます。\n\n" +
      "よろしいですか？"
    );
    if (!ok) return;
  
    // ★ 中断フラグ初期化
    bulkCancelRequested = false;
  
    // ★ ボタンを「中断」に変更
    if (bulkBtn) {
      bulkBtn.textContent = "中断";
      bulkBtn.style.background = "#d9534f";
      bulkBtn.style.color = "#fff";
    
      // 中断ボタンクリックで中断要求
      bulkBtn.onclick = () => {
        bulkCancelRequested = true;
        bulkBtn.textContent = "中断（処理中…）";
        bulkBtn.disabled = true;
      };
    
      // ======== 追加：Enterキー無効化（誤作動防止）========
      bulkBtn.onkeydown = (e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          e.stopPropagation();
        }
      };
    
      // ======== 追加：フォーカス除去（Enter連打対策）========
      bulkBtn.onfocus = () => bulkBtn.blur();
    }

  
    for (const rec of records) {
  
      if (bulkCancelRequested) {
        alert("一括印刷を中断しました。");
        break;
      }
  
      const id = rec.$id.value;
  
      const result = await printRecordById(
        id,
        templateHTML,
        styleCSS,
        { skipMissingCheck: true }
      );
  console.log(`Record ID ${id} → elapsed = ${result.elapsed} ms, status = ${result.status}`);
  
      // ★ キャンセル・スキップは wait を長くする
      let waitTime;
      
      if (result.status === "skip") {
        // ★ キャンセル・スキップ時は長めに待つ
        waitTime = 3000;
      } else {
        // ★ 通常は短めで次へ
        waitTime = 200;
      }
    
      await new Promise(r => setTimeout(r, waitTime));
      
    }
  
    // ★ ボタンを元に戻す
    if (bulkBtn) {
      bulkBtn.textContent = "一括印刷";
      bulkBtn.style.background = ""; 
      bulkBtn.style.color = "";
      bulkBtn.disabled = false;
      bulkBtn.onclick = () => bulkPrint(records, templateHTML, styleCSS);
    }
  
    if (!bulkCancelRequested) {
      alert("一括印刷が完了しました。");
    }
  }


  //=====================================================
  // 一覧画面：ボタン設置
  //=====================================================
  kintone.events.on("app.record.index.show", async function (event) {
    const records = event.records || [];
    if (!records.length) return event;

    let templateContent, styleContent;
    try {
      const tmpl = await getTemplateFile(TEMPLATE_TYPE);
      templateContent = tmpl.templateContent;
      styleContent    = tmpl.styleContent;
    } catch (e) {
      console.error("テンプレート取得エラー:", e);
      return event;
    }

    // 個別ボタン
    const cells = kintone.app.getFieldElements(PRINT_FIELD_CODE);
    if (cells && cells.length) {
      cells.forEach((cell, i) => {
        const id = records[i].$id.value;

        cell.innerHTML = "";
        cell.style.display = "flex";
        cell.style.justifyContent = "center";
        cell.style.alignItems = "flex-start";
        cell.style.padding = "8px 0";

        const btn = document.createElement("button");
        btn.className = "invoice-print-btn";
        btn.textContent = "印 刷";
        btn.onclick = () => printRecordById(id, templateContent, styleContent);

        cell.appendChild(btn);
      });
    }

    // 一括ボタン
    const bar = document.querySelector(".gaia-argoui-app-index-toolbar");
    if (bar && !document.getElementById("bulk-print-btn")) {
      bulkBtn = document.createElement("button");
      bulkBtn.id = "bulk-print-btn";
      bulkBtn.className = "invoice-bulk-btn";
      bulkBtn.textContent = "一括印刷";
    
      bulkBtn.onclick = () => bulkPrint(records, templateContent, styleContent);
    
      bar.appendChild(bulkBtn);
    }


    return event;
  });

  //------------------------------------------------------
  // 詳細画面：「請求書印刷」ボタンをツールバーに設置（スタイル内包）
  //------------------------------------------------------
  kintone.events.on('app.record.detail.show', async function (event) {
    const record = event.record;
  
    // ===== すでにボタンがある場合は再生成しない =====
    if (document.getElementById('invoice-print-btn')) return event;
  
    // ===== ツールバー領域 =====
    const toolbar = document.querySelector('.gaia-argoui-app-toolbar');
    if (!toolbar) {
      console.error('ツールバーが見つかりません');
      return event;
    }
  
    // ===== 既存ボタンを削除 =====
    const oldBtn = document.getElementById('generate-invoice-btn');
    if (oldBtn) oldBtn.remove();
  
    // ===== テンプレート読み込み =====
    let templateContent, styleContent;
    try {
      const tmpl = await getTemplateFile('請求書');
      templateContent = tmpl.templateContent;
      styleContent = tmpl.styleContent;
    } catch (e) {
      console.error('テンプレート取得エラー:', e);
      return event;
    }
  
    // ===== ボタン生成 =====
    const btn = document.createElement('button');
    btn.id = 'invoice-print-btn';
    btn.textContent = '請求書印刷';
  
    // ===== 💡ここでデザイン指定 =====
    Object.assign(btn.style, {
      background: 'linear-gradient(135deg, #ff9800, #f57c00)',
      color: '#fff',
      fontSize: '17px',         // ← 文字大きめ
      fontWeight: '600',
      border: 'none',
      borderRadius: '30px',     // ← 丸みを強調
      padding: '12px 32px',     // ← 少し大きめに
      cursor: 'pointer',
      boxShadow: '0 3px 8px rgba(0,0,0,0.2)',
      transition: 'all 0.2s ease',
      marginLeft: '8px'         // ← ツールバー内の余白調整
    });
  
    // ホバー効果（動的）
    btn.addEventListener('mouseenter', () => {
      btn.style.background = 'linear-gradient(135deg, #ffa726, #ef6c00)';
      btn.style.transform = 'translateY(-1px)';
      btn.style.boxShadow = '0 6px 12px rgba(0,0,0,0.3)';
    });
    btn.addEventListener('mouseleave', () => {
      btn.style.background = 'linear-gradient(135deg, #ff9800, #f57c00)';
      btn.style.transform = 'none';
      btn.style.boxShadow = '0 3px 8px rgba(0,0,0,0.2)';
    });
  
    // ===== ボタン動作 =====
    btn.onclick = async () => {
      btn.disabled = true;
      btn.textContent = '印刷中…';
      try {
        const recordId = record.$id.value;
        await printRecordById(recordId, templateContent, styleContent);
      } catch (err) {
        console.error('印刷エラー:', err);
        alert('印刷中にエラーが発生しました（詳細はコンソール）');
      } finally {
        btn.disabled = false;
        btn.textContent = '請求書印刷';
      }
    };
  
    // ===== ツールバーに追加 =====
    toolbar.prepend(btn);
  
    return event;
  });





})();
