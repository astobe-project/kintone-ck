(function () {
  "use strict";
  const DEBUG = true;

  const EVENTS = [
    "app.record.create.submit",
    "app.record.edit.submit"
  ];

  const TARGET_TABLE_CODE = "伝票チェック";
  const TARGET_DENPYO_NO_FIELD_CODE = "伝票No_";
  const TARGET_TOTAL_AMOUNT_FIELD_CODE = "集計金額";
  const TARGET_DENPYO_TOTAL_FIELD_CODE = "伝票合計金額";
  const TARGET_CHECK_FIELD_CODE = "チェック欄";

  const SOURCE_TABLE_CODE = "明細";
  const SOURCE_DENPYO_NO_FIELD_CODE = "伝票No_";
  const SOURCE_AMOUNT_FIELD_CODE = "仕入額";

  function debugLog(...args) {
    if (!DEBUG) return;
    console.log("[aggregateByDenpyoNo]", ...args);
  }

  function toNumber(value) {
    if (value === null || value === undefined || value === "") return 0;
    if (typeof value === "number") return value;
    const normalized = String(value).replace(/,/g, "").trim();
    const n = Number(normalized);
    return Number.isFinite(n) ? n : 0;
  }

  function setCheckValueByType(rowValue, fieldCode, statusText) {
    const cell = rowValue?.[fieldCode];
    if (!cell) return;

    // kintone subtable field types:
    // SINGLE_LINE_TEXT / MULTI_LINE_TEXT / RICH_TEXT / NUMBER / DROP_DOWN / RADIO_BUTTON / CHECK_BOX
    if (cell.type === "CHECK_BOX") {
      // CHECK_BOX expects array
      cell.value = [statusText];
      return;
    }
    cell.value = statusText;
  }

  function buildTotalsByDenpyoNo(
    rows,
    denpyoNoFieldCode,
    sourceAmountFieldCode
  ) {
    const totals = new Map();
    const counts = new Map();

    rows.forEach((row) => {
      const denpyoNo = (row.value[denpyoNoFieldCode]?.value || "").trim();
      if (!denpyoNo) return;

      const amount = toNumber(row.value[sourceAmountFieldCode]?.value);
      totals.set(denpyoNo, (totals.get(denpyoNo) || 0) + amount);
      counts.set(denpyoNo, (counts.get(denpyoNo) || 0) + 1);
    });

    return { totals, counts };
  }

  function detectTarget(record) {
    const targetTableCode = TARGET_TABLE_CODE;
    if (!record[targetTableCode]) return null;

    const targetRows = record[targetTableCode]?.value || [];
    const firstTargetRow = targetRows[0]?.value || {};
    const targetDenpyoNoFieldCode = TARGET_DENPYO_NO_FIELD_CODE;
    const targetTotalAmountFieldCode = TARGET_TOTAL_AMOUNT_FIELD_CODE;
    const targetDenpyoTotalFieldCode = TARGET_DENPYO_TOTAL_FIELD_CODE;
    const targetCheckFieldCode = TARGET_CHECK_FIELD_CODE;
    if (
      !targetDenpyoNoFieldCode ||
      !targetTotalAmountFieldCode ||
      !targetDenpyoTotalFieldCode ||
      !targetCheckFieldCode
    ) return null;
    if (
      !firstTargetRow[targetDenpyoNoFieldCode] ||
      !firstTargetRow[targetTotalAmountFieldCode] ||
      !firstTargetRow[targetDenpyoTotalFieldCode] ||
      !firstTargetRow[targetCheckFieldCode]
    ) return null;

    return {
      targetTableCode,
      targetRows,
      targetDenpyoNoFieldCode,
      targetTotalAmountFieldCode,
      targetDenpyoTotalFieldCode,
      targetCheckFieldCode
    };
  }

  function detectSource(record, target) {
    const rows = record[SOURCE_TABLE_CODE]?.value || [];
    const firstRow = rows[0]?.value || {};
    if (
      rows.length &&
      firstRow[SOURCE_DENPYO_NO_FIELD_CODE] &&
      firstRow[SOURCE_AMOUNT_FIELD_CODE]
    ) {
      return {
        sourceTableCode: SOURCE_TABLE_CODE,
        sourceRows: rows,
        denpyoNoFieldCode: SOURCE_DENPYO_NO_FIELD_CODE,
        sourceAmountFieldCode: SOURCE_AMOUNT_FIELD_CODE
      };
    }

    // fallback
    return {
      sourceTableCode: target.targetTableCode,
      sourceRows: target.targetRows,
      denpyoNoFieldCode: target.targetDenpyoNoFieldCode,
      sourceAmountFieldCode: target.targetTotalAmountFieldCode
    };
  }

  function applyRecalculationToRecord(record) {
    debugLog("table candidates:", Object.keys(record || {}));
    const target = detectTarget(record);
    if (!target || !target.targetRows.length) return false;

    const source = detectSource(record, target);
    if (!source || !source.sourceRows.length) return false;

    debugLog("resolved targetTableCode:", target.targetTableCode);
    debugLog("resolved sourceTableCode:", source.sourceTableCode);
    debugLog("resolved targetDenpyoNoFieldCode:", target.targetDenpyoNoFieldCode);
    debugLog("resolved targetTotalAmountFieldCode:", target.targetTotalAmountFieldCode);
    debugLog("resolved targetDenpyoTotalFieldCode:", target.targetDenpyoTotalFieldCode);
    debugLog("resolved targetCheckFieldCode:", target.targetCheckFieldCode);
    debugLog("resolved sourceDenpyoNoFieldCode:", source.denpyoNoFieldCode);
    debugLog("resolved sourceAmountFieldCode:", source.sourceAmountFieldCode);

    const { totals: totalsByNo, counts: countsByNo } = buildTotalsByDenpyoNo(
      source.sourceRows,
      source.denpyoNoFieldCode,
      source.sourceAmountFieldCode
    );
    const duplicateNos = Array.from(countsByNo.entries()).filter(([, c]) => c > 1);
    debugLog(
      "totalsByNo:",
      Array.from(totalsByNo.entries()).map(([k, v]) => `${k}:${v}`)
    );
    debugLog(
      "duplicate denpyoNo count:",
      duplicateNos.length,
      duplicateNos.map(([k, c]) => `${k}(${c})`)
    );

    const targetType = target.targetRows[0]?.value?.[target.targetTotalAmountFieldCode]?.type;
    debugLog("target field type:", target.targetTotalAmountFieldCode, targetType);
    const checkType = target.targetRows[0]?.value?.[target.targetCheckFieldCode]?.type;
    debugLog("check field type:", target.targetCheckFieldCode, checkType);
    if (targetType === "CALC") {
      console.warn(
        "[aggregateByDenpyoNo] 集計金額が計算(CALC)フィールドのため書き換えできません。フィールドタイプを数値にしてください。"
      );
      return false;
    }

    target.targetRows.forEach((row) => {
      const denpyoNo = (row.value[target.targetDenpyoNoFieldCode]?.value || "").trim();
      const total = denpyoNo
        ? totalsByNo.get(denpyoNo) || 0
        : toNumber(row.value[target.targetTotalAmountFieldCode]?.value);
      const denpyoTotal = toNumber(row.value[target.targetDenpyoTotalFieldCode]?.value);
      const status = Math.abs(denpyoTotal - total) < 0.001 ? "一致" : "不一致";
      debugLog("apply row:", { denpyoNo, total });
      row.value[target.targetTotalAmountFieldCode].value = String(total);
      setCheckValueByType(row.value, target.targetCheckFieldCode, status);
    });

    return true;
  }

  kintone.events.on(EVENTS, (event) => {
    try {
      debugLog("event type:", event.type);
      const updated = applyRecalculationToRecord(event.record);
      if (!updated) {
        console.warn(
          "[aggregateByDenpyoNo] 固定したフィールドコードのいずれかが見つからず、再集計をスキップしました。"
        );
      } else {
        debugLog("recalculation applied");
      }
    } catch (err) {
      console.error("[aggregateByDenpyoNo] 再集計に失敗:", err);
    }
    return event;
  });
})();
