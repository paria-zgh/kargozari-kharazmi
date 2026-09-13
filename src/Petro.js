import React, { useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import ExcelJS from "exceljs";
import "./App.css";

export const Petro=()=>{
  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);
  const [loading, setLoading] = useState(false);
  const [alertMsg, setAlertMsg] = useState("");

  const handleFile1 = (e) => setFile1(e.target.files[0]);
  const handleFile2 = (e) => setFile2(e.target.files[0]);

  const showAlert = (msg) => {
    setAlertMsg(msg);
    setTimeout(() => setAlertMsg(""), 3000);
  };

  const isNumeric = (v) => {
    if (v === null || v === undefined || v === "") return false;
    const s = String(v).toString().replace(/,/g, "").trim();
    return !isNaN(parseFloat(s)) && isFinite(s);
  };

  const getExcelColumn = (index) => {
    let column = "";
    while (index > 0) {
      let remainder = (index - 1) % 26;
      column = String.fromCharCode(65 + remainder) + column;
      index = Math.floor((index - 1) / 26);
    }
    return column;
  };

  const processFiles = async () => {
    if (!file1 || !file2) {
      showAlert("لطفاً هر دو فایل را انتخاب کنید.");
      return;
    }

    setLoading(true);
    try {
      const data1 = await file1.arrayBuffer();
      const wb1 = XLSX.read(data1, { type: "array" });
      const ws1 = wb1.Sheets[wb1.SheetNames[0]];
      const df1 = XLSX.utils.sheet_to_json(ws1);

      const df1Map = {};
      df1.forEach((row) => {
        const code = String(row["کد عرضه"] || "").trim();
        if (code) df1Map[code] = row;
      });

      const order = [...new Set(df1.map((row) => String(row["کد عرضه"])))]
        .filter((x) => x && x !== "undefined");

      const df1HasOfferQty = df1.some((r) =>
        Object.prototype.hasOwnProperty.call(r, "مقدار عرضه")
      );
      const df1HasDemandQty = df1.some((r) =>
        Object.prototype.hasOwnProperty.call(r, "مقدار تقاضا")
      );
      const df1HasSellerPrice = df1.some((r) =>
        Object.prototype.hasOwnProperty.call(r, "قیمت فروشنده (ریال)")
      );

      const data2 = await file2.arrayBuffer();
      const wb2 = XLSX.read(data2, { type: "array" });
      const ws2 = wb2.Sheets[wb2.SheetNames[0]];
      const df2 = XLSX.utils.sheet_to_json(ws2);

      const df2Map = {};
      df2.forEach((r) => {
        const code = String(r["کد عرضه"] || "").trim();
        if (!df2Map[code]) df2Map[code] = [];
        df2Map[code].push(r);
      });

      const keepColumns = [
        "عرضه",
        "تقاضا",
        "نام کالا",
        "محموله",
        "تولید کننده",
        "نام عرضه کننده",
        "قیمت فروشنده (ریال)",
        "نام مشتری",

      ];

      const sampleRowDf2 = df2[0] || {};
      const availableCols = keepColumns.filter((col) => {
        if (col === "قیمت فروشنده (ریال)" && df1HasSellerPrice) return true;
        if (col === "عرضه" && df1HasOfferQty) return true;
        if (col === "تقاضا" && df1HasDemandQty) return true;
        return Object.prototype.hasOwnProperty.call(sampleRowDf2, col);
      });

      let result = [];
      order.forEach((code) => {
        const subset = (df2Map[code] || []).map((row) => {

          let filtered = {};
          availableCols.forEach((col) => {
            let val = "";

            if (col === "عرضه" && df1HasOfferQty) {
              val = df1Map[code]?.["مقدار عرضه"] ?? row[col] ?? "";
            } else if (col === "تقاضا" && df1HasDemandQty) {
              val = df1Map[code]?.["مقدار تقاضا"] ?? row[col] ?? "";
            } else if (col === "قیمت فروشنده (ریال)" && df1HasSellerPrice) {
              val = df1Map[code]?.["قیمت فروشنده (ریال)"] ?? row[col] ?? "";
            } else {
              val = row[col] ?? "";
            }

            const outputCol = col === "قیمت فروشنده (ریال)" ? "قیمت پایه" : col;

            filtered[outputCol] = isNumeric(val)
              ? Number(String(val).replace(/,/g, ""))
              : val;
          });

          return filtered;
        });

        if (subset.length > 0) {
          result.push(...subset);
          result.push({});
        }
      });

      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("نتیجه مرتب‌سازی");
      sheet.views = [{ rightToLeft: true }];

      const priceColumns = [
        "قیمت1",
        "قیمت2",
        "قیمت3",
        "قیمت4",
        "قیمت5",
        "قیمت6",
        "قیمت7",
      ];

      const extraCols = ["سقف"];

      const renamedCols = availableCols.map(c =>
        c === "قیمت فروشنده (ریال)" ? "قیمت پایه" : c
      );

      const allCols = [...renamedCols, ...extraCols, ...priceColumns];

      sheet.addRow(allCols);

      result.forEach((row) => {
        if (Object.keys(row).length === 0) {
          sheet.addRow(allCols.map(() => ""));
        } else {
          sheet.addRow(allCols.map((c) => row[c] ?? ""));
        }
      });

      const fontName = "B Nazanin";
      const offerColIdx = allCols.indexOf("عرضه") + 1;
      const demandColIdx = allCols.indexOf("تقاضا") + 1;

      sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          row.eachCell((cell) => {
            cell.font = { name: fontName, bold: true, size: 12 };
            cell.alignment = { vertical: "middle", horizontal: "center" };
            cell.border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFCCCCCC" },
            };
          });
          return;
        }

        const isBlankRow = allCols.every((col) => {
          const v = row.getCell(allCols.indexOf(col) + 1).value;
          return v === "" || v === null || v === undefined;
        });

        if (isBlankRow) return;

        const offerVal = row.getCell(offerColIdx).value;
        const demandVal = row.getCell(demandColIdx).value;

        const shouldBeGreen =
          typeof offerVal === "number" &&
          typeof demandVal === "number" &&
          demandVal <= offerVal;

        row.eachCell((cell, colNumber) => {
          const colName = allCols[colNumber - 1];
          const cellValue = cell.value;

          const numFmtCols = [
            "قیمت1","قیمت2","قیمت3","قیمت4","قیمت5","قیمت6","قیمت7","سقف","قیمت پایه","عرضه","تقاضا"
          ];

          if (colName === "قیمت2") {
            const p7 = getExcelColumn(allCols.indexOf("قیمت7") + 1);
            const p1 = getExcelColumn(allCols.indexOf("قیمت1") + 1);
            cell.value = {
              formula: `IFERROR((${p7}${rowNumber}-${p1}${rowNumber})/6+${p1}${rowNumber},0)`,
            };
          } else if (colName.match(/^قیمت[3-6]$/)) {
            const n = parseInt(colName.replace("قیمت", ""));
            const prev = getExcelColumn(allCols.indexOf(`قیمت${n - 1}`) + 1);
            const p7 = getExcelColumn(allCols.indexOf("قیمت7") + 1);
            const p1 = getExcelColumn(allCols.indexOf("قیمت1") + 1);

            cell.value = {
              formula: `IFERROR((${p7}${rowNumber}-${p1}${rowNumber})/6+${prev}${rowNumber},0)`,
            };
          }

          if (numFmtCols.includes(colName)) {
            cell.numFmt = "#,##0";
          }

          if (colName === "سقف") {
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFD3D3D3" },
            };
          } else if (shouldBeGreen) {
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFC6EFCE" },
            };
            cell.font = {
              name: fontName,
              size: 12,
              color: { argb: "FF006100" },
            };
          } else {
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFFFFFFF" },
            };
          }

          cell.font = cell.font || { name: fontName, size: 12 };

          cell.alignment = {
            vertical: "middle",
            horizontal: "center",
            wrapText: true,
          };

          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        });
      });

      const fontFactor = 1;

      sheet.columns.forEach((col) => {
        let maxLen = 0;

        col.eachCell({ includeEmpty: true }, (c) => {
          const v = c.value;
          let len = 0;

          if (v) {
            if (typeof v === "string") len = v.length;
            else if (typeof v === "number") len = String(v).length;
            else if (v.formula) len = v.formula.length;
          }

          if (len > maxLen) maxLen = len;
        });

        col.width = Math.max(10, Math.ceil(maxLen * fontFactor));
      });

      ["سقف","قیمت1","قیمت2","قیمت3","قیمت4","قیمت5","قیمت6","قیمت7"].forEach((colName)=>{
        const idx = allCols.indexOf(colName);
        if(idx !== -1){
          sheet.getColumn(idx + 1).width = 15;
        }
      });

      const mergeCols = ["عرضه", "تقاضا", "نام کالا", "تولید کننده","نام عرضه کننده", "قیمت پایه"];

      mergeCols.forEach((col) => {

        const idx = allCols.indexOf(col) + 1;

        if (!idx) return;

        let start = 2;

        let prevVal = sheet.getRow(start).getCell(idx).value;

        for (let r = start + 1; r <= sheet.rowCount; r++) {

          const curr = sheet.getRow(r).getCell(idx).value;

          if (curr !== prevVal || curr === null || curr === "") {

            if (r - 1 > start) sheet.mergeCells(start, idx, r - 1, idx);

            start = r;

            prevVal = curr;
          }
        }

        if (sheet.rowCount >= start + 1) {
          sheet.mergeCells(start, idx, sheet.rowCount, idx);
        }
      });

      const buffer = await workbook.xlsx.writeBuffer();

      saveAs(new Blob([buffer]), "اکسل_مرتب.xlsx");

      showAlert("فایل آماده شد و دانلود شد.");

      setFile1(null);
      setFile2(null);

    } catch (e) {

      console.error("خطا در پردازش:", e);

      showAlert("خطا در پردازش فایل‌ها!");

    } finally {

      setLoading(false);

    }
  };

  return (
    <div className="container">
      <h2>📊 مرتب‌سازی اکسل دوم بر اساس اکسل اول</h2>

      <div className="file-input">
        <label htmlFor="file1" className="custom-file-btn">انتخاب فایل اول</label>
        <input id="file1" type="file" accept=".xlsx,.xls" onChange={handleFile1}/>
        <p className="file-name">{file1 ? file1.name : "هیچ فایلی انتخاب نشده"}</p>
      </div>

      <div className="file-input">
        <label htmlFor="file2" className="custom-file-btn">انتخاب فایل دوم</label>
        <input id="file2" type="file" accept=".xlsx,.xls" onChange={handleFile2}/>
        <p className="file-name">{file2 ? file2.name : "هیچ فایلی انتخاب نشده"}</p>
      </div>

      <button onClick={processFiles} disabled={loading || !file1 || !file2}>
        {loading ? "در حال پردازش..." : "دانلود اکسل مرتب‌شده"}
      </button>

      {alertMsg && <div className="custom-alert">{alertMsg}</div>}
    </div>
  );
}


