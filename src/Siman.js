import React, { useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import ExcelJS from "exceljs";
import "./App.css";

function Siman() {
  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);
  const [file3, setFile3] = useState(null);
  const [loading, setLoading] = useState(false);
  const [alertMsg, setAlertMsg] = useState("");

  const handleFile1 = (e) => setFile1(e.target.files[0]);
  const handleFile2 = (e) => setFile2(e.target.files[0]);
  const handleFile3 = (e) => setFile3(e.target.files[0]);

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

  //----------------------------------------------------
  // توابع نرمال‌سازی برای عدم حساسیت به فاصله، فونت و علائم
  //----------------------------------------------------
  const normalizeKey = (str) => {
    if (!str) return "";
    return String(str)
      .replace(/[\s\u00a0\u200c\u200d\u200e\u200f]+/g, "") // حذف تمام فاصله‌ها، نیم‌فاصله، فاصله نشکن
      .replace(/ي/g, "ی") // یکسان‌سازی ی عربی
      .replace(/ك/g, "ک") // یکسان‌سازی ک عربی
      .replace(/[()（）[\]]/g, "") // حذف پرانتزها
      .trim();
  };

  const isSellerPriceCol = (key) => {
    const norm = normalizeKey(key);
    return norm.includes("قیمتفروشنده") || norm.includes("قیمتپایه");
  };


  const findColumnKey = (rows, predicate) => {
    for (const r of rows) {
      const found = Object.keys(r).find(predicate);
      if (found) return found;
    }
    return null;
  };

  const processFiles = async () => {
    if (!file1 || !file2 || !file3) {
      showAlert("لطفاً هر سه فایل را انتخاب کنید.");
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
        if (code) {
          df1Map[code] = row;
          // برای اطمینان از تطابق عددی
          const numericCode = code.replace(/\D/g, "");
          if (numericCode) df1Map[numericCode] = row;
        }
      });

      const order = [...new Set(df1.map((row) => String(row["کد عرضه"])))]
        .filter((x) => x && x !== "undefined");

      const df1HasOfferQty = df1.some((r) =>
        Object.prototype.hasOwnProperty.call(r, "مقدار عرضه")
      );
      const df1HasDemandQty = df1.some((r) =>
        Object.prototype.hasOwnProperty.call(r, "مقدار تقاضا")
      );

      // تشخیص ستون قیمت فروشنده در فایل اول (حتی با فاصله یا فونت متفاوت)
      const df1SellerPriceKey = findColumnKey(df1, isSellerPriceCol);
      const df1HasSellerPrice = Boolean(df1SellerPriceKey);

      const data2 = await file2.arrayBuffer();
      const wb2 = XLSX.read(data2, { type: "array" });
      const ws2 = wb2.Sheets[wb2.SheetNames[0]];
      const df2 = XLSX.utils.sheet_to_json(ws2);

      const df2Map = {};
      df2.forEach((r) => {
        const code = String(r["کد عرضه"] || "").trim();
        const numericCode = code.replace(/\D/g, "");
        const finalKey = numericCode || code;
        if (!df2Map[finalKey]) df2Map[finalKey] = [];
        df2Map[finalKey].push(r);
      });

      // تشخیص ستون قیمت فروشنده در فایل دوم در صورت وجود
      const df2SellerPriceKey = findColumnKey(df2, isSellerPriceCol);

      const keepColumns = [
        "عرضه",
        "تقاضا",
        "نام کالا",
        "تولید کننده",
        "نام عرضه کننده",
        "نام مشتری",
        "محموله",
        "قیمت پایه", // مستقیماً از قیمت پایه استفاده می‌کنیم
      ];

      const sampleRowDf2 = df2[0] || {};
      const availableCols = keepColumns.filter((col) => {
        if (col === "قیمت پایه") return df1HasSellerPrice || Boolean(df2SellerPriceKey);
        if (col === "عرضه" && df1HasOfferQty) return true;
        if (col === "تقاضا" && df1HasDemandQty) return true;
        return Object.prototype.hasOwnProperty.call(sampleRowDf2, col);
      });

      //-----------------------------------------
      // فایل سوم – آمار معاملات
      //-----------------------------------------
      const data3 = await file3.arrayBuffer();
      const wb3 = XLSX.read(data3, { type: "array" });
      const ws3 = wb3.Sheets[wb3.SheetNames[0]];
      const df3 = XLSX.utils.sheet_to_json(ws3);

      const df3Map = {};
      df3.forEach((row) => {
        const contractType = String(row["نوع قرارداد"] || "").trim();

        if (contractType === "نقدی (مچینگ)") return;

        const product = String(row["نام کالا"] || "").trim();
        const producer = String(row["تولید کننده"] || "").trim();
        const supplier3 = String(row["عرضه کننده"] || "").trim();

        const keys = [];
        if (product && producer) keys.push(product + "|" + producer);
        if (product && supplier3) keys.push(product + "|" + supplier3);

        keys.forEach((key) => {
          if (!df3Map[key]) df3Map[key] = [];

          const date = row["تاریخ معامله"];
          if (date) {
            const exists = df3Map[key].some((d) => d.date === date);
            if (exists) return;

            df3Map[key].push({
              date: row["تاریخ معامله"],
              high: row["بالاترین"],
              low: row["پایین ترین"],
              avg: row["قیمت پایانی میانگین موزون"],
              base: row["قیمت پایه عرضه"],
              max: row["سقف قیمت"],
              offer: row["حجم عرضه"],
              demand: row["تقاضا"],
            });
          }
        });
      });

      //-----------------------------------------
      // پردازش اصلی
      //-----------------------------------------
      let result = [];

      order.forEach((code) => {
        const subset = (df2Map[code] || []).map((row) => {
          let filtered = {};

          let tradeDate = "";
          if (row["تاریخ عرضه"]) {
            tradeDate = row["تاریخ عرضه"];
          } else if (row["زمان عرضه"]) {
            const timeVal = String(row["زمان عرضه"]);
            tradeDate = timeVal.split(" ")[0];
          }

          filtered["تاریخ معامله"] = tradeDate;

          availableCols.forEach((col) => {
            let val = "";

            if (col === "عرضه" && df1HasOfferQty) {
              val = df1Map[code]?.["مقدار عرضه"] ?? row[col] ?? "";
            } else if (col === "تقاضا" && df1HasDemandQty) {
              val = df1Map[code]?.["مقدار تقاضا"] ?? row[col] ?? "";
            } else if (col === "قیمت پایه") {
              const cleanCode = String(code || "").trim();
              const numCode = cleanCode.replace(/\D/g, "");
              const row1 = df1Map[cleanCode] || df1Map[numCode];

              const p1 = df1SellerPriceKey && row1 ? row1[df1SellerPriceKey] : undefined;
              const p2 = df2SellerPriceKey ? row[df2SellerPriceKey] : undefined;

              val = (p1 !== undefined && p1 !== "") ? p1 : ((p2 !== undefined && p2 !== "") ? p2 : "");
            }
else {
              val = row[col] ?? "";
            }

            filtered[col] = isNumeric(val)
              ? Number(String(val).replace(/,/g, ""))
              : val;
          });

          return filtered;
        });

        if (subset.length > 0) {
          result.push(...subset);

          const firstRow = subset[0];
          const demand = Number(firstRow["تقاضا"] ?? 0);
          const supply = Number(firstRow["عرضه"] ?? 0);

          if (!isNaN(demand) && !isNaN(supply) && demand > supply) {
            result.push({
              "تاریخ معامله": "تاریخ معامله",
              "عرضه": "عرضه",
              "تقاضا": "تقاضا",
              "نسبت رقابت": "نسبت رقابت",
              "نام کالا": "بالاترین",
              "تولید کننده": "پایین‌ ترین",
              "نام عرضه کننده": "پایین‌ ترین",
              "نام مشتری": "متوسط",
              "قیمت پایه": "قیمت پایه",
              "سقف": "سقف پیشنهادی",
            });

            const product = String(firstRow["نام کالا"] || "").trim();
            let producer = String(firstRow["تولید کننده"] || "").trim();
            const supplier = String(firstRow["نام عرضه کننده"] || "").trim();
            if (!producer) producer = supplier;

            const key = product + "|" + producer;
            const dates = df3Map[key] || [];
            dates.sort((a, b) => {
              const da = new Date(a.date);
              const db = new Date(b.date);
              return db - da;
            });

            dates.forEach((item) => {
              result.push({
                "تاریخ معامله": item.date,
                "عرضه": isNumeric(item.offer) ? Number(String(item.offer).replace(/,/g, "")) : "",
                "تقاضا": isNumeric(item.demand) ? Number(String(item.demand).replace(/,/g, "")) : "",
                "نام کالا": isNumeric(item.high) ? Number(String(item.high).replace(/,/g, "")) : "",
                "تولید کننده": isNumeric(item.low) ? Number(String(item.low).replace(/,/g, "")) : "",
                "نام عرضه کننده": isNumeric(item.low) ? Number(String(item.low).replace(/,/g, "")) : "",
                "نام مشتری": isNumeric(item.avg) ? Number(String(item.avg).replace(/,/g, "")) : "",
                "قیمت پایه": isNumeric(item.base) ? Number(String(item.base).replace(/,/g, "")) : "",
                "سقف": isNumeric(item.max) ? Number(String(item.max).replace(/,/g, "")) : "",
              });
            });

            result.push({ "__blank__": true });
          } else {
            result.push({ "__blank__": true });
          }
        }
      });

      //-----------------------------------------
      // ساخت اکسل خروجی
      //-----------------------------------------
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
      const renamedCols = availableCols; // نیازی به رینیم مجدد نیست

      const allCols = [
        "تاریخ معامله",
        ...renamedCols,
        ...extraCols,
        ...priceColumns,
      ];

      const demandIndex = allCols.indexOf("تقاضا");
      if (demandIndex !== -1) {
        allCols.splice(demandIndex + 1, 0, "نسبت رقابت");
      }

      const ratioColIndex = allCols.indexOf("نسبت رقابت") + 1;
      if (ratioColIndex > 0) {
        sheet.getColumn(ratioColIndex).numFmt = "0.00";
      }

      sheet.addRow(allCols);

      result.forEach((row) => {
        if (row["__blank__"]) {
          sheet.addRow(allCols.map(() => ""));

          const headerLikeRow = allCols.map((c) => c);
          const newRow = sheet.addRow(headerLikeRow);

          newRow.eachCell((cell) => {
            cell.font = { name: "B Nazanin", bold: true, size: 12 };
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFCCCCCC" },
            };
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
        } else {
          const newRow = sheet.addRow(allCols.map((c) => row[c] ?? ""));
          const rowNumber = newRow.number;

          const offerVal = row["عرضه"];
          const demandVal = row["تقاضا"];
          const diffCol = allCols.indexOf("نسبت رقابت") + 1;

          const prevRow = sheet.getRow(rowNumber - 1);
          const prevOffer = prevRow ? prevRow.getCell(allCols.indexOf("عرضه") + 1).value : null;
          const prevDemand = prevRow ? prevRow.getCell(allCols.indexOf("تقاضا") + 1).value : null;

          const isSameGroup =
            prevOffer === offerVal &&
            prevDemand === demandVal &&
            prevOffer !== null &&
            prevDemand !== null;

          if (!isSameGroup) {
            if (isNumeric(offerVal) && isNumeric(demandVal)) {
              const offerCol = getExcelColumn(allCols.indexOf("عرضه") + 1);
              const demandCol = getExcelColumn(allCols.indexOf("تقاضا") + 1);

              sheet.getRow(rowNumber).getCell(diffCol).value = {
                formula: `${demandCol}${rowNumber}/${offerCol}${rowNumber}`,
              };
            }
          } else {
            const firstRowValue = prevRow.getCell(diffCol).value;
            sheet.getRow(rowNumber).getCell(diffCol).value =
              typeof firstRowValue === "object" && firstRowValue.formula
                ? ""
                : firstRowValue;
          }
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

        const rowValues = row.values.slice(1);
        const isHeaderLikeRow = rowValues.every((v, i) => v === allCols[i]);
        const statsHeader =
          row.getCell(1).value === "تاریخ معامله" &&
          row.getCell(allCols.indexOf("نام کالا") + 1).value === "بالاترین";

        const isStatsRow =
          !statsHeader &&
          isNumeric(row.getCell(allCols.indexOf("نام کالا") + 1).value);

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
          const isPriceColumn = colName.match(/^قیمت[1-7]$/);

          if (!isStatsRow) {
            if (colName === "قیمت2") {
              const p7 = getExcelColumn(allCols.indexOf("قیمت7") + 1);
              const p1 = getExcelColumn(allCols.indexOf("قیمت1") + 1);
              cell.value = {
                formula: `IFERROR((${p7}${rowNumber}-${p1}${rowNumber})/6+${p1}${rowNumber},0)`,
              };
            } else if (colName === "قیمت7") {
              const سقفColLetter = getExcelColumn(allCols.indexOf("سقف") + 1);
              cell.value = { formula: `${سقفColLetter}${rowNumber}` };
            } else if (colName.match(/^قیمت[3-6]$/)) {
              const n = parseInt(colName.replace("قیمت", ""));
              const prev = getExcelColumn(allCols.indexOf(`قیمت${n - 1}`) + 1);
              const p7 = getExcelColumn(allCols.indexOf("قیمت7") + 1);
              const p1 = getExcelColumn(allCols.indexOf("قیمت1") + 1);
              cell.value = {
                formula: `IFERROR((${p7}${rowNumber}-${p1}${rowNumber})/6+${prev}${rowNumber},0)`,
              };
            }
          }

          if (!isStatsRow || !isPriceColumn) {
            const numFmtCols = [
              "قیمت1",
              "قیمت2",
              "قیمت3",
              "قیمت4",
              "قیمت5",
              "قیمت6",
              "قیمت7",
              "سقف",
              "قیمت پایه",
              "عرضه",
              "تقاضا",
              "نام کالا",
              "تولید کننده",
              "نام عرضه کننده",
              "نام مشتری",
            ];
            if (numFmtCols.includes(colName)) {
              cell.numFmt = "#,##0";
            }
          }

          if (isHeaderLikeRow) {
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFCCCCCC" },
            };
            cell.font = { name: fontName, bold: true, size: 12 };
          } else if (statsHeader) {
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFCCE5FF" },
            };
            cell.font = { name: fontName, bold: true, size: 12 };
          } else if (colName === "سقف" || colName === "سقف پیشنهادی") {
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFD3D3D3" },
            };
            cell.font = {
              name: fontName,
              size: 12,
              bold: true,
              color: { argb: "FFFF0000" },
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

          if (isStatsRow && isPriceColumn) {
            cell.border = {};
          } else {
            cell.border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }

          cell.font = cell.font || { name: fontName, size: 12 };
          cell.alignment = {
            vertical: "middle",
            horizontal: "center",
            wrapText: true,
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

      const ratioIdx = allCols.indexOf("نسبت رقابت") + 1;
      if (ratioIdx > 0) {
        sheet.getColumn(ratioIdx).width = 15;
      }

      [
        "تاریخ معامله",
        "سقف",
        "قیمت1",
        "قیمت2",
        "قیمت3",
        "قیمت4",
        "قیمت5",
        "قیمت6",
        "قیمت7",
      ].forEach((colName) => {
        const idx = allCols.indexOf(colName);
        if (idx !== -1) sheet.getColumn(idx + 1).width = 15;
      });

      const mergeCols = [
        "عرضه",
        "تقاضا",
        "نسبت رقابت",
        "نام کالا",
        "تولید کننده",
        "نام عرضه کننده",
        "قیمت پایه",
        "تاریخ معامله",
      ];

      mergeCols.forEach((col) => {
        const idx = allCols.indexOf(col) + 1;
        if (!idx) return;

        let startRow = 2;
        let isCurrentColMerged = false;

        const isStatsRow = (rowNum) => {
          const productNameCell = sheet
            .getRow(rowNum)
            .getCell(allCols.indexOf("نام کالا") + 1);
          return (
            isNumeric(productNameCell.value) &&
            productNameCell.value !== null &&
            productNameCell.value !== ""
          );
        };

        for (let r = startRow; r <= sheet.rowCount; r++) {
          const currentRow = sheet.getRow(r);
          const currentCell = currentRow.getCell(idx);
          const currentValue = currentCell.value;

          if (isStatsRow(r)) {
            if (col === "نسبت رقابت") {
              if (isCurrentColMerged && r - 1 > startRow) {
                sheet.mergeCells(startRow, idx, r - 1, idx);
              }
              isCurrentColMerged = false;
              startRow = r + 1;
              continue;
            }

            if (isCurrentColMerged && r - 1 > startRow) {
              sheet.mergeCells(startRow, idx, r - 1, idx);
            }
            isCurrentColMerged = false;
            startRow = r + 1;
            continue;
          }

          if (r === startRow) {
            startRow = r;
            isCurrentColMerged = true;
            continue;
          }

          const prevRow = sheet.getRow(r - 1);
          const prevCell = prevRow.getCell(idx);
          const prevValue = prevCell.value;

          let breakMerge = false;

          if (col === "نسبت رقابت") {
            const offerIdx = allCols.indexOf("عرضه") + 1;
            const demandIdx = allCols.indexOf("تقاضا") + 1;

            const currentOffer = currentRow.getCell(offerIdx).value;
            const currentDemand = currentRow.getCell(demandIdx).value;

            const prevOffer = sheet.getRow(r - 1).getCell(offerIdx).value;
            const prevDemand = sheet.getRow(r - 1).getCell(demandIdx).value;

            if (currentOffer !== prevOffer || currentDemand !== prevDemand) {
              breakMerge = true;
            }
          } else {
            breakMerge =
              currentValue !== prevValue ||
              currentValue === null ||
              currentValue === "";
          }

          if (breakMerge) {
            if (isCurrentColMerged && r - 1 > startRow) {
              sheet.mergeCells(startRow, idx, r - 1, idx);
            }
            startRow = r;
            isCurrentColMerged = true;
            continue;
          }
        }

        if (
          isCurrentColMerged &&
          sheet.rowCount >= startRow &&
          !isStatsRow(sheet.rowCount)
        ) {
          if (sheet.rowCount > startRow) {
            sheet.mergeCells(startRow, idx, sheet.rowCount, idx);
          }
        }
      });

      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), "اکسل_مرتب.xlsx");
      showAlert("فایل آماده شد و دانلود شد.");
      setFile1(null);
      setFile2(null);
      setFile3(null);
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
        <label htmlFor="file1" className="custom-file-btn">
          بارگزاری شود tts فایل
        </label>
        <input
          id="file1"
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFile1}
          className="form-control mb-2"
        />
        <p className="file-name">
          {file1 ? file1.name : "هیچ فایلی انتخاب نشده"}
        </p>
      </div>

      <div className="file-input">
        <label htmlFor="file2" className="custom-file-btn">
          فایل خروجی ثبت سفارش‌های سیمرغ وارد شود
        </label>
        <input
          id="file2"
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFile2}
          className="form-control mb-2"
        />
        <p className="file-name">
          {file2 ? file2.name : "هیچ فایلی انتخاب نشده"}
        </p>
      </div>

      <div className="file-input">
        <label htmlFor="file3" className="custom-file-btn">
          فایل آمار معاملات بارگزاری شود
        </label>
        <input
          id="file3"
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFile3}
          className="form-control mb-2"
        />
        <p className="file-name">
          {file3 ? file3.name : "هیچ فایلی انتخاب نشده"}
        </p>
      </div>

      <button
        className="btn btn-primary"
        onClick={processFiles}
        disabled={loading || !file1 || !file2 || !file3}
      >
        {loading ? "در حال پردازش..." : "دانلود اکسل مرتب‌شده"}
      </button>

      {alertMsg && <div className="custom-alert">{alertMsg}</div>}
    </div>
  );
}

export default Siman;
