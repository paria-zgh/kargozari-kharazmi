// Enerzhi.js
import React, { useState } from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import jalaali from "jalaali-js";

const AmarEnergy = () => {
  const [data, setData] = useState([]);
  const [domesticRows, setDomesticRows] = useState([]);
  const [internationalRows, setInternationalRows] = useState([]);
  const [otherRows, setOtherRows] = useState([]);

  const [titleText] = useState("آمار انرژی");
  const [SuplyText] = useState("آمار معاملات ");

  // ستون نوع معامله به لیست اصلی ستون‌ها قبل از درصد رقابت اضافه شد
  const columns = [
    "نام تجاری کالا",
    "عرضه کننده",
    "واحد عرضه",
    "حجم عرضه",
    "حجم تقاضا",
    "حجم کل معامله شده",
    "بیشترین قیمت معامله",
    "کمترین قیمت معامله",
    "متوسط قيمت معامله",
    "قیمت پایه",
    "قیمت مبنا",
    "مقصد",
    "نوع معامله",
  ];

  // ------------------------- نرمال‌سازی متن ----------------------------
  const normalizeText = (str) => {
    if (!str) return "";
    return str
      .toString()
      .replace(/ك/g, "ک")
      .replace(/ي/g, "ی")
      .replace(/\u200c/g, " ")
      .replace(/[()]/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  };

  // ------------------------- واترمارک ----------------------------
  const createWatermark = (text1, text2) => {
    const canvas = document.createElement("canvas");
    canvas.width = 3600;
    canvas.height = 1000;
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.translate(canvas.width / 2, canvas.height / 2);
    ctx.rotate((-8 * Math.PI) / 180);
    ctx.textAlign = "center";
    ctx.fillStyle = "rgba(140,140,140,0.25)";
    ctx.shadowColor = "rgba(0,0,0,0.25)";
    ctx.shadowBlur = 15;
    ctx.shadowOffsetX = 6;
    ctx.shadowOffsetY = 6;
    ctx.font = "bold 250px Arial";
    ctx.fillText(text1, 0, -180);
    ctx.font = "bold 230px Arial";
    ctx.fillText(text2, 0, 260);
    return canvas.toDataURL("image/png");
  };

  const fixProductName = (name) => {
    if (!name) return "";

    let normalized = name.replace(/ك/g, "ک").replace(/ي/g, "ی");

    // حذف هر چیزی بعد از "حاصل از"
    normalized = normalized.replace(/حاصل از.*/i, "").trim();

    // حذف پرانتزهایی که داخلشان "حاصل" است
    normalized = normalized
      .replace(/\(([^)]+)\)/g, (match, p1) => {
        if (p1.includes("حاصل")) return "";
        return match;
      })
      .trim();

    const stopWords = ["شرکت", "مجتمع"];
    const parts = normalized.split(" ");
    let filtered = [];

    for (let i = 0; i < parts.length; i++) {
      if (stopWords.includes(parts[i])) break;
      filtered.push(parts[i]);
    }

    const trimmed = name.trim();

    if (
      trimmed.endsWith("عمده") &&
      filtered.length &&
      filtered[filtered.length - 1] !== "(عمده)"
    ) {
      filtered.push("(عمده)");
    }

    return filtered.join(" ");
  };

  const formatPriceBase = (text) => {
    if (!text) return "";
    let str = text.toString();
    str = str.replace(/\\([^)]*\\)/g, "").trim();
    const match = str.match(/(-?\d+)(.*)/);
    if (match) {
      const isNegative = match[1].startsWith("-");
      let numberPart = match[1].replace("-", "");
      let rest = match[2].trim();
      return `${isNegative ? "-" : ""}${numberPart}${rest}`;
    }
    return str;
  };

  const extractCellValue = (v) => {
    if (v === null || v === undefined) return "";

    // RichText اکسل
    if (typeof v === "object" && v.richText) {
      return v.richText.map((t) => t.text).join("");
    }

    // اگر سلول فرمول داشت و نتیجه هم داشت
    if (typeof v === "object" && v.formula !== undefined) {
      // ExcelJS معمولاً نتیجه رو در v.result می‌گذارد
      if (v.result !== undefined && v.result !== null) return v.result;
      return "";
    }

    // صفر واقعی
    if (v === 0 || v === "0") return 0;

    return v;
  };

  // تبدیل عدد فارسی به انگلیسی
  const faDigitsToEn = (s) =>
    s.replace(/[۰-۹]/g, (d) => "۰۱۲۳۴۵۶۷۸۹".indexOf(d));

  // استخراج «عدد واقعی» از سلول یک ستون خاص در همان row
  const getCleanNumberFromRow = (row, headerMap, colKey) => {
    const cIndex = headerMap[colKey];
    if (!cIndex) return null;

    const raw = extractCellValue(row.getCell(cIndex).value);
    if (raw === null || raw === undefined || raw === "") return 0;

    // اگر همینجا عدد بود
    if (typeof raw === "number") return raw;

    // رشته: پاکسازی کامل
    const s = faDigitsToEn(raw.toString())
      .replace(/,/g, "")
      .replace(/\u200c/g, "")
      .replace(/\s+/g, "")
      .trim();

    // اگر عملاً چیزی نمانده
    if (!s) return 0;

    // فقط رقم/نقطه/منفی
    const cleaned = s.replace(/[^\d.-]/g, "");
    if (!cleaned) return 0;

    const num = Number(cleaned);
    return Number.isFinite(num) ? num : 0;
  };

  // ------------------------------------- خواندن فایل اکسل -------------------------------------
  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = async (evt) => {
      const buffer = evt.target.result;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      const worksheet = workbook.worksheets[0];
      const headerRow = worksheet.getRow(1);

      // --- ۱) ساخت مپ هدرهای واقعی فایل ---
      const realHeaders = {};
      headerRow.eachCell((cell, colNumber) => {
        const norm = normalizeText(cell.value);
        realHeaders[norm] = colNumber;
      });

      // --- ۲) alias برای انواع اسامی ستون‌ها ---
      const columnAliases = {
        "نام تجاری کالا": ["نام تجاری کالا"],
        "عرضه کننده": ["عرضه کننده"],
        "واحد عرضه": ["واحد عرضه"],
        "حجم عرضه": ["حجم عرضه", "حجم عرضه بر حسب واحد عرضه"],
        "حجم تقاضا": ["حجم تقاضا", "تقاضا(بر حسب واحد عرضه)", "تقاضا بر حسب واحد عرضه"],
        "حجم کل معامله شده": [
          "حجم کل معامله شده بر حسب واحد عرضه",
          "حجم كل معامله شده بر حسب واحد عرضه",
        ],
        "کمترین قیمت معامله": ["کمترین قیمت معامله"],
        "متوسط قيمت معامله": ["متوسط قيمت معامله", "متوسط قیمت معامله"],
        "بیشترین قیمت معامله": ["بیشترین قیمت معامله"],
        "قیمت مبنا": [
          "قیمت مبنا",
          "قيمت مبنا",
          "قيمت مبنا جزئيات محاسبه قيمت نهايي معاملات کشف پریمیوم",
          "قیمت مبنا جزئیات محاسبه",
        ],
        "قیمت پایه": ["قیمت پایه", "قیمت پایه (ریال)"],
        "مقصد": ["مقصد"],
        "نوع معامله": ["نوع معامله", "نوع معاملـه"],
      };

      // --- ۳) ساخت مپ نهایی ستون‌ها بر اساس alias ---
      const headerMap = {};
      Object.keys(columnAliases).forEach((mainKey) => {
        const aliases = columnAliases[mainKey];
        for (let alias of aliases) {
          const normAlias = normalizeText(alias);
          const found = Object.keys(realHeaders).find((h) => h === normAlias);
          if (found) {
            headerMap[mainKey] = realHeaders[found];
            break;
          }
        }
      });

      // --- ۴) خواندن داده‌ها + حذف ردیف‌هایی که حجم عرضه 0/خالی دارند ---
      const jsonData = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;

        // اگر ستون "حجم عرضه" اصلاً پیدا نشد، بهتره چیزی رو حذف نکنیم
        if (!headerMap["حجم عرضه"]) {
          // fallback: بدون فیلتر
        } else {
          const supplyNum = getCleanNumberFromRow(row, headerMap, "حجم عرضه");

          // شرط حذف قطعی
          if (supplyNum === 0) {
            return; // این ردیف کلاً حذف می‌شود
          }
        }

        const rowData = {};

        Object.keys(columnAliases).forEach((col) => {
          const colIndex = headerMap[col];
          let value = colIndex ? extractCellValue(row.getCell(colIndex).value) : "";

          if (col === "نام تجاری کالا") value = fixProductName(value);

          rowData[col] = value;
        });

        // اگر نام کالا خالی بود، حذف
        if (!rowData["نام تجاری کالا"]) return;

        jsonData.push(rowData);
      });

      // --- ۵) دسته‌بندی داخلی / بین‌الملل ---
      const domestic = [],
        international = [],
        other = [];

      jsonData.forEach((row) => {
        const dest = normalizeText(row["مقصد"]);
        if (dest.includes("داخلی")) domestic.push(row);
        else if (dest.includes("بین")) international.push(row);
        else other.push(row);
      });

      // سورت بر اساس نام تجاری کالا
      domestic.sort(sortByProductName);
      international.sort(sortByProductName);
      other.sort(sortByProductName);

      setDomesticRows(domestic);
      setInternationalRows(international);
      setOtherRows(other);

      // --- ۶) ساخت جدول نمایش صفحه ---
      const emptyRow = Object.fromEntries(columns.map((c) => [c, ""]));
      const createHeaderRow = () => {
        const r = {};
        columns.forEach((col) => (r[col] = col));
        r.__isHeader = true;
        return r;
      };

      const finalData = [];
      finalData.push(createHeaderRow());
      if (domestic.length > 0) finalData.push(...domestic);
      finalData.push(emptyRow);
      finalData.push(createHeaderRow());
      if (international.length > 0) finalData.push(...international);
      if (other.length > 0) finalData.push(...other);

      setData(finalData);
    };

    reader.readAsArrayBuffer(file);
  };

  const numericColumns = [
    "حجم عرضه",
    "حجم تقاضا",
    "کمترین قیمت معامله",
    "متوسط قيمت معامله",
    "بیشترین قیمت معامله",
    "حجم کل معامله شده",
    "قیمت پایه",
    "قیمت مبنا",
  ];

  const toNumber = (val) => {
    if (val === null || val === undefined || val === "") return null;
    if (typeof val === "number") return val;

    const cleaned = faDigitsToEn(val.toString())
      .replace(/,/g, "")
      .replace(/[^\d.-]/g, "");

    const num = Number(cleaned);
    return isNaN(num) ? null : num;
  };

  const sortByProductName = (a, b) => {
    const nameA = (a["نام تجاری کالا"] || "")
      .toString()
      .replace(/ي/g, "ی")
      .replace(/ك/g, "ک")
      .trim();

    const nameB = (b["نام تجاری کالا"] || "")
      .toString()
      .replace(/ي/g, "ی")
      .replace(/ك/g, "ک")
      .trim();

    return nameA.localeCompare(nameB, "fa");
  };

  // -------------------------------------------------- تابع خروجی اکسل --------------------------------------------------
  const exportExcel = async () => {
    if (data.length === 0) return;

    const workbook = new ExcelJS.Workbook();

    const autoFitColumns = (worksheet, cols) => {
      const skipRows = 2;
      cols.forEach((col, index) => {
        let maxLength = 1;
        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
          if (rowNumber <= skipRows) return;
          const cell = row.getCell(index + 1);
          let cellValue = cell.value;
          if (cellValue && typeof cellValue === "object" && cellValue.richText) {
            cellValue = cellValue.richText.map((t) => t.text).join("");
          } else if (cellValue && typeof cellValue !== "string") {
            cellValue = cellValue.toString();
          }
          if (cellValue) {
            const cellLength = cellValue.length;
            if (cellLength > maxLength) maxLength = cellLength;
          }
        });
        worksheet.getColumn(index + 1).width = Math.max(maxLength + 4, 12);
      });
    };

    const addMergedHeader = (worksheet, cols, topText, secondText) => {
      const lastColLetter = String.fromCharCode(64 + cols.length);
      worksheet.mergeCells(`A1:${lastColLetter}1`);
      const topCell = worksheet.getCell("A1");
      topCell.value = topText;
      topCell.font = { name: "Arial", bold: true, size: 14 };
      topCell.alignment = { horizontal: "center", vertical: "middle" };
      topCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0FAF8E" } };
      worksheet.getRow(1).height = 30;
      worksheet.getRow(1).eachCell((cell) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });

      worksheet.mergeCells(`A2:${lastColLetter}2`);
      const secondCell = worksheet.getCell("A2");
      secondCell.value = secondText;
      secondCell.font = { name: "Arial", bold: true, size: 18 };
      secondCell.alignment = { horizontal: "center", vertical: "middle" };
      secondCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3FD9B8" } };
      worksheet.getRow(2).height = 40;
      worksheet.getRow(2).eachCell((cell) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });

      const headerRow = worksheet.addRow(cols.map((col) => col));
      headerRow.eachCell((cell) => {
        cell.font = { name: "Arial", bold: true, size: 12 };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFCAC8C9" } };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
      worksheet.getRow(headerRow.number).height = 28;

      return headerRow.number;
    };

    const addMergedHeaderWithColors = (worksheet, cols, topText, secondText, colors) => {
      const lastColLetter = String.fromCharCode(64 + cols.length);

      worksheet.mergeCells(`A1:${lastColLetter}1`);
      const topCell = worksheet.getCell("A1");
      topCell.value = topText;
      topCell.font = { name: "Arial", bold: true, size: 14 };
      topCell.alignment = { horizontal: "center", vertical: "middle" };
      topCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.header1 } };
      worksheet.getRow(1).height = 33;
      worksheet.getRow(1).eachCell((cell) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });

      worksheet.mergeCells(`A2:${lastColLetter}2`);
      const secondCell = worksheet.getCell("A2");
      secondCell.value = secondText;
      secondCell.font = { name: "Arial", bold: true, size: 18 };
      secondCell.alignment = { horizontal: "center", vertical: "middle" };
      secondCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.header2 } };
      worksheet.getRow(2).height = 40;
      worksheet.getRow(2).eachCell((cell) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });

      const headerRow = worksheet.addRow(cols.map((col) => col));
      headerRow.eachCell((cell) => {
        cell.font = { name: "Arial", bold: true, size: 12 };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
      worksheet.getRow(headerRow.number).height = 28;

      return headerRow.number;
    };

    const getTodayShamsi = () => {
      const today = new Date();
      const { jy, jm, jd } = jalaali.toJalaali(today);
      const dayNames = ["یکشنبه", "دوشنبه", "سه‌شنبه", "چهارشنبه", "پنجشنبه", "جمعه", "شنبه"];
      const dayOfWeek = dayNames[today.getDay()];
      return {
        day: dayOfWeek,
        date: `${jy}/${jm.toString().padStart(2, "0")}/${jd.toString().padStart(2, "0")}`,
      };
    };

    const { day, date } = getTodayShamsi();

    // =================== شیت داخلی ===================
    if (domesticRows.length > 0) {
      const internalColumns = [
        "نام تجاری کالا",
        "عرضه کننده",
        "واحد عرضه",
        "حجم عرضه",
        "حجم تقاضا",
        "حجم کل معامله شده",
        "کمترین قیمت معامله",
        "متوسط قيمت معامله",
        "بیشترین قیمت معامله",
        "قیمت پایه",
        "قیمت مبنا",
        "مقصد",
        "نوع معامله",
        "درصد رقابت",
      ];

      const domesticSheet = workbook.addWorksheet("داخلی", {
        properties: { defaultRowHeight: 25 },
        views: [{ rightToLeft: true }],
      });

      const displayInternalColumns = internalColumns.map((col) => {
        if (col === "حجم عرضه بر حسب واحد عرضه") return "حجم عرضه";
        if (col === "تقاضا(بر حسب واحد عرضه)") return "حجم تقاضا";
        if (col === "قيمت مبنا جزئيات محاسبه قيمت نهايي معاملات کشف پریمیوم") return "قيمت مبنا";
        return col;
      });

      const headerRowNumber = addMergedHeader(
        domesticSheet,
        displayInternalColumns,
        "کارگزاری آینده نگر خوارزمی به مدیریت دکتر ذوقی  09123011311 و 61914000-021 (داخلی 401)",
        `آمار معاملات محصولات بورس انرژی در رینگ داخلی روز ${day} مورخ ${date}`
      );

      const wmBase64 = createWatermark("کارگزاری آینده نگر خوارزمی", "09123011311");
      const wmId = workbook.addImage({ base64: wmBase64, extension: "png" });

      domesticSheet.addImage(wmId, {
        tl: { col: 2, row: 2 },
        ext: { width: 900, height: 300 },
        editAs: "oneCell",
      });

      domesticRows.forEach((rowData, i) => {
        const cleanedRow = internalColumns.map((col) => {
          if (col === "درصد رقابت") return "";
          const val = rowData[col];
          if (val === "" || val === null || val === undefined) return "";
          if (val === 0 || val === "0") return 0;
          return val;
        });

        const row = domesticSheet.addRow(cleanedRow);
        const rowNum = row.number;

        const avgPriceColLetter = String.fromCharCode(
          64 + internalColumns.indexOf("متوسط قيمت معامله") + 1
        );
        const basePriceColLetter = String.fromCharCode(
          64 + internalColumns.indexOf("قیمت پایه") + 1
        );
        const compColIndex = internalColumns.indexOf("درصد رقابت") + 1;

        row.getCell(compColIndex).value = {
          formula: `IFERROR(IF(AND(ISNUMBER(${avgPriceColLetter}${rowNum}),ISNUMBER(${basePriceColLetter}${rowNum}),${basePriceColLetter}${rowNum}<>0),(${avgPriceColLetter}${rowNum}-${basePriceColLetter}${rowNum})/${basePriceColLetter}${rowNum},0),0)`,
        };
        row.getCell(compColIndex).numFmt = "[=0]0%;0.0%";

        const fillColor = i % 2 === 0 ? "FFFFFFFF" : "FFF5F5F5";

        row.eachCell((cell, colNumber) => {
          const colName = internalColumns[colNumber - 1];

          if (numericColumns.includes(colName)) {
            const numValue = toNumber(cell.value);
            if (numValue !== null) {
              cell.value = numValue;
              cell.numFmt = "#,##0";
            }
          }

          if (colName === "درصد رقابت") {
            cell.numFmt = "[=0]0%;0.0%";
          }

          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fillColor } };
          cell.font = { name: "Arial", bold: true };
          cell.alignment = { horizontal: "center", vertical: "middle" };
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        });
      });

      const internalCompColLetter = String.fromCharCode(64 + internalColumns.indexOf("درصد رقابت") + 1);
      domesticSheet.addConditionalFormatting({
        ref: `${internalCompColLetter}${headerRowNumber + 1}:${internalCompColLetter}${domesticSheet.rowCount}`,
        rules: [
          {
            type: "cellIs",
            operator: "greaterThan",
            formulae: ["0"],
            style: { font: { color: { argb: "FF008000" }, bold: true } },
          },
        ],
      });

      domesticSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber > headerRowNumber) row.height = 25;
      });

      autoFitColumns(domesticSheet, internalColumns);

      const priceMabnaIndex = internalColumns.indexOf("قیمت مبنا") + 1;
      let isColumnEmpty = true;

      domesticSheet.eachRow((row, rowNumber) => {
        if (rowNumber <= headerRowNumber) return;
        const val = row.getCell(priceMabnaIndex).value;
        if (val !== "" && val !== null && val !== 0 && val !== "0") isColumnEmpty = false;
      });

      if (isColumnEmpty) {
        domesticSheet.getColumn(priceMabnaIndex).hidden = true;
      }
    }

    // =================== شیت بین الملل ===================
    let mainColumns = [...columns, "درصد رقابت"];
    const kashfPremiumRows = [...internationalRows, ...otherRows].filter((row) =>
      row["نوع معامله"]?.includes("کشف پریمیوم")
    );
    const otherNonPremiumRows = [...internationalRows, ...otherRows].filter(
      (row) => !row["نوع معامله"]?.includes("کشف پریمیوم")
    );

    const mainDataRows = [...kashfPremiumRows, ...otherNonPremiumRows];

    const basePriceKey = "قيمت مبنا جزئيات محاسبه قيمت نهايي معاملات کشف پریمیوم";
    const priceColIndex = mainColumns.findIndex((c) => c === basePriceKey);

    if (priceColIndex !== -1) {
      const allEmpty = mainDataRows.every((row) => !row[basePriceKey] || row[basePriceKey].toString().trim() === "");
      if (allEmpty) mainColumns.splice(priceColIndex, 1);
    }

    const mainSheet = workbook.addWorksheet("بین الملل", {
      properties: { defaultRowHeight: 25 },
      views: [{ rightToLeft: true }],
    });

    const displayColumnsMap = (col) => {
      if (col === "حجم عرضه بر حسب واحد عرضه") return "حجم عرضه";
      if (col === "تقاضا(بر حسب واحد عرضه)") return "حجم تقاضا";
      if (col === "قيمت مبنا جزئيات محاسبه قيمت نهايي معاملات کشف پریمیوم") return "قيمت مبنا";
      return col;
    };

    const headerRowNumber = addMergedHeaderWithColors(
      mainSheet,
      mainColumns.map(displayColumnsMap),
      "کارگزاری آینده نگر خوارزمی به مدیریت دکتر ذوقی 09123011311 و 61914000-021 (داخلی 401)",
      `آمار معاملات محصولات بورس انرژی در رینگ صادراتی روز ${day} مورخ ${date}`,
      { header1: "FFFF6A00", header2: "FFFFB873", header3: "FFD9D9D9" }
    );

    const wmBase64 = createWatermark("کارگزاری آینده نگر خوارزمی", "09123011311");
    const wmId = workbook.addImage({ base64: wmBase64, extension: "png" });

    mainSheet.addImage(wmId, {
      tl: { col: 2, row: 2 },
      ext: { width: 900, height: 300 },
      editAs: "oneCell",
    });

    mainDataRows.forEach((rowData, i) => {
      const rowDataCopy = { ...rowData };
      if (rowDataCopy["قیمت پایه"]) rowDataCopy["قیمت پایه"] = formatPriceBase(rowDataCopy["قیمت پایه"]);

      const cleanedRow = mainColumns.map((col) => {
        if (col === "درصد رقابت") return "";
        const val = rowDataCopy[col];
        if (val === "" || val === null || val === undefined) return "";
        if (val === 0 || val === "0") return 0;
        return val;
      });

      const row = mainSheet.addRow(cleanedRow);
      const rowNum = row.number;

      const avgPriceColLetter = String.fromCharCode(64 + mainColumns.indexOf("متوسط قيمت معامله") + 1);
      const basePriceColLetter = String.fromCharCode(64 + mainColumns.indexOf("قیمت پایه") + 1);
      const compColIndex = mainColumns.indexOf("درصد رقابت") + 1;

      row.getCell(compColIndex).value = {
        formula: `IFERROR(IF(AND(ISNUMBER(${avgPriceColLetter}${rowNum}),ISNUMBER(${basePriceColLetter}${rowNum}),${basePriceColLetter}${rowNum}<>0),(${avgPriceColLetter}${rowNum}-${basePriceColLetter}${rowNum})/${basePriceColLetter}${rowNum},0),0)`,
      };
      row.getCell(compColIndex).numFmt = "[=0]0%;0.0%";

      const fillColor = i % 2 === 0 ? "FFFFFFFF" : "FFFAF7F3";

      row.eachCell((cell, colNumber) => {
        const colName = mainColumns[colNumber - 1];

        if (numericColumns.includes(colName)) {
          const numValue = toNumber(cell.value);
          if (numValue !== null) {
            cell.value = numValue;
            cell.numFmt = "#,##0";
          }
        }

        if (colName === "درصد رقابت") {
          cell.numFmt = "[=0]0%;0.0%";
        }

        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fillColor } };
        cell.font = { name: "Arial", bold: true };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
    });

    const mainCompColLetter = String.fromCharCode(64 + mainColumns.indexOf("درصد رقابت") + 1);
    mainSheet.addConditionalFormatting({
      ref: `${mainCompColLetter}${headerRowNumber + 1}:${mainCompColLetter}${mainSheet.rowCount}`,
      rules: [
        {
          type: "cellIs",
          operator: "greaterThan",
          formulae: ["0"],
          style: { font: { color: { argb: "FF008000" }, bold: true } },
        },
      ],
    });

    mainSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (rowNumber > headerRowNumber) row.height = 25;
    });

    autoFitColumns(mainSheet, mainColumns);

    const buf = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buf]), "Enerzhi_Output_Separated.xlsx");
  };

  // ------------------------------------- رابط کاربری (UI) ------------------------------------------
  return (
    <div style={{ padding: "20px", fontFamily: "Arial" }}>
      <h2>{titleText}</h2>
      <h4>{SuplyText}</h4>

      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      <button onClick={exportExcel} style={{ marginLeft: "10px" }}>
        خروجی اکسل
      </button>

      <table style={{ width: "100%", marginTop: "20px", borderCollapse: "collapse" }}>
        <thead>
          <tr>
            {columns.map((col, idx) => (
              <th
                key={idx}
                style={{
                  backgroundColor: "#D3D3D3",
                  padding: "8px",
                  border: "1px solid #ccc",
                }}
              >
                {col}
              </th>
            ))}
          </tr>
        </thead>

        <tbody>
          {data.map((row, idx) => (
            <tr key={idx}>
              {columns.map((col, i) => (
                <td
                  key={i}
                  style={{
                    padding: "8px",
                    border: "1px solid #ccc",
                    textAlign: "center",
                  }}
                >
                  {row[col]}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default AmarEnergy;
