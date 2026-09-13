// Enerzhi.js
import React, { useState } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import jalaali from 'jalaali-js';

const AmarEnergy = () => {
  const [data, setData] = useState([]);
  const [domesticRows, setDomesticRows] = useState([]);
  const [internationalRows, setInternationalRows] = useState([]);
  const [otherRows, setOtherRows] = useState([]);

  const [titleText] = useState('آمار انرژی');
  const [SuplyText] = useState('آمار معاملات ');

  const columns = [
    'نام تجاری کالا',
    'عرضه کننده',
    'واحد عرضه',
    'حجم عرضه',
    'حجم تقاضا',
    'بیشترین قیمت معامله',
    'کمترین قیمت معامله',
    'متوسط قيمت معامله',
    'قیمت مبنا',
    'قیمت پایه',
    'مقصد',
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
    ctx.rotate(-8 * Math.PI / 180);
    ctx.textAlign = "center";
    ctx.fillStyle = "rgba(0,70,180,0.25)";
    ctx.shadowColor = "rgba(0,0,0,0.25)";
    ctx.shadowBlur = 15;
    ctx.shadowOffsetX = 6;
    ctx.shadowOffsetY = 6;
    ctx.font = "bold 250px B Nazanin";
    ctx.fillText(text1, 0, -180);
    ctx.font = "bold 230px B Nazanin";
    ctx.fillText(text2, 0, 260);
    return canvas.toDataURL("image/png");
  };

  const fixProductName = (name) => {
    if (!name) return "";
    let normalized = name.replace(/ك/g, "ک").replace(/ي/g, "ی");
    normalized = normalized.replace(/\\(([^)]+)\\)/g, (match, p1) => {
      if (p1.includes("حاصل")) return "";
      return match;
    }).trim();
    const stopWords = ["شرکت", "مجتمع"];
    const parts = normalized.split(" ");
    let filtered = [];
    for (let i = 0; i < parts.length; i++) {
      if (stopWords.includes(parts[i].toLowerCase())) break;
      filtered.push(parts[i]);
    }
    const trimmed = name.trim();
    if (trimmed.endsWith("عمده") && filtered.length && filtered[filtered.length - 1] !== ".عمده") {
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

  // ------------------------------------- نسخه جدید و اصلاح‌شده handleFileUpload -------------------------------------
  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
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

      // --- ۲) alias حرفه‌ای برای انواع اسامی ستون‌ها ---
      const columnAliases = {
        "نام تجاری کالا": ["نام تجاری کالا"],
      
        "عرضه کننده": ["عرضه کننده"],
      
       "واحد عرضه": ["واحد عرضه"],
      
        "حجم عرضه": [
          "حجم عرضه",
          "حجم كل معامله شده بر حسب واحد عرضه",
          "حجم کل معامله شده بر حسب واحد عرضه"
        ],
      
        "حجم تقاضا": [
          "حجم تقاضا",
          "تقاضا(بر حسب واحد عرضه)",
          "تقاضا بر حسب واحد عرضه"
        ],
      
        "قیمت مبنا": [
          "قیمت مبنا",
          "قيمت مبنا",
          "قيمت مبنا جزئيات محاسبه قيمت نهايي معاملات کشف پریمیوم",
          "قیمت مبنا جزئیات محاسبه"
        ],
      
        "قیمت پایه": [
          "قیمت پایه",
          "قیمت پایه (ریال)"
        ],
      
        "بیشترین قیمت معامله": [
          "بیشترین قیمت معامله"
        ],
      
        "کمترین قیمت معامله": [
          "کمترین قیمت معامله"
        ],
      
        "متوسط قيمت معامله": [
          "متوسط قيمت معامله",
          "متوسط قیمت معامله"
        ],
      
        "مقصد": ["مقصد"],
      
        "نوع معامله": ["نوع معامله"]
      };
      

      // --- ۳) ساخت مپ نهایی ستون‌ها بر اساس alias ---
      const headerMap = {};

      Object.keys(columnAliases).forEach((mainKey) => {
        const aliases = columnAliases[mainKey];

        for (let alias of aliases) {
          const normAlias = normalizeText(alias);
          const found = Object.keys(realHeaders).find((h) =>
          h === normAlias
        );
        
        
                  if (found) {
            headerMap[mainKey] = realHeaders[found];
            break;
          }
        }
      });

      // --- ۴) خواندن داده‌ها ---
      const jsonData = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;

        const rowData = {};

        Object.keys(columnAliases).forEach((col) => {
          const colIndex = headerMap[col];
          let value = colIndex ? row.getCell(colIndex).value || "" : "";

          if (col === "نام تجاری کالا") value = fixProductName(value);

          rowData[col] = value;
        });

        jsonData.push(rowData);
      });

      // --- ۵) دسته‌بندی داخلی / بین‌الملل ---
      const domestic = [], international = [], other = [];

      jsonData.forEach(row => {
        const dest = normalizeText(row["مقصد"]);
        if (dest.includes("داخلی")) domestic.push(row);
        else if (dest.includes("بین")) international.push(row);
        else other.push(row);
      });

      setDomesticRows(domestic);
      setInternationalRows(international);
      setOtherRows(other);

      // --- ۶) ساخت جدول نمایش صفحه ---
      const emptyRow = Object.fromEntries(columns.map(c => [c, ""]));
      const createHeaderRow = () => {
        const row = {};
        columns.forEach(col => row[col] = col);
        row.__isHeader = true;
        return row;
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

  // -------------------------------------------------- بقیه کد دست‌نخورده --------------------------------------------------
  const exportExcel = async () => {
    if (data.length === 0) return;

    const workbook = new ExcelJS.Workbook();

    const autoFitColumns = (worksheet, columns) => {
      const skipRows = 2;
      columns.forEach((col, index) => {
        let maxLength = 1;
        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
          if (rowNumber <= skipRows) return;
          const cell = row.getCell(index + 1);
          let cellValue = cell.value;
          if (cellValue && typeof cellValue === 'object' && cellValue.richText) {
            cellValue = cellValue.richText.map(t => t.text).join('');
          } else if (cellValue && typeof cellValue !== 'string') {
            cellValue = cellValue.toString();
          }
          if (cellValue) {
            const cellLength = cellValue.length;
            if (cellLength > maxLength) maxLength = cellLength;
          }
        });
        worksheet.getColumn(index + 1).width = maxLength + 2;
      });
    };

    const addMergedHeader = (worksheet, columns, topText, secondText) => {
      const lastColLetter = String.fromCharCode(64 + columns.length);
      worksheet.mergeCells(`A1:${lastColLetter}1`);
      const topCell = worksheet.getCell("A1");
      topCell.value = topText;
      topCell.font = { name: "B Nazanin", bold: true, size: 14 };
      topCell.alignment = { horizontal: "center", vertical: "middle" };
      topCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF6e8bc5" } };
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
      secondCell.font = { name: "B Nazanin", bold: true, size: 18 };
      secondCell.alignment = { horizontal: "center", vertical: "middle" };
      secondCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFB3C6E7" } };
      worksheet.getRow(2).height = 40;
      worksheet.getRow(2).eachCell((cell) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });

      const headerRow = worksheet.addRow(columns.map((col) => col));
      headerRow.eachCell((cell) => {
        cell.font = { name: "B Nazanin", bold: true, size: 12 };
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

    const addMergedHeaderWithColors = (
      worksheet,
      columns,
      topText,
      secondText,
      colors
    ) => {
      const lastColLetter = String.fromCharCode(64 + columns.length);

      worksheet.mergeCells(`A1:${lastColLetter}1`);
      const topCell = worksheet.getCell("A1");
      topCell.value = topText;
      topCell.font = { name: "B Nazanin", bold: true, size: 14 };
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
      secondCell.font = { name: "B Nazanin", bold: true, size: 18 };
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

      const headerRow = worksheet.addRow(columns.map((col) => col));
      headerRow.eachCell((cell) => {
        cell.font = { name: "B Nazanin", bold: true, size: 12 };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.header3 } };
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

    const getTomorrowShamsi = () => {
      const today = new Date();
      const tomorrow = new Date(today);
      tomorrow.setDate(today.getDate() + 1);
      const { jy, jm, jd } = jalaali.toJalaali(tomorrow);
      const dayNames = ["یکشنبه", "دوشنبه", "سه‌شنبه", "چهارشنبه", "شنبه"];
      const dayOfWeek = dayNames[tomorrow.getDay()];
      return {
        day: dayOfWeek,
        date: `${jy}/${jm.toString().padStart(2, "0")}/${jd
          .toString()
          .padStart(2, "0")}`,
      };
    };

    const { day, date } = getTomorrowShamsi();

    // =================== شیت داخلی ===================
    if (domesticRows.length > 0) {
      const internalColumns = [
        "نام تجاری کالا",
        "عرضه کننده",
        "واحد عرضه",
        "حجم عرضه",
        "حجم تقاضا",
        "قیمت مبنا",
        "قیمت پایه",
        "بیشترین قیمت معامله",
        "کمترین قیمت معامله",
        "متوسط قيمت معامله",
        "مقصد",
      ];

      const domesticSheet = workbook.addWorksheet("داخلی", {
        properties: { defaultRowHeight: 25 },
        views: [{ rightToLeft: true }],
      });

      const displayInternalColumns = internalColumns.map((col) => {
        if (col === "حجم كل معامله شده بر حسب واحد عرضه") return "حجم عرضه";
        if (col === "تقاضا(بر حسب واحد عرضه)") return "حجم تقاضا";
        if (col === "قيمت مبنا جزئيات محاسبه قيمت نهايي معاملات کشف پریمیوم")
          return "قيمت مبنا";
        return col;
      });

      const headerRowNumber = addMergedHeader(
        domesticSheet,
        displayInternalColumns,
        "کارگزاری آینده نگر خوارزمی به مدیریت دکتر ذوقی  09123011311 و 61914000-021 (داخلی 401)",
        `عرضه محصولات بورس انرژی در رینگ داخلی روز ${day} مورخ ${date}`
      );

      const wmBase64 = createWatermark(
        "کارگزاری آینده نگر خوارزمی",
        "09123011311"
      );
      const wmId = workbook.addImage({
        base64: wmBase64,
        extension: "png",
      });

      domesticSheet.addImage(wmId, {
        tl: { col: 2, row: 2 },
        ext: { width: 600, height: 300 },
        editAs: "oneCell",
      });

      domesticRows.forEach((rowData, i) => {
        const row = domesticSheet.addRow(
          internalColumns.map((col) => rowData[col] || "")
        );
        const fillColor = i % 2 === 0 ? "FFFFFFFF" : "FFF5F5F5";

        row.eachCell((cell) => {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fillColor } };
          cell.font = { name: "B Nazanin" };
          cell.alignment = { horizontal: "center", vertical: "middle" };
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        });
      });

      domesticSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber > headerRowNumber) row.height = 25;
      });

      autoFitColumns(domesticSheet, internalColumns);
    }

    // =================== شیت بین الملل ===================
    let mainColumns = [...columns];
    const kashfPremiumRows = [...internationalRows, ...otherRows].filter(
      (row) => row["نوع معامله"]?.includes("کشف پریمیوم")
    );
    const otherNonPremiumRows = [...internationalRows, ...otherRows].filter(
      (row) => !row["نوع معامله"]?.includes("کشف پریمیوم")
    );

    const mainDataRows = [...kashfPremiumRows, ...otherNonPremiumRows];

    const basePriceKey =
      "قيمت مبنا جزئيات محاسبه قيمت نهايي معاملات کشف پریمیوم";
    const priceColIndex = mainColumns.findIndex((c) => c === basePriceKey);

    if (priceColIndex !== -1) {
      const allEmpty = mainDataRows.every(
        (row) =>
          !row[basePriceKey] || row[basePriceKey].toString().trim() === ""
      );
      if (allEmpty) mainColumns.splice(priceColIndex, 1);
    }

    const mainSheet = workbook.addWorksheet("بین الملل", {
      properties: { defaultRowHeight: 25 },
      views: [{ rightToLeft: true }],
    });

    const displayColumnsMap = (col) => {
      if (col === "حجم كل معامله شده بر حسب واحد عرضه") return "حجم عرضه";
      if (col === "تقاضا(بر حسب واحد عرضه)") return "حجم تقاضا";
      if (col === "قيمت مبنا جزئيات محاسبه قيمت نهايي معاملات کشف پریمیوم")
        return "قيمت مبنا";
      return col;
    };

    const headerRowNumber = addMergedHeaderWithColors(
      mainSheet,
      mainColumns.map(displayColumnsMap),
      "کارگزاری آینده نگر خوارزمی به مدیریت دکتر ذوقی 09123011311 و 61914000-021 (داخلی 401)",
      `عرضه محصولات بورس انرژی در رینگ صادراتی روز ${day} مورخ ${date}`,
      { header1: "FFCC2B52", header2: "FFD9A299", header3: "FFCAC8C9" }
    );

    const wmBase64 = createWatermark(
      "کارگزاری آینده نگر خوارزمی",
      "09123011311"
    );
    const wmId = workbook.addImage({
      base64: wmBase64,
      extension: "png",
    });

    mainSheet.addImage(wmId, {
      tl: { col: 2, row: 2 },
      ext: { width: 600, height: 300 },
      editAs: "oneCell",
    });

    mainDataRows.forEach((rowData, i) => {
      const rowDataCopy = { ...rowData };
      if (rowDataCopy["قیمت پایه"])
        rowDataCopy["قیمت پایه"] = formatPriceBase(rowDataCopy["قیمت پایه"]);

      const row = mainSheet.addRow(
        mainColumns.map((col) => rowDataCopy[col] || "")
      );

      const fillColor = i % 2 === 0 ? "FFFFFFFF" : "FFFAF7F3";

      row.eachCell((cell) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fillColor } };
        cell.font = { name: "B Nazanin" };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
    });

    mainSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (rowNumber > headerRowNumber) row.height = 25;
    });

    autoFitColumns(mainSheet, mainColumns);

    const buf = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buf]), "Enerzhi_Output_Separated.xlsx");
  };

  // ------------------------------------- UI ------------------------------------------
  return (
    <div style={{ padding: "20px", fontFamily: "Arial" }}>
      <h2>{titleText}</h2>
      <h4>{SuplyText}</h4>

      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      <button onClick={exportExcel} style={{ marginLeft: "10px" }}>
        خروجی اکسل
      </button>

      <table
        style={{ width: "100%", marginTop: "20px", borderCollapse: "collapse" }}
      >
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
