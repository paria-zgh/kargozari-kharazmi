import React, { useState } from "react";
import ExcelJS from "exceljs";
import jalaali from "jalaali-js";

export default function Smain() {
  const [data, setData] = useState([]);

  // ----- helper: normalize any ExcelJS cell value to a trimmed string -----
  const normalizeCellValue = (val) => {
    if (val == null) return "";
    // ExcelJS richText or { text: "..."} or array etc.
    if (typeof val === "object") {
      // richText => { richText: [{text: '...'}, ...] }
      if (Array.isArray(val.richText)) {
        return val.richText.map((t) => t.text || "").join("").trim();
      }
      // {text: "..." }
      if (val.text) return String(val.text).trim();
      // some other object, try to stringify safely
      if (typeof val.toString === "function") return String(val.toString()).trim();
      return "";
    }
    return String(val).trim();
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      const arrayBuffer = evt.target?.result;
      if (!arrayBuffer) return;

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);
      const worksheet = workbook.worksheets[0];

      const jsonData = [];
      // read headers from row 1 and normalize
      const rawHeaderValues = worksheet.getRow(1).values.slice(1);
      const columns = rawHeaderValues.map((v) => normalizeCellValue(v));

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const obj = {};
        columns.forEach((col, idx) => {
          const cellValue = row.getCell(idx + 1).value;
          obj[col] = normalizeCellValue(cellValue);
        });
        jsonData.push(obj);
      });

      setData(jsonData);
    };
    reader.readAsArrayBuffer(file);
  };

  const applyAllBorders = (cell) => {
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  };
// --- helper: create watermark image as base64 ---
const createWatermark = (text1, text2) => {
  const canvas = document.createElement("canvas");
  canvas.width = 3600; // کمی پهن‌تر برای جا شدن کامل متن
  canvas.height = 1600; // کمی بلندتر برای فاصله خطوط
  
  const ctx = canvas.getContext("2d");

  ctx.clearRect(0, 0, canvas.width, canvas.height);

  // انتقال به وسط + کمی به راست تا حرف اول حذف نشه
  ctx.translate(canvas.width / 2 , canvas.height / 2);
  ctx.rotate(-8 * Math.PI / 180);

  ctx.textAlign = "center";
  ctx.fillStyle = "rgba(0,70,180,0.25)"; // آبی شفاف همان استایل قبلی

  // سایه نرم برای برجسته شدن
  ctx.shadowColor = "rgba(0,0,0,0.25)";
  ctx.shadowBlur = 15;
  ctx.shadowOffsetX = 6;
  ctx.shadowOffsetY = 6;

  // === خط اول ===
  ctx.font = "bold 250px Arial";
  ctx.fillText(text1, 0, -180); // فاصله از مرکز بیشتر تا جدا شود

  // === خط دوم ===
  ctx.font = "bold 230px Arial";
  ctx.fillText(text2, 0, 260); // فاصله زیادتر تا در هم نرود

  return canvas.toDataURL("image/png");
};

  const processData = async () => {
    if (data.length === 0) {
      alert("ابتدا یک فایل آپلود کنید.");
      return;
    }

    // Sort and block logic unchanged but data is already normalized strings
    const sorted = [...data].sort((a, b) =>
      String(a["تالار"] || "").localeCompare(String(b["تالار"] || ""), "fa")
    );

    const keywordsSub = [
      "مس","فولاد","ضایعات","پالت چوبی","بشکه خالی","ورق","MDF","M.D.F",
      "پشم شیشه","کنسانتره","اقلام تجهیزات","اقلام تکمیلی خودرو",
      "شمش","سپری","تختال","بتنی","گرانول نقره","فلز","بیلت","چدن","لوله"
    ];
    const petrochemKeywords = ["پلی","الیاف استیپل اکریلیک","تیشو","پی وی سی"];
    const petroleumKeywords = ["قیر", "نفت", "واکس", "گوگرد", "اسلاک" ,"پارافین","وکیوم","روغن","لوب","PDA"];

    const blocks = {
      keywordBlockSub: [],
      keywordBlockPetroleumFromSub: [],
      petroleumFromPetrochem: [], // <-- این خط اضافه شود

      auctionBlockStoneOrCathode: [],
      auctionBlockVacuum: [],
      petrochemBlock: [],
      otherRows: [],
    };

      sorted.forEach((row) => {
      const curTalar = row["تالار"] || "";
      const namaKala = row["نام کالا"] || "";
    
      // 👈 بررسی تالار صنعتی و معدنی برای تیشو
      if (curTalar === "تالار صنعتی و معدنی") {
        if (namaKala.includes("تیشو")) {
          blocks.petrochemBlock.push(row); // انتقال تیشو از صنعتی به پتروشیمی/پلیمری
        } else {
          blocks.otherRows.push(row);
        }
      }
      else if (curTalar === "تالار فرعی") {
        if (["نفتی", "نفت", "وکیوم", "قیر", "روغن","PDA"].some((kw) => namaKala.includes(kw))) {
          blocks.keywordBlockPetroleumFromSub.push(row);
        } else if (petrochemKeywords.some((kw) => namaKala.includes(kw))) {
          blocks.petrochemBlock.push(row);
        } else if (keywordsSub.some((kw) => namaKala.includes(kw))) {
          blocks.keywordBlockSub.push(row);
        } else {
          blocks.petrochemBlock.push(row);
        }
      } 
      else if (curTalar === "تالار حراج باز") {
        if (
          namaKala.includes("سنگ") ||
          namaKala.includes("مس کاتد") ||
          namaKala.includes("فلز") ||
          namaKala.includes("تختال") ||
          namaKala.includes("اکسید مولیبدن")
        ) {
          blocks.auctionBlockStoneOrCathode.push(row);
        } else if (namaKala.includes("وکیوم")) {
          blocks.auctionBlockVacuum.push(row);
        } else if (petrochemKeywords.some((kw) => namaKala.includes(kw))) {
          blocks.petrochemBlock.push(row);
        } else {
          blocks.petrochemBlock.push(row);
        }
      } 
      else if (curTalar === "تالار پتروشیمی و فرآورده های نفتی") {
        if (petroleumKeywords.some((kw) => namaKala.includes(kw))) {
          blocks.petroleumFromPetrochem.push(row);
        } else {
          blocks.otherRows.push(row);
        }
      } 
      else {
        blocks.otherRows.push(row);
      }
    });

    // ۱. تغییر تالار تمام اقلام فرعی و حراج باز پتروشیمی به تالار اصلی
    const subPetrochemRows = blocks.petrochemBlock.map((r) => ({
      ...r,
      تالار: "تالار پتروشیمی و فرآورده های نفتی",
    }));

    const finalRows = [...blocks.otherRows];

    // ۲. تزریق این ردیف‌ها به داخل بلوک تالار پتروشیمی در کنار بقیه اقلام
    if (subPetrochemRows.length > 0) {
      let insertIdx = -1;
      finalRows.forEach((r, i) => {
        if ((r["تالار"] || "") === "تالار پتروشیمی و فرآورده های نفتی") {
          insertIdx = i;
        }
      });

      if (insertIdx !== -1) {
        finalRows.splice(insertIdx + 1, 0, ...subPetrochemRows);
      } else {
        finalRows.unshift(...subPetrochemRows);
      }
    }

    // ردیف‌های صنعتیِ آمده از تالار فرعی
    const subIndustrialRows = blocks.keywordBlockSub.map((r) => ({
      ...r,
      تالار: "تالار صنعتی و معدنی",
    }));

    // ردیف‌های صنعتیِ آمده از حراج باز (سنگ/کاتد/...)
    const auctionIndustrialRows = blocks.auctionBlockStoneOrCathode.map((r) => ({
      ...r,
      تالار: "تالار صنعتی و معدنی",
    }));

    const allIndustrialRows = [...subIndustrialRows, ...auctionIndustrialRows];

    if (allIndustrialRows.length > 0) {
      // اگر خود تالار صنعتی در داده هست: بعد از آخرین صنعتی درج کن
      const lastIndustrialIdx = finalRows
        .map((r) => r["تالار"] || "")
        .lastIndexOf("تالار صنعتی و معدنی");

      if (lastIndustrialIdx !== -1) {
        finalRows.splice(lastIndustrialIdx + 1, 0, ...allIndustrialRows);
      } else {
        // اگر صنعتی وجود ندارد: قبل از اولین تالار مهم قرار بده تا بلوک‌بندی/شیت‌سازی به هم نریزد
        const fallbackIdx = finalRows.findIndex((r) =>
          ["تالار سیمان", "تالار کالای صادراتی کيش", "تالار فرآورده های نفتی"]
            .includes(r["تالار"] || "")
        );

        if (fallbackIdx !== -1) finalRows.splice(fallbackIdx, 0, ...allIndustrialRows);
        else finalRows.unshift(...allIndustrialRows);
      }
    }
      
    // ۴. ساخت بلوک مستقل فرآورده‌های نفتی (شامل نفتی‌های تالار فرعی)
    const petroleumRows = [
      ...blocks.petroleumFromPetrochem,
      ...blocks.keywordBlockPetroleumFromSub,
      ...blocks.auctionBlockVacuum
    ].map((row) => ({
      ...row,
      تالار: "تالار فرآورده های نفتی",
    }));
    
    if (petroleumRows.length > 0) {
      finalRows.push(...petroleumRows);
    }
  
    const insertTalarNames = [
      "تالار صنعتی و معدنی",
      "تالار فرآورده های نفتی",
      "تالار سیمان",
      "تالار کالای صادراتی کيش"
    ];
    
    let processed = [];
    const inserted = new Set();
    
    finalRows.forEach((row) => {
      const curTalar = row["تالار"] || "";
    
      if (
        insertTalarNames.includes(curTalar) &&
        !inserted.has(curTalar)
      ) {
        processed.push({});
        inserted.add(curTalar);
      }
    
      processed.push(row);
    });
    
    const specialKeywords=["بوتادین استایرن","استایرن بوتادین","تیشو","پلی","الیاف استیپل اکریلیک","ABS","پی وی سی"];
    const firstInsertIndex=processed.findIndex(row=>Object.keys(row).length===0);
    if(firstInsertIndex>0){
      const beforeInsert=processed.slice(0,firstInsertIndex);
      const afterInsert=processed.slice(firstInsertIndex);
      const specialRows=[], normalRows=[];
      beforeInsert.forEach(row=>{
        const namaKala=row["نام کالا"]||"";
        if (specialKeywords.some(kw => namaKala.includes(kw)) && !namaKala.includes("پلیمریک"))  specialRows.push(row);
        else normalRows.push(row);
      });
      if(specialRows.length>0) processed=[...normalRows,{},...specialRows,...afterInsert];
    }

    const blocksSorted=[];
    let currentBlock=[];
    processed.forEach(row=>{
      const isInsert=Object.keys(row).length===0;
      if(isInsert){
        if(currentBlock.length>0){
          currentBlock.sort((a,b)=>{
            if((a["تالار"]||"")==="تالار سیمان"&&(b["تالار"]||"")==="تالار سیمان"){
              return String(a["تولید کننده کالا"]||"").localeCompare(String(b["تولید کننده کالا"]||""),"fa");
            }
            return String(a["نام کالا"]||"").localeCompare(String(b["نام کالا"]||""),"fa");
          });
          blocksSorted.push({__HEADER__:true});
          blocksSorted.push(...currentBlock);
          currentBlock=[];
        }
        // سطر خالی جداکننده اینجا نباید به صورت داده به blocksSorted برود
      } else currentBlock.push(row);
    });
    if(currentBlock.length>0){
      currentBlock.sort((a,b)=>{
        if((a["تالار"]||"")==="تالار سیمان"&&(b["تالار"]||"")==="تالار سیمان"){
          return String(a["تولید کننده کالا"]||"").localeCompare(String(b["تولید کننده کالا"]||""),"fa");
        }
        return String(a["نام کالا"]||"").localeCompare(String(b["نام کالا"]||""),"fa");
      });
      blocksSorted.push({__HEADER__:true});
      blocksSorted.push(...currentBlock);
    }

    let headersArr=Array.from(new Set(blocksSorted.flatMap(r=>Object.keys(r)))).filter(
      h=>h!=="تاریخ عرضه" && h!=="__HEADER__" && h!=="تعداد محموله" && h!=="قیمت پایه"
    );
    headersArr=headersArr.map(h=>{
      if(h==="حجم") return "مقدار پایه";
      if(h==="قیمت") return "قیمت پایه";
      if(h==="کد") return "کد عرضه";
      if(h==="تسویه") return "نوع تسویه";
      if(h==="حداکثر افزایش حجم سفارش") return "حداکثر افزایش عرضه";
      return h;
    });

    const producer="تولید کننده کالا", delivery="محل تحویل";
    if(headersArr.includes(producer) && headersArr.includes(delivery)){
      const withoutProducer=headersArr.filter(h=>h!==producer);
      const idxDelivery=withoutProducer.indexOf(delivery);
      if(idxDelivery!==-1) withoutProducer.splice(idxDelivery,0,producer);
      headersArr=withoutProducer;
    }

    const reordered=blocksSorted.map(row=>{
      if(row.__HEADER__) return row;
      const nr={};
      headersArr.forEach(h=>{
        let originalKey=
          h==="مقدار پایه"?"حجم":
          h==="قیمت پایه"?"قیمت":
          h==="کد عرضه"?"کد":
          h==="نوع تسویه"?"تسویه":
          h==="حداکثر افزایش عرضه"?"حداکثر افزایش حجم سفارش":
          h;
        let value=row[originalKey]??"";
        if(h==="مقدار پایه"){
          const vol=Number(String(value).replace(/,/g,'')); // حذف کاما
          nr[h]=isNaN(vol) || value==="" ? null : vol/1000; 
                
        }else if(h==="قیمت پایه"){
          const price=Number(String(value).replace(/,/g,''));
          nr[h]=isNaN(price) || value==="" ? null : price; 
          
        } else nr[h]=value;
      });
      
      return nr;
    });

    // --- بلوک‌های ویژه انتهایی ---
    const addSpecialBlock=(rows)=> {
      if(rows.length>0){
        reordered.push({__HEADER__:true});
        rows.forEach(r=>reordered.push({...r}));
      }
    };

    // حفظ شرط‌های خاص شما
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("بطری")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی پروپیلن") && r["نام کالا"]?.includes("نساجی") && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی پروپیلن") && r["نام کالا"]?.includes("نساجی") && r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی پروپیلن") && r["نام کالا"]?.includes("شیمیایی") && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی پروپیلن") && r["نام کالا"]?.includes("شیمیایی") && r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r => 
      (r["نام کالا"]?.includes("پی وی سی") || r["نام کالا"]?.includes("پلی وینیل کلراید")) 
      && r["نام کالا"]?.includes("S")
    ));
    
    addSpecialBlock(reordered.filter(r => 
      (r["نام کالا"]?.includes("پی وی سی") || r["نام کالا"]?.includes("پلی وینیل کلراید")) 
      && r["نام کالا"]?.includes("E")
    ));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("استایرن"))) ;
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("منو اتیلن گلایکول") || r["نام کالا"]?.includes("دی اتیلن گلایکول")||r["نام کالا"]?.includes("تری اتیلن گلایکول")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("اسید ترفتالیک"))) ;
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی اتیلن سبک")&& r["نام کالا"]?.includes("تزریقی") && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی اتیلن سبک") && r["نام کالا"]?.includes("خطی") && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی اتیلن سبک") && r["نام کالا"]?.includes("فیلم") && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی اتیلن سبک") && r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی اتیلن سنگین") &&
      (r["نام کالا"]?.includes("اکستروژن") || r["نام کالا"]?.includes("لوله")) && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("سنگین") &&
      (r["نام کالا"]?.includes("PEWAX") || r["نام کالا"]?.includes("کلوخه")|| r["نام کالا"]?.includes("پلی اتیلن سنگین پودر")) && !r["نام کالا"]?.toLowerCase().includes("off")));
    ["بادی","تزریقی","فیلم","دورانی"].forEach(k=>{
      addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی اتیلن سنگین") &&
        r["نام کالا"]?.includes(k) && !r["نام کالا"]?.toLowerCase().includes("off")));
    });
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("سنگین") &&
      r["نام کالا"]?.toLowerCase().includes("off")));

    // بلوک ویژه تولوئن / متیلن
    addSpecialBlock(reordered.filter(r => 
      r["نام کالا"]?.includes("تولوئن دی ایزو سیانات") || 
      r["نام کالا"]?.includes("متیلن دی فنیل دی ایزوسیانات خالص") ||
      r["نام کالا"]?.includes("متیلن دی فنیل ایزوسیانات خالص")
    ));

    // --- تولید اکسل با چند شیت ---
    const workbook = new ExcelJS.Workbook();
    workbook.views = [{ rightToLeft: true }];

    const titleText="کارگزاری آینده نگر خوارزمی به مدیریت دکتر ذوقی 09123011311-02161914000 داخلی(408 الی 412) ";
    const today = new Date();
    let nextDay = new Date(today);
    let diffDays = today.getDay()===3?3:1;
    nextDay.setDate(today.getDate()+diffDays);
    const j=jalaali.toJalaali(nextDay);
    const jalaliDate=`${j.jy}/${String(j.jm).padStart(2,"0")}/${String(j.jd).padStart(2,"0")}`;
    const daysFa=["یکشنبه","دوشنبه","سه‌شنبه","چهارشنبه","پنجشنبه","جمعه","شنبه"];
    const dayFa=daysFa[nextDay.getDay()];
    const supplyText=`عرضه X روز ${dayFa} مورخ ${jalaliDate}`;

    const styleRow=(row,bgColor,bold=true)=>{
      row.eachCell(cell=>{
        if(bgColor) cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:bgColor}};
        cell.font={name:"Arial", bold, size: 14};
        cell.alignment={vertical:"middle",horizontal:"center"};
        applyAllBorders(cell);
      });
    };
    
    const addTitleRows = (ws) => {
      const titleRow = ws.addRow([titleText]);
      titleRow.height = 33;
      ws.mergeCells(titleRow.number, 1, titleRow.number, headersArr.length);
    
      titleRow.eachCell((cell) => {
        cell.font = { name: "Arial", bold: true, size: 12 };
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFA500" } };
        applyAllBorders(cell);
      });
    
      const supplyRow = ws.addRow([supplyText]);
      supplyRow.height = 40;
      ws.mergeCells(supplyRow.number,1,supplyRow.number,headersArr.length);
    
      supplyRow.eachCell((cell) => {
        cell.font = { name: "Arial", bold: true, size: 18 };
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
        applyAllBorders(cell);
      });
    
      const headerRow = ws.addRow(headersArr);
      headerRow.height = 31;
      headerRow.eachCell((cell) => {
        cell.font = { name: "Arial", bold: true,size:14 };
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD3D3D3" } };
        applyAllBorders(cell);
      });
    };
    
    let sheetCounter = 1;
    let currentSheet = null;
    let rowCounter = 0;

    // تابع کمکی برای تشخیص ردیف دارای داده واقعی (جلوگیری از ساخت ردیف‌های پوچ با مقدار 0)
    const hasValidData = (row) => {
      if (!row || row.__HEADER__) return false;
      return Object.values(row).some(v => v !== null && v !== "" && v !== undefined);
    };
    
    reordered.forEach((row, index) => {
      if (row.__HEADER__) {
        const nextHeaderIndex = reordered.slice(index + 1).findIndex(r => r.__HEADER__);
        const dataForSheet = nextHeaderIndex === -1 
          ? reordered.slice(index + 1)
          : reordered.slice(index + 1, index + 1 + nextHeaderIndex);
    
        // فقط اگر داده واقعی معتبر هست شیت بساز
        if(dataForSheet.some(r => hasValidData(r))){
          sheetCounter++;
          currentSheet = workbook.addWorksheet(`Sheet ${sheetCounter}`, {views:[{rightToLeft:true}]});
          addTitleRows(currentSheet);
          rowCounter = 0;
        }
      } else if(hasValidData(row)){
        // اضافه کردن فقط ردیف‌های معتبر
        if(!currentSheet){
          sheetCounter = 1;
          currentSheet = workbook.addWorksheet(`Sheet ${sheetCounter}`, {views:[{rightToLeft:true}]});
          addTitleRows(currentSheet);
          rowCounter = 0;
        }
    
        const dataRow = currentSheet.addRow(
          headersArr.map(h => {
            const val = row[h];
            if (val !== null && val !== "" && !isNaN(val)) return Number(val);
            return val ?? "";
          })
        );
        styleRow(dataRow, null, true);
        dataRow.height = 30;
    
        if(rowCounter % 2 === 0){
          dataRow.eachCell(cell => {
            cell.fill = {type:"pattern", pattern:"solid", fgColor:{argb:"FFF5DEB3"}};
          });
        }
    
        rowCounter++;
    
        // تنظیم فرمت عددی
        const colMiqdar = headersArr.indexOf("مقدار پایه") + 1;
        const colGheymat = headersArr.indexOf("قیمت پایه") + 1;
        if (colMiqdar > 0) dataRow.getCell(colMiqdar).numFmt = "#,##0";
        if (colGheymat > 0) dataRow.getCell(colGheymat).numFmt = "#,##0";
      }
    });
    
    // --- اصلاح هدر شیت‌ها بر اساس ستون "تالار" و نام کالا ---
    const normalizeExcelCell = (cellVal) => {
      if (cellVal == null) return "";
      if (typeof cellVal === "object") {
        if (cellVal.text) return String(cellVal.text).trim();
        if (Array.isArray(cellVal.richText)) return cellVal.richText.map(t => t.text || "").join("").trim();
        return String(cellVal.toString()).trim();
      }
      return String(cellVal).trim();
    };
    
    workbook.eachSheet((ws) => {
      const petrochemKeywords = ["بوتادین استایرن","استایرن بوتادین","تیشو","پلی","الیاف استیپل اکریلیک"];
      const colNamaKala = headersArr.indexOf("نام کالا") + 1;
      const colTalar = headersArr.indexOf("تالار") + 1;
    
      let hasIndustrial = false;
      let hasSeman = false;
      let hasPetroleum = false;
      let hasExport = false;
      let hasPetrochemBlock = false;
      let notPetrochemBlock = false;

      let allPPTextile = true;
      let allPPTextileOff = true;
      let allPPChemical = true;
      let allPPChemicalOff = true;
      let allPVC_S = true;
      let allPVC_E = true;
      let allDEG = true;
      let allMEG = true;
      let allDEGandMEG = true;
      let allPET_Bottle = true;
      let allPolystyrene = true;
      let asid = true;

      // فلگ شناسایی ایزوسیانات برای نام‌گذاری کارون
      let allIsocyanate = true;
    
      // فلگ‌ها برای پلی اتیلن سبک
      const specialPEBlocks = [
        { keyword: "پلی اتیلن سبک خطی", headerText: "پلی اتیلن سبک خطی", allPresent: true, allOff: false },
        { keyword: "پلی اتیلن سبک فیلم", headerText: "پلی اتیلن سبک فیلم", allPresent: true, allOff: false },
        { keyword: "پلی اتیلن سبک تزریقی", headerText: "پلی اتیلن سبک تزریقی", allPresent: true, allOff: false },
        { keyword: "پلی اتیلن سبک", headerText: "پلی اتیلن سبک", allPresent: true, allOff: true, offHeader: "های OFF پلی اتیلن سبک" },
      ];
    
      // فلگ‌ها برای پلی اتیلن سنگین
      const heavyPEBlocks = [
        { keywords: ["پلی اتیلن سنگین اکستروژن","پلی اتیلن سنگین لوله"], headerText: "پلی اتیلن سنگین لوله و اکستروژن", allPresent: true },
        { keywords: ["پلی اتیلن سنگین کلوخه","پلی اتیلن سنگین پودر","پلی اتیلن سنگین PEWAX"], headerText: "پلی اتیلن سنگین", allPresent: true },
        { keywords: ["پلی اتیلن سنگین دورانی"], headerText: "پلی اتیلن سنگین دورانی", allPresent: true },
        { keywords: ["پلی اتیلن سنگین بادی"], headerText: "پلی اتیلن سنگین بادی", allPresent: true },
        { keywords: ["پلی اتیلن سنگین فیلم"], headerText: "پلی اتیلن سنگین فیلم", allPresent: true },
        { keywords: ["پلی اتیلن سنگین تزریقی"], headerText: "پلی اتیلن سنگین تزریقی", allPresent: true },
        { keywords: ["پلی اتیلن سنگین"], headerText: "های OFF پلی اتیلن سنگین", allPresent: true, checkOff: true }
      ];
    
      // iterate rows
      ws.eachRow((row, rowNumber) => {
        if (rowNumber <= 3) return; 
        const namaKalaRaw = colNamaKala > 0 ? normalizeExcelCell(row.getCell(colNamaKala).value) : "";
        const nkLower = namaKalaRaw.toLowerCase().replace(/[\s\u200C]+/g, "");
        const containsOff = nkLower.includes("off");
    
        const rawTalar = colTalar > 0 ? normalizeExcelCell(row.getCell(colTalar).value) : "";
        if (rawTalar === "تالار صنعتی و معدنی") hasIndustrial = true;
        if (rawTalar === "تالار سیمان") hasSeman = true;
        if (rawTalar === "تالار فرآورده های نفتی") hasPetroleum = true;
        if (rawTalar === "تالار کالای صادراتی کيش") hasExport = true;
        if (rawTalar ==="تالار پتروشیمی و فرآورده های نفتی" && petrochemKeywords.some(kw => namaKalaRaw.includes(kw))) {
          hasPetrochemBlock = true;
        }
        if (
          rawTalar === "تالار پتروشیمی و فرآورده های نفتی" ||
          rawTalar === "تالار حراج باز" ||
          (
            rawTalar === "تالار فرعی" &&
            !petrochemKeywords.some((kw) => namaKalaRaw.includes(kw))
          )
        ) {
          notPetrochemBlock = true;
        }

        // فلگ ایزوسیانات‌های کارون
        const isIsocyanate =
          namaKalaRaw.includes("تولوئن دی ایزو سیانات") ||
          namaKalaRaw.includes("متیلن دی فنیل دی ایزوسیانات خالص") ||
          namaKalaRaw.includes("متیلن دی فنیل ایزوسیانات خالص");
        allIsocyanate = allIsocyanate && isIsocyanate;
    
        // پلی اتیلن سبک
        specialPEBlocks.forEach(b => {
          const keywordNorm = b.keyword.toLowerCase().replace(/[\s\u200C]+/g, "");
          const containsKeyword = nkLower.includes(keywordNorm);
          b.allPresent = b.allPresent && containsKeyword;
          if (b.offHeader) b.allOff = b.allOff && (containsKeyword && containsOff);
        });
    
        // پلی اتیلن سنگین
        heavyPEBlocks.forEach(b => {
          const keywordsNorm = b.keywords.map(k => k.toLowerCase().replace(/[\s\u200C]+/g, ""));
          const match = keywordsNorm.some(k => nkLower.includes(k));
          if (b.checkOff) {
            b.allPresent = b.allPresent && (nkLower.includes("پلیاتیلنسنگین".replace(/[\s\u200C]+/g, "")) && containsOff);
          } else if (b.keywords.length === 3) {
            b.allPresent = b.allPresent && match && !containsOff;
          } else {
            b.allPresent = b.allPresent && match && !containsOff;
          }
        });
    
        // سایر محصولات
        const containsPPTextile = namaKalaRaw.includes("پلی پروپیلن نساجی");
        const containsPPChemical = namaKalaRaw.includes("پلی پروپیلن شیمیایی");
        allPPTextile = allPPTextile && containsPPTextile;
        allPPTextileOff = allPPTextileOff && (containsPPTextile && containsOff);
        allPPChemical = allPPChemical && containsPPChemical;
        allPPChemicalOff = allPPChemicalOff && (containsPPChemical && containsOff);
        asid = asid && namaKalaRaw.includes("اسید ترفتالیک");
        
        allPVC_S = allPVC_S && ((namaKalaRaw.includes("پلی وینیل کلراید") || namaKalaRaw.includes("پی وی سی")) && namaKalaRaw.includes("S"));
        allPVC_E = allPVC_E && ((namaKalaRaw.includes("پلی وینیل کلراید") || namaKalaRaw.includes("پی وی سی")) && namaKalaRaw.includes("E"));

        allDEG = allDEG && namaKalaRaw.includes("دی اتیلن گلایکول");
        allMEG = allMEG && namaKalaRaw.includes("منو اتیلن گلایکول");
        allDEGandMEG = allDEGandMEG && (namaKalaRaw.includes("دی اتیلن گلایکول") || namaKalaRaw.includes("منو اتیلن گلایکول")||namaKalaRaw.includes("تری اتیلن گلایکول"));
        allPET_Bottle = allPET_Bottle && namaKalaRaw.includes("پلی اتیلن ترفتالات بطری");
        allPolystyrene = allPolystyrene && namaKalaRaw.includes("پلی استایرن");
      });
    
      // جایگزینی X در هدر
      const headerRow = ws.getRow(2);
      headerRow.eachCell((cell) => {
        if (typeof cell.value === "string" && cell.value.includes("X")) {
          // اول پلی اتیلن سنگین
          for (let b of heavyPEBlocks) {
            if (b.allPresent) {
              cell.value = cell.value.replace("X", b.headerText);
              return;
            }
          }
    
          // سپس پلی اتیلن سبک
          for (let b of specialPEBlocks) {
            if (b.offHeader && b.allPresent && b.allOff) {
              cell.value = cell.value.replace("X", b.offHeader);
              return;
            } else if (b.allPresent) {
              cell.value = cell.value.replace("X", b.headerText);
              return;
            }
          }

          // سپس ایزوسیانات‌های کارون
          if (allIsocyanate) {
            cell.value = cell.value.replace("X", "عرضه محصولات پتروشیمی کارون");
            return;
          }
    
          // سایر محصولات قبلی
          if (allPPTextileOff) cell.value = cell.value.replace("X", "های OFF پلی پروپیلن نساجی");
          else if (allPPTextile) cell.value = cell.value.replace("X", "پلی پروپیلن نساجی");
          else if (allPPChemicalOff) cell.value = cell.value.replace("X", "های OFF پلی پروپیلن شیمیایی");
          else if (allPPChemical) cell.value = cell.value.replace("X", "پلی پروپیلن شیمیایی");
          else if (allPVC_S) cell.value = cell.value.replace("X", "پلی وینیل کلراید PVC(S)");
          else if (allPVC_E) cell.value = cell.value.replace("X", "پلی وینیل کلراید PVC(E)");
          else if (allDEGandMEG) cell.value = cell.value.replace("X", "اتیلن گلایکول ها");
          else if (allDEG) cell.value = cell.value.replace("X", "DEG");
          else if (allMEG) cell.value = cell.value.replace("X", "MEG");
          else if (allPET_Bottle) cell.value = cell.value.replace("X", "پلی اتیلن ترفتالات بطری");
          else if (allPolystyrene) cell.value = cell.value.replace("X", "پلی استایرن");
          else if (asid) cell.value = cell.value.replace("X", "اسید ترفتالیک");
          else if (hasIndustrial) cell.value = cell.value.replace("X", "محصولات صنعتی");
          else if (hasSeman) cell.value = cell.value.replace("X", "سیمان");
          else if (hasPetroleum) cell.value = cell.value.replace("X", "فرآورده های نفتی و قیر");
          else if (hasExport) cell.value = cell.value.replace("X", "تالار صادراتی");
          else if (hasPetrochemBlock) cell.value = cell.value.replace("X", "محصولات پلیمری");
          else if (notPetrochemBlock) cell.value = cell.value.replace("X", "محصولات شیمیایی");
        }
      });
    });

    // === افزودن جمع کل عرضه‌ها و اختلاف به پایین شیت‌های خاص و سیمان ===
    workbook.eachSheet((ws) => {
      const volumeColIndex = headersArr.indexOf("مقدار پایه") + 1;
      const increaseColIndex = headersArr.indexOf("حداکثر افزایش عرضه") + 1;

      if (volumeColIndex <= 0 || increaseColIndex <= 0) {
        return;
      }

      const colNamaKala = headersArr.indexOf("نام کالا") + 1;
      const colTalar = headersArr.indexOf("تالار") + 1;

      let isTargetSheet = false;

      let isSeman = false;
      let isPET_Bottle = true;
      let isPPTextile = true;
      let isPPTextileOff = true;
      let isPPChemical = true;
      let isPPChemicalOff = true;
      let isPVC_S = true;
      let isPVC_E = true;
      let isPolystyrene = true;
      let isDEGandMEG = true;
      let isAsid = true;

      let isPE_Light_Injection = true;
      let isPE_Light_Linear = true;
      let isPE_Light_Film = true;
      let isPE_Light_Off = true;

      let isPE_Heavy_Pipe = true;
      let isPE_Heavy_WaxPowder = true;
      let isPE_Heavy_Off = true;
      let isPE_Heavy_Badi = true;
      let isPE_Heavy_Injection = true;
      let isPE_Heavy_Film = true;
      let isPE_Heavy_Dorani = true;

      let dataRowCount = 0;

      ws.eachRow((row, rowNumber) => {
        if (rowNumber <= 3) return;
        dataRowCount++;

        const namaKala = colNamaKala > 0 ? normalizeExcelCell(row.getCell(colNamaKala).value) : "";
        const talar = colTalar > 0 ? normalizeExcelCell(row.getCell(colTalar).value) : "";
        const nkLower = namaKala.toLowerCase().replace(/[\s\u200C]+/g, "");
        const containsOff = nkLower.includes("off");

        if (talar === "تالار سیمان") isSeman = true;

        isPET_Bottle = isPET_Bottle && (namaKala.includes("پلی") && namaKala.includes("بطری"));

        const containsPPTextile = namaKala.includes("پلی پروپیلن") && namaKala.includes("نساجی");
        isPPTextile = isPPTextile && (containsPPTextile && !containsOff);
        isPPTextileOff = isPPTextileOff && (containsPPTextile && containsOff);

        const containsPPChem = namaKala.includes("پلی پروپیلن") && namaKala.includes("شیمیایی");
        isPPChemical = isPPChemical && (containsPPChem && !containsOff);
        isPPChemicalOff = isPPChemicalOff && (containsPPChem && containsOff);

        const containsPVC = namaKala.includes("پی وی سی") || namaKala.includes("پلی وینیل کلراید");
        isPVC_S = isPVC_S && (containsPVC && namaKala.includes("S"));
        isPVC_E = isPVC_E && (containsPVC && namaKala.includes("E"));

        isPolystyrene = isPolystyrene && (namaKala.includes("پلی") && namaKala.includes("استایرن"));

        isDEGandMEG = isDEGandMEG && (namaKala.includes("منو اتیلن گلایکول") || namaKala.includes("دی اتیلن گلایکول") || namaKala.includes("تری اتیلن گلایکول"));

        isAsid = isAsid && namaKala.includes("اسید ترفتالیک");

        const isPE_Light = namaKala.includes("پلی اتیلن سبک");
        isPE_Light_Injection = isPE_Light_Injection && (isPE_Light && namaKala.includes("تزریقی") && !containsOff);
        isPE_Light_Linear = isPE_Light_Linear && (isPE_Light && namaKala.includes("خطی") && !containsOff);
        isPE_Light_Film = isPE_Light_Film && (isPE_Light && namaKala.includes("فیلم") && !containsOff);
        isPE_Light_Off = isPE_Light_Off && (isPE_Light && containsOff);

        const isPE_Heavy = namaKala.includes("پلی اتیلن سنگین") || (namaKala.includes("پلی") && namaKala.includes("سنگین"));
        isPE_Heavy_Pipe = isPE_Heavy_Pipe && (isPE_Heavy && (namaKala.includes("اکستروژن") || namaKala.includes("لوله")) && !containsOff);
        isPE_Heavy_WaxPowder = isPE_Heavy_WaxPowder && (isPE_Heavy && (namaKala.includes("PEWAX") || namaKala.includes("کلوخه") || namaKala.includes("پودر")) && !containsOff);
        isPE_Heavy_Off = isPE_Heavy_Off && (isPE_Heavy && containsOff);
        isPE_Heavy_Badi = isPE_Heavy_Badi && (isPE_Heavy && namaKala.includes("بادی") && !containsOff);
        isPE_Heavy_Injection = isPE_Heavy_Injection && (isPE_Heavy && namaKala.includes("تزریقی") && !containsOff);
        isPE_Heavy_Film = isPE_Heavy_Film && (isPE_Heavy && namaKala.includes("فیلم") && !containsOff);
        isPE_Heavy_Dorani = isPE_Heavy_Dorani && (isPE_Heavy && namaKala.includes("دورانی") && !containsOff);
      });

      if (dataRowCount === 0) return;

      isTargetSheet =
        isSeman ||
        isPET_Bottle ||
        isPPTextile ||
        isPPTextileOff ||
        isPPChemical ||
        isPPChemicalOff ||
        isPVC_S ||
        isPVC_E ||
        isPolystyrene ||
        isDEGandMEG ||
        isAsid ||
        isPE_Light_Injection ||
        isPE_Light_Linear ||
        isPE_Light_Film ||
        isPE_Light_Off ||
        isPE_Heavy_Pipe ||
        isPE_Heavy_WaxPowder ||
        isPE_Heavy_Off ||
        isPE_Heavy_Badi ||
        isPE_Heavy_Injection ||
        isPE_Heavy_Film ||
        isPE_Heavy_Dorani;

      if (!isTargetSheet) {
        return;
      }

      const lastDataRow = ws.lastRow?.number || 3;
      const volumeLetter = ws.getColumn(volumeColIndex).letter;
      const mergeStartCol = Math.min(volumeColIndex, increaseColIndex);
      const mergeEndCol = Math.max(volumeColIndex, increaseColIndex);

      const clearSummaryRow = (row) => {
        for (let colIndex = 1; colIndex <= headersArr.length; colIndex++) {
          const cell = row.getCell(colIndex);
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFFFF" }
          };
          cell.border = {};
        }
      };

      // ===== ردیف عنوان جمع =====
      const titleRow = ws.addRow([]);
      clearSummaryRow(titleRow);

      ws.mergeCells(titleRow.number, mergeStartCol, titleRow.number, mergeEndCol);
      const titleCell = ws.getCell(titleRow.number, mergeStartCol);
      titleCell.value = "جمع کل عرضه ها";
      titleCell.font = { name: "Arial", bold: true, size: 14 };
      titleCell.alignment = { horizontal: "center", vertical: "middle" };
      titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFCC99" } };
      applyAllBorders(titleCell);
      titleRow.height = 28;

      // ===== ردیف مقدار جمع =====
      const sumRow = ws.addRow([]);
      clearSummaryRow(sumRow);

      ws.mergeCells(sumRow.number, mergeStartCol, sumRow.number, mergeEndCol);
      const sumCell = ws.getCell(sumRow.number, mergeStartCol);
      sumCell.value = {
        formula: `SUM(${volumeLetter}4:${volumeLetter}${lastDataRow})`
      };
      sumCell.numFmt = "#,##0";
      sumCell.font = { name: "Arial", bold: true, size: 13 };
      sumCell.alignment = { horizontal: "center", vertical: "middle" };
      sumCell.fill = { type: "pattern", pattern: "solid", fgColor: "white" };
      applyAllBorders(sumCell);
      sumRow.height = 25;

      // ===== ردیف عنوان اختلاف =====
      const diffTitleRow = ws.addRow([]);
      clearSummaryRow(diffTitleRow);

      ws.mergeCells(diffTitleRow.number, mergeStartCol, diffTitleRow.number, mergeEndCol);
      const diffTitleCell = ws.getCell(diffTitleRow.number, mergeStartCol);
      diffTitleCell.value = "اختلاف با هفته گذشته";
      diffTitleCell.font = { name: "Arial", bold: true, size: 14 };
      diffTitleCell.alignment = { horizontal: "center", vertical: "middle" };
      diffTitleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFCC99" } };
      applyAllBorders(diffTitleCell);
      diffTitleRow.height = 28;

      // ===== ردیف مقدار اختلاف =====
      const diffRow = ws.addRow([]);
      clearSummaryRow(diffRow);

      ws.mergeCells(diffRow.number, mergeStartCol, diffRow.number, mergeEndCol);
      const diffCell = ws.getCell(diffRow.number, mergeStartCol);

      const previousWeekCell = ws.getCell(diffRow.number + 1, mergeStartCol);
      diffCell.value = {
        formula: `${sumCell.address}-${previousWeekCell.address}`
      };
      diffCell.numFmt = "#,##0";
      diffCell.font = { name: "Arial", bold: true, size: 13 };
      diffCell.alignment = { horizontal: "center", vertical: "middle" };
      diffCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFCCFFCC" } };
      applyAllBorders(diffCell);
      diffRow.height = 25;
    });

    // مخفی‌کردن ستون «تالار» در تمام شیت‌ها
    const talarColumnIndex = headersArr.indexOf("تالار") + 1;
    if (talarColumnIndex > 0) {
      workbook.eachSheet((ws) => {
        ws.getColumn(talarColumnIndex).hidden = true;
      });
    }

    // AutoFit عرض ستون‌ها
    const volumeColIndex = headersArr.indexOf("مقدار پایه") + 1;
    const increaseColIndex = headersArr.indexOf("حداکثر افزایش عرضه") + 1;

    workbook.eachSheet((ws) => {
      ws.columns.forEach((column) => {
        let maxLength = 0;

        column.eachCell({ includeEmpty: false }, (cell, rowNumber) => {
          if (rowNumber >= 3) {
            const isSummaryRow = 
              cell.value === "جمع کل عرضه ها" || 
              cell.value === "اختلاف با هفته گذشته" ||
              (cell.value && typeof cell.value === "object" && cell.value.formula);

            if (!isSummaryRow) {
              let strVal = "";
              if (typeof cell.value === "number") {
                strVal = cell.value.toLocaleString("en-US");
              } else if (cell.value) {
                strVal = cell.value.toString();
              }

              if (strVal.length > maxLength) {
                maxLength = strVal.length;
              }
            }
          }
        });

        const calculatedWidth = Math.ceil(maxLength * 1.35) + 4;

        if (column.number === volumeColIndex || column.number === increaseColIndex) {
          column.width = Math.max(calculatedWidth, 15);
        } else {
          column.width = Math.max(calculatedWidth, 12);
        }
      });
    });

    // تنظیمات چاپ (A4 افقی + Fit to Width + فوتر)
    workbook.eachSheet((ws) => {
      const lastRow = ws.lastRow?.number || 1;
      const lastColumn = headersArr.length;
      const lastColLetter = ws.getColumn(lastColumn).letter;

      ws.properties.pageSetUpPr = {
        fitToPage: true,
        autoPageBreaks: false
      };

      ws.pageSetup = {
        paperSize: 9,
        orientation: "landscape",
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        horizontalCentered: true,
        verticalCentered: false,
        margins: {
          left: 0.25,
          right: 0.25,
          top: 0.35,
          bottom: 0.35,
          header: 0.2,
          footer: 0.2
        },
        printTitlesRow: "1:3",
        printArea: `A1:${lastColLetter}${lastRow}`
      };

      ws.printOptions = {
        horizontalCentered: true,
        verticalCentered: false,
        gridLines: false,
        headings: false
      };

      ws.headerFooter = {
        oddFooter: '&L&"Arial"صفحه &P از &N&R&"Arial"تاریخ چاپ: &D',
        differentFirst: false,
        differentOddEven: false
      };
    });

    // درج واترمارک
    const watermarkBase64 = createWatermark("کارگزاری آینده نگر خوارزمی", "09123011311");
    const watermarkImageId = workbook.addImage({
      base64: watermarkBase64,
      extension: "png",
    });

    const ROWS_PER_PAGE = 22; 

    workbook.eachSheet((ws) => {
      const colCount = headersArr.length || 10;
      const centerCol = Math.max(1, Math.floor(colCount / 2) - 2);
      const totalRows = ws.rowCount || 1;

      const estimatedPages = Math.ceil(totalRows / ROWS_PER_PAGE);

      for (let page = 0; page < estimatedPages; page++) {
        const startRowOfPage = page * ROWS_PER_PAGE + 1;
        const centerRowForThisPage = startRowOfPage + Math.floor(ROWS_PER_PAGE / 2) - 3;

        ws.addImage(watermarkImageId, {
          tl: { col: centerCol, row: Math.max(1, centerRowForThisPage) },
          ext: { width: 700, height: 280 },
          editAs: "oneCell",
        });
      }
    });

    const dateStr = new Date().toISOString().slice(0,10);
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {type:"application/octet-stream"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `خروجی-${dateStr}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="container mt-4" dir="rtl">
      <h3>پردازش اکسل (RTL)</h3>
      <input
        type="file"
        accept=".xlsx,.xls"
        className="form-control mb-2"
        onChange={handleFileUpload}
      />
      <button
        className="btn btn-primary"
        onClick={processData}
        disabled={data.length===0}
      >
        پردازش و دانلود خروجی
      </button>
    </div>
  );
}
