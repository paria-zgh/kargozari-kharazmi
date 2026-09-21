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
    ctx.font = "bold 250px B Nazanin";
    ctx.fillText(text1, 0, -180); // فاصله از مرکز بیشتر تا جدا شود

    // === خط دوم ===
    ctx.font = "bold 230px B Nazanin";
    ctx.fillText(text2, 0, 260); // فاصله زیادتر تا در هم نرود

    return canvas.toDataURL("image/png");
  };

  // --- helper: افزودن واترمارک به صورت کاملاً مرکزچین در تمامی شیت‌ها و صفحات ---
  const addCenteredWatermarks = (workbook, worksheet, watermarkBase64, totalColumns) => {
    const imageId = workbook.addImage({
      base64: watermarkBase64,
      extension: "png",
    });

    const rowsPerPage = 32; // تعداد ردیف در هر صفحه چاپی استاندارد
    const totalRows = Math.max(worksheet.rowCount || 1, 10);
    const pageCount = Math.max(1, Math.ceil((totalRows - 3) / rowsPerPage));

    // ستون شروع برای قرارگیری دقیق در وسط صفحه
    const centerCol = Math.max(1, Math.floor(totalColumns / 2) - 2);

    for (let page = 0; page < pageCount; page++) {
      const startRow = 4 + page * rowsPerPage;
      const endRow = Math.min(startRow + rowsPerPage - 1, totalRows);
      const centerRow = Math.max(4, Math.floor((startRow + endRow) / 2) - 3);

      worksheet.addImage(imageId, {
        tl: { col: centerCol, row: centerRow },
        ext: { width: 700, height: 250 },
        editAs: "undefined",
      });
    }
  };

  const processData = async () => {
    if (data.length === 0) {
      alert("ابتدا یک فایل آپلود کنید.");
      return;
    }

    const cleanStr = (s) =>
      String(s || "")
        .replace(/ي/g, "ی")
        .replace(/ك/g, "ک")
        .replace(/\u200C/g, " ")
        .trim();

    // Sort and block logic unchanged but data is already normalized strings
    const sorted = [...data].sort((a, b) =>
      String(a["تالار"] || "").localeCompare(String(b["تالار"] || ""), "fa")
    );

    const keywordsSub = [
      "مس","فولاد","ضایعات","پالت چوبی","بشکه خالی","ورق",
      "پشم شیشه","کنسانتره","اقلام تجهیزات","اقلام تکمیلی خودرو",
      "شمش","سپری","تختال","بتنی","گرانول نقره","فلز","بیلت"
    ];
   // const petrochemKeywords = ["پلی","الیاف استیپل اکریلیک"];
    const petroleumKeywords = ["قیر", "نفت", "واکس", "گوگرد", "اسلاک" ,"پارافین","وکیوم","روغن","لوب"];

    const blocks = {
      keywordBlockSub: [],
      keywordBlockPetroleumFromSub: [],
      petroleumFromPetrochem: [],
      auctionBlockStoneOrCathode: [],
      auctionBlockVacuum: [],
      petrochemBlock: [],
      otherRows: [],
    };

    sorted.forEach((row) => {
      const curTalar = cleanStr(row["تالار"]);
      const namaKala = cleanStr(row["نام کالا"]);
    
      if (curTalar.includes("فرعی")) {
        if (["نفتی", "نفت", "وکیوم", "قیر", "روغن"].some((kw) => namaKala.includes(kw))) {
          blocks.keywordBlockPetroleumFromSub.push(row);
        } else if (keywordsSub.some((kw) => namaKala.includes(kw))) {
          blocks.keywordBlockSub.push({
            ...row,
            تالار: "تالار صنعتی و معدنی",
          });
        } else {
          blocks.petrochemBlock.push({
            ...row,
            تالار: "تالار پتروشیمی و فرآورده های نفتی",
          });
        }
      } 
      else if (curTalar.includes("حراج باز")) {
        if (
          namaKala.includes("سنگ") ||
          namaKala.includes("مس کاتد") ||
          namaKala.includes("فلز") ||
          namaKala.includes("تختال") ||
          namaKala.includes("اکسید مولیبدن")
        ) {
          blocks.auctionBlockStoneOrCathode.push({
            ...row,
            تالار: "تالار صنعتی و معدنی",
          });
        } else if (namaKala.includes("وکیوم")) {
          blocks.auctionBlockVacuum.push(row);
        } else {
          blocks.petrochemBlock.push(row);
        }
      } 
      else if (curTalar.includes("پتروشیمی")) {
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

    const industrialRows = blocks.otherRows.filter((r) =>
      cleanStr(r["تالار"]).includes("صنعتی")
    );
    const restOfOtherRows = blocks.otherRows.filter(
      (r) => !cleanStr(r["تالار"]).includes("صنعتی")
    );

    // ساخت یک بلوک مستقل برای فرآورده‌های نفتی
    const petroleumRows = [
      ...blocks.petroleumFromPetrochem,
      ...blocks.keywordBlockPetroleumFromSub,
      ...blocks.auctionBlockVacuum
    ].map((row) => ({
      ...row,
      تالار: "تالار فرآورده های نفتی",
    }));

    const finalRows = [
      ...blocks.petrochemBlock,
      ...industrialRows,
      ...blocks.keywordBlockSub,
      ...blocks.auctionBlockStoneOrCathode,
      ...restOfOtherRows,
      ...petroleumRows,
    ];
    
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
    

    const specialKeywords=["بوتادین استایرن","استایرن بوتادین","تیشو","پلی","الیاف استیپل اکریلیک","ABS"];
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
        blocksSorted.push(row);
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
          nr[h]=isNaN(vol) ? null : vol/1000; // 👈 null به جای "" برای سلول خالی
                
        }else if(h==="قیمت پایه"){
            const price=Number(String(value).replace(/,/g,''));
            nr[h]=isNaN(price) ? null : price; // 👈 عدد واقعی
          
          
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
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی وینیل")  && r["نام کالا"]?.includes("کلراید")&& r["نام کالا"]?.includes("S")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی وینیل")  && r["نام کالا"]?.includes("کلراید")&& r["نام کالا"]?.includes("E")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی") && r["نام کالا"]?.includes("استایرن"))) ;
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("منو اتیلن گلایکول") || r["نام کالا"]?.includes("دی اتیلن گلایکول")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("تری اتیلن گلایکول"))) ;
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("اسید ترفتالیک"))) ;
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی اتیلن سبک")&& r["نام کالا"]?.includes("تزریقی") && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی اتیلن سبک") && r["نام کالا"]?.includes("خطی") && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی اتیلن سبک") && r["نام کالا"]?.includes("فیلم") && !r["نام کالا"]?.toLowerCase().includes("off")));
    addSpecialBlock(reordered.filter(r=>r["نام کالا"]?.includes("پلی اتیلن سبک") && r["نام کالا"]?.includes("off")));
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
    const newSpecialRows = reordered.filter(r => 
      r["نام کالا"]?.includes("تولوئن دی ایزو سیانات") || 
      r["نام کالا"]?.includes("متیلن دی فنیل دی ایزوسیانات خالص")||
      r["نام کالا"]?.includes("متیلن دی فنیل ایزوسیانات خالص")
    );

    if(newSpecialRows.length > 0){
      reordered.push({__HEADER__:true});
      const excludeColumns = ["کد عرضه", "تولید کننده کالا", "محل تحویل", "حداقل خرید"];
      const specialHeaders = headersArr.filter(h => !excludeColumns.includes(h));

      newSpecialRows.forEach(r=>{
        const rowData = specialHeaders.map(h => r[h] ?? "");
        const newRow = {};
        specialHeaders.forEach((h,i)=>{ newRow[h] = rowData[i]; });
        reordered.push(newRow);
      });
    }

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
        cell.font={name:"B Nazanin",bold};
        cell.alignment={vertical:"middle",horizontal:"center"};
        applyAllBorders(cell);
      });
    };

    const addTitleRows = (ws) => {
      const titleRow = ws.addRow([titleText]);
      titleRow.height = 33;
      ws.mergeCells(titleRow.number, 1, titleRow.number, headersArr.length);
    
      // --- اصلاح اعمال فونت روی سلول‌های merge شده ---
      titleRow.eachCell((cell) => {
        cell.font = { name: "B Nazanin", bold: true, size: 12 };
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFA500" } };
        applyAllBorders(cell);
      });
    
      const supplyRow = ws.addRow([supplyText]);
      supplyRow.height = 40;
      ws.mergeCells(supplyRow.number,1,supplyRow.number,headersArr.length);
    
      // --- همین اصلاح برای ردیف دوم ---
      supplyRow.eachCell((cell) => {
        cell.font = { name: "B Nazanin", bold: true, size: 18 };
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
        applyAllBorders(cell);
      });
    
      const headerRow = ws.addRow(headersArr);
      headerRow.height = 30;
      headerRow.eachCell((cell) => {
        cell.font = { name: "B Nazanin", bold: true };
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD3D3D3" } };
        applyAllBorders(cell);
      });
    };
    
    let sheetCounter = 1;
    let currentSheet = null;
    let rowCounter = 0;
    
    reordered.forEach((row, index) => {
      if (row.__HEADER__) {
        // پیدا کردن داده‌های بعد از این HEADER تا HEADER بعدی یا انتهای آرایه
        const nextHeaderIndex = reordered.slice(index + 1).findIndex(r => r.__HEADER__);
        const dataForSheet = nextHeaderIndex === -1 
          ? reordered.slice(index + 1)
          : reordered.slice(index + 1, index + 1 + nextHeaderIndex);
    
        // فقط اگر داده واقعی هست شیت بساز
        if(dataForSheet.some(r => Object.keys(r).length > 0)){
          sheetCounter++;
          currentSheet = workbook.addWorksheet(`Sheet ${sheetCounter}`, {views:[{rightToLeft:true}]});
          addTitleRows(currentSheet);
          rowCounter = 0;
        }
      } else if(Object.keys(row).length > 0){
        // اضافه کردن ردیف داده واقعی
        if(!currentSheet){
          // برای اولین شیت که هنوز ساخته نشده
          sheetCounter = 1;
          currentSheet = workbook.addWorksheet(`Sheet ${sheetCounter}`, {views:[{rightToLeft:true}]});
          addTitleRows(currentSheet);
          rowCounter = 0;
        }
    
        const dataRow = currentSheet.addRow(
          headersArr.map(h => {
            const val = row[h];
            if (val !== "" && !isNaN(val)) return Number(val);
            return val ?? "";
          })
        );
        styleRow(dataRow, null, false);
        dataRow.height = 23;
    
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
    // Helper to normalize cell/row values coming from the generated workbook (still safe)
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
    
      // فلگ‌ها برای سایر محصولات
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
      let allteg = true;
      let asid = true;
    
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
        if (rowNumber <= 3) return; // skip title/supply/header rows
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
        allteg = allteg && namaKalaRaw.includes("تری اتیلن گلایکول");
        allPVC_S = allPVC_S && (namaKalaRaw.includes("پلی وینیل کلراید") && namaKalaRaw.includes("S"));
        allPVC_E = allPVC_E && (namaKalaRaw.includes("پلی وینیل کلراید") && namaKalaRaw.includes("E"));
        allDEG = allDEG && namaKalaRaw.includes("دی اتیلن گلایکول");
        allMEG = allMEG && namaKalaRaw.includes("منو اتیلن گلایکول");
        allDEGandMEG = allDEGandMEG && (namaKalaRaw.includes("دی اتیلن گلایکول") || namaKalaRaw.includes("منو اتیلن گلایکول"));
        allPET_Bottle = allPET_Bottle && namaKalaRaw.includes("پلی اتیلن ترفتالات بطری");
        allPolystyrene = allPolystyrene && namaKalaRaw.includes("پلی استایرن");
      });
    
      let isAllowedForSummary = false;

      // جایگزینی X در هدر
      const headerRow = ws.getRow(2);
      headerRow.eachCell((cell) => {
        if (typeof cell.value === "string" && cell.value.includes("X")) {
          // اول پلی اتیلن سنگین
          for (let b of heavyPEBlocks) {
            if (b.allPresent) {
              cell.value = cell.value.replace("X", b.headerText);
              isAllowedForSummary = true;
              return;
            }
          }
    
          // سپس پلی اتیلن سبک
          for (let b of specialPEBlocks) {
            if (b.offHeader && b.allPresent && b.allOff) {
              cell.value = cell.value.replace("X", b.offHeader);
              isAllowedForSummary = true;
              return;
            } else if (b.allPresent) {
              cell.value = cell.value.replace("X", b.headerText);
              isAllowedForSummary = true;
              return;
            }
          }
    
          // سایر محصولات قبلی
          if (allPPTextileOff) { cell.value = cell.value.replace("X", "های OFF پلی پروپیلن نساجی"); isAllowedForSummary = true; }
          else if (allPPTextile) { cell.value = cell.value.replace("X", "پلی پروپیلن نساجی"); isAllowedForSummary = true; }
          else if (allPPChemicalOff) { cell.value = cell.value.replace("X", "های OFF پلی پروپیلن شیمیایی"); isAllowedForSummary = true; }
          else if (allPPChemical) { cell.value = cell.value.replace("X", "پلی پروپیلن شیمیایی"); isAllowedForSummary = true; }
          else if (allPVC_S) { cell.value = cell.value.replace("X", "پلی وینیل کلراید PVC(S)"); isAllowedForSummary = true; }
          else if (allPVC_E) { cell.value = cell.value.replace("X", "پلی وینیل کلراید PVC(E)"); isAllowedForSummary = true; }
          else if (allDEGandMEG) { cell.value = cell.value.replace("X", "DEG & MEG"); isAllowedForSummary = true; }
          else if (allDEG) { cell.value = cell.value.replace("X", "DEG"); isAllowedForSummary = true; }
          else if (allMEG) { cell.value = cell.value.replace("X", "MEG"); isAllowedForSummary = true; }
          else if (allPET_Bottle) { cell.value = cell.value.replace("X", "پلی اتیلن ترفتالات بطری"); isAllowedForSummary = true; }
          else if (allPolystyrene) { cell.value = cell.value.replace("X", "پلی استایرن"); isAllowedForSummary = true; }
          else if (allteg) { cell.value = cell.value.replace("X", "تری اتیلن گلایکول"); isAllowedForSummary = true; }
          else if (asid) { cell.value = cell.value.replace("X", "اسید ترفتالیک"); isAllowedForSummary = true; }
          else if (hasIndustrial) { cell.value = cell.value.replace("X", "محصولات صنعتی"); }
          else if (hasSeman) { cell.value = cell.value.replace("X", "سیمان"); isAllowedForSummary = true; }
          else if (hasPetroleum) { cell.value = cell.value.replace("X", "فرآورده های نفتی و قیر"); }
          else if (hasExport) { cell.value = cell.value.replace("X", "تالار صادراتی"); }
          else if (hasPetrochemBlock) { cell.value = cell.value.replace("X", "محصولات پلیمری"); isAllowedForSummary = true; }
          else if (notPetrochemBlock) { cell.value = cell.value.replace("X", "محصولات شیمیایی"); isAllowedForSummary = true; }

        }
      });

      ws._isAllowedForSummary = isAllowedForSummary;
    });

    // === افزودن جمع کل عرضه‌ها و اختلاف به پایین شیت‌های مجاز ===
    workbook.eachSheet((ws) => {
      if (!ws._isAllowedForSummary) {
        return;
      }

      // در این کد، ستون مقدار پایه معادل ستون حجم عرضه در کد قبلی است
      const volumeColIndex = headersArr.indexOf("مقدار پایه") + 1;
      const increaseColIndex =
        headersArr.indexOf("حداکثر افزایش عرضه") + 1;

      // فقط در صورتی اجرا شود که هر دو ستون وجود داشته باشند
      if (volumeColIndex <= 0 || increaseColIndex <= 0) {
        return;
      }

      // آخرین ردیف داده قبل از اضافه‌شدن ردیف‌های جمع
      const lastDataRow = ws.lastRow?.number || 3;

      // حرف ستون مقدار پایه، برای مثال I یا J
      const volumeLetter = ws.getColumn(volumeColIndex).letter;

      // ابتدا و انتهای محدوده Merge
      const mergeStartCol = Math.min(
        volumeColIndex,
        increaseColIndex
      );

      const mergeEndCol = Math.max(
        volumeColIndex,
        increaseColIndex
      );

      // سفید و بدون Border کردن سلول‌های کناری
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

      ws.mergeCells(
        titleRow.number,
        mergeStartCol,
        titleRow.number,
        mergeEndCol
      );

      const titleCell = ws.getCell(
        titleRow.number,
        mergeStartCol
      );

      titleCell.value = "جمع کل عرضه ها";

      titleCell.font = {
        name: "B Nazanin",
        bold: true,
        size: 14
      };

      titleCell.alignment = {
        horizontal: "center",
        vertical: "middle"
      };

      titleCell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFCC99" }
      };

      applyAllBorders(titleCell);

      titleRow.height = 28;

      // ===== ردیف مقدار جمع =====
      const sumRow = ws.addRow([]);

      clearSummaryRow(sumRow);

      ws.mergeCells(
        sumRow.number,
        mergeStartCol,
        sumRow.number,
        mergeEndCol
      );

      const sumCell = ws.getCell(
        sumRow.number,
        mergeStartCol
      );

      sumCell.value = {
        formula: `SUM(${volumeLetter}4:${volumeLetter}${lastDataRow})`
      };

      sumCell.numFmt = "#,##0";

      sumCell.font = {
        name: "B Nazanin",
        bold: true,
        size: 13
      };

      sumCell.alignment = {
        horizontal: "center",
        vertical: "middle"
      };

      sumCell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: "white",
      };

      applyAllBorders(sumCell);

      sumRow.height = 25;

      // ===== ردیف عنوان اختلاف =====
      const diffTitleRow = ws.addRow([]);

      clearSummaryRow(diffTitleRow);

      ws.mergeCells(
        diffTitleRow.number,
        mergeStartCol,
        diffTitleRow.number,
        mergeEndCol
      );

      const diffTitleCell = ws.getCell(
        diffTitleRow.number,
        mergeStartCol
      );

      diffTitleCell.value = "اختلاف با هفته گذشته";

      diffTitleCell.font = {
        name: "B Nazanin",
        bold: true,
        size: 14
      };

      diffTitleCell.alignment = {
        horizontal: "center",
        vertical: "middle"
      };

      diffTitleCell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFCC99" }
      };

      applyAllBorders(diffTitleCell);

      diffTitleRow.height = 28;

      // ===== ردیف مقدار اختلاف =====
      const diffRow = ws.addRow([]);

      clearSummaryRow(diffRow);

      ws.mergeCells(
        diffRow.number,
        mergeStartCol,
        diffRow.number,
        mergeEndCol
      );

      const diffCell = ws.getCell(
        diffRow.number,
        mergeStartCol
      );

      // جمع عرضه فعلی منهای مقدار هفته گذشته
      // مقدار هفته گذشته باید در ردیف زیر این سلول وارد شود
      const previousWeekCell = ws.getCell(
        diffRow.number + 1,
        mergeStartCol
      );

      diffCell.value = {
        formula: `${sumCell.address}-${previousWeekCell.address}`
      };

      diffCell.numFmt = "#,##0";

      diffCell.font = {
        name: "B Nazanin",
        bold: true,
        size: 13
      };

      diffCell.alignment = {
        horizontal: "center",
        vertical: "middle"
      };

      diffCell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFCCFFCC" }
      };

      applyAllBorders(diffCell);

      diffRow.height = 25;
    });

    // === درج واترمارک مرکزچین در تمام شیت‌ها و صفحات ===
    const watermarkBase64 = createWatermark("کارگزاری آینده نگر خوارزمی", "09123011311");
    workbook.eachSheet((ws) => {
      addCenteredWatermarks(workbook, ws, watermarkBase64, headersArr.length);
    });

    // === AutoFit عرض ستون‌ها بر اساس محتوا ===
    // ==========================================
    // ۱) ابتدا: بولد (Bold) کردن تمامی سلول‌ها در همه شیت‌ها
    // ==========================================
    workbook.eachSheet((ws) => {
      ws.eachRow({ includeEmpty: false }, (row) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
          cell.font = {
            ...(cell.font || {}),
            bold: true
          };
        });
      });
    });

    // مخفی‌کردن ستون «تالار» در تمام شیت‌ها
    const talarColumnIndex = headersArr.indexOf("تالار") + 1;
    if (talarColumnIndex > 0) {
      workbook.eachSheet((ws) => {
        ws.getColumn(talarColumnIndex).hidden = true;
      });
    }

    // ==========================================
    // ۲) سپس: تنظیم AutoFit عرض ستون‌ها (بعد از Bold شدن)
    // ==========================================
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
              const cellValue = cell.value ? cell.value.toString() : "";
              maxLength = Math.max(maxLength, cellValue.length);
            }
          }
        });

        // چون متن‌ها بولد شده‌اند، فضای اضافه (padding) را به +4 و حداقل را به 12 افزایش دادیم تا متن بریده نشود
        let finalWidth = Math.min(Math.max(maxLength + 7, 12), 36);
        if (column.number === volumeColIndex || column.number === increaseColIndex) {
          finalWidth = Math.min(Math.max(maxLength + 7, 15), 20);
        }
        column.width = finalWidth;
      });
    });


    // ==========================================
    // ۳) تنظیمات کامل و استاندارد چاپ (A4 افقی + Fit to Width + فوتر)
    // ==========================================
    workbook.eachSheet((ws) => {
      const lastRow = ws.lastRow?.number || 1;
      const lastColumn = headersArr.length;
      const lastColLetter = ws.getColumn(lastColumn).letter;

      ws.properties.pageSetUpPr = {
        fitToPage: true,
        autoPageBreaks: false
      };

      ws.pageSetup = {
        paperSize: 9,              // اندازه A4
        orientation: "landscape",  // حالت افقی
        fitToPage: true,
        fitToWidth: 1,             // تمام ستون‌ها در یک عرض جا شوند
        fitToHeight: 0,            // طول صفحه بر اساس تعداد سطرها آزاد باشد
        horizontalCentered: true,  // وسط‌چین در چاپ
        verticalCentered: false,
        margins: {
          left: 0.25,
          right: 0.25,
          top: 0.35,
          bottom: 0.35,
          header: 0.2,
          footer: 0.2
        },
        printTitlesRow: "1:3",     // تکرار ۳ سطر عنوان در بالای همه صفحات
        printArea: `A1:${lastColLetter}${lastRow}`
      };

      ws.printOptions = {
        horizontalCentered: true,
        verticalCentered: false,
        gridLines: false,
        headings: false
      };

      ws.headerFooter = {
        oddFooter: '&L&"B Nazanin"صفحه &P از &N&R&"B Nazanin"تاریخ چاپ: &D',
        differentFirst: false,
        differentOddEven: false
      };
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
