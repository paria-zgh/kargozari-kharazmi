import React, { useState } from "react";
import ExcelJS from "exceljs";
import jalaali from "jalaali-js";

export default function Main() {
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
const createWatermark = (line1, line2, width=650, height=100) => {
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");

  ctx.globalAlpha = 0.25; // شفافیت کم
  ctx.fillStyle = "green";

  ctx.translate(canvas.width / 6, canvas.height / 2);
  ctx.rotate(-10 * Math.PI / 180); // زاویه 30 درجه
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";

  ctx.font = "bold 24px B Nazanin";
  ctx.fillText(line1, 0, -10);  // خط اول
  ctx.fillText(line2, 0, 20);   // خط دوم

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
      "مس","فولاد","ضایعات","پالت چوبی","بشکه خالی","ورق",
      "پشم شیشه","کنسانتره","اقلام تجهیزات","اقلام تکمیلی خودرو",
      "شمش","سپری","تختال","بتنی","گرانول نقره","فلز","بیلت"
    ];
    const petrochemKeywords = ["پلی","الیاف استیپل اکریلیک"];

    const blocks = {
      keywordBlockSub: [],
      keywordBlockPetroleumFromSub: [],
      auctionBlockStoneOrCathode: [],
      auctionBlockVacuum: [],
      petrochemBlock: [],
      otherRows: [],
    };

    sorted.forEach((row) => {
      const curTalar = row["تالار"] || "";
      const namaKala = row["نام کالا"] || "";

       
      if(curTalar==="تالار فرعی"){
        if(["نفتی","نفت","وکیوم","قیر","روغن"].some(kw=>namaKala.includes(kw))){
          blocks.keywordBlockPetroleumFromSub.push(row);
        } else if(keywordsSub.some(kw=>namaKala.includes(kw))){
          blocks.keywordBlockSub.push(row);
        } else if(petrochemKeywords.some(kw=>namaKala.includes(kw))){
          blocks.petrochemBlock.push(row);
        } else blocks.petrochemBlock.push(row);
      } else if(curTalar==="تالار حراج باز"){
        if(namaKala.includes("سنگ") || namaKala.includes("مس کاتد") || namaKala.includes("فلز")|| namaKala.includes("تختال")
        || namaKala.includes("اکسید مولیبدن")
        ){
          blocks.auctionBlockStoneOrCathode.push(row);
        } else if(namaKala.includes("وکیوم")){
          blocks.auctionBlockVacuum.push(row);
        } else if(petrochemKeywords.some(kw=>namaKala.includes(kw))){
          blocks.petrochemBlock.push(row);
        } else blocks.petrochemBlock.push(row);
      } else {
        blocks.otherRows.push(row);
      }
    });

    const finalRows = [...blocks.otherRows];
    let industrialIndex = finalRows.map(r => r["تالار"] || "").lastIndexOf("تالار صنعتی");
    if(industrialIndex===-1) industrialIndex=finalRows.length-1;
    if(blocks.keywordBlockSub.length>0){
      finalRows.splice(industrialIndex+1,0,...blocks.keywordBlockSub);
      industrialIndex+=blocks.keywordBlockSub.length;
    }
    if(blocks.auctionBlockStoneOrCathode.length>0){
      finalRows.splice(industrialIndex+1,0,...blocks.auctionBlockStoneOrCathode);
      industrialIndex+=blocks.auctionBlockStoneOrCathode.length;
    }

    let petroleumIndex = finalRows.map(r=>r["تالار"]||"").lastIndexOf("تالار فرآورده های نفتی");
    if(petroleumIndex===-1) petroleumIndex=finalRows.length-1;
    if(blocks.auctionBlockVacuum.length>0){
      finalRows.splice(petroleumIndex+1,0,...blocks.auctionBlockVacuum);
      petroleumIndex+=blocks.auctionBlockVacuum.length;
    }
    if(blocks.keywordBlockPetroleumFromSub.length>0){
      finalRows.splice(petroleumIndex+1,0,...blocks.keywordBlockPetroleumFromSub);
      petroleumIndex+=blocks.keywordBlockPetroleumFromSub.length;
    }

    let petroIndex = finalRows.map(r=>r["تالار"]||"").lastIndexOf("تالار پتروشیمی");
    if(petroIndex===-1) petroIndex=finalRows.length-1;
    if(blocks.petrochemBlock.length>0){
      finalRows.splice(petroIndex+1,0,...blocks.petrochemBlock);
      petroIndex+=blocks.petrochemBlock.length;
    }

    const insertTalarNames=["تالار صنعتی","تالار فرآورده های نفتی","تالار سیمان","تالار کالای صادراتی کيش"];
    let processed=[];
    const inserted=new Set();
    finalRows.forEach(row=>{
      const curTalar=row["تالار"]||"";
      if(insertTalarNames.includes(curTalar) && !inserted.has(curTalar)){
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

    // *** حذف ستون تالار پس از تکمیل تمام مرتب‌سازی‌ها ***


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
// --- حذف ستون تالار از headersArr ---

    // --- تولید اکسل با چند شیت ---
    const workbook = new ExcelJS.Workbook();
    workbook.views = [{ rightToLeft: true }];

    const titleText="کارگزاری آینده نگر خوارزمی به مدیریت دکتر ذوقی 09123011311";
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
      titleRow.height = 35;
      ws.mergeCells(titleRow.number, 1, titleRow.number, headersArr.length);
    
      // --- اصلاح اعمال فونت روی سلول‌های merge شده ---
      titleRow.eachCell((cell) => {
        cell.font = { name: "B Nazanin", bold: true, size: 12 };
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFA500" } };
        applyAllBorders(cell);
      });
    
      const supplyRow = ws.addRow([supplyText]);
      supplyRow.height = 30;
      ws.mergeCells(supplyRow.number,1,supplyRow.number,headersArr.length);
    
      // --- همین اصلاح برای ردیف دوم ---
      supplyRow.eachCell((cell) => {
        cell.font = { name: "B Nazanin", bold: true, size: 13 };
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
    
          // ---- اضافه کردن واترمارک ----
          const watermarkBase64 = createWatermark("کارگزاری آینده نگر خوارزمی", "09123011311");
          const imageId = workbook.addImage({
            base64: watermarkBase64,
            extension: "png",
          });
          currentSheet.addImage(imageId, {
            tl: { col: 1, row: 2 },
            ext: { width: 700, height: 250 },
            editAs: "oneCell"
          });
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
        if (rawTalar === "تالار صنعتی") hasIndustrial = true;
        if (rawTalar === "تالار سیمان") hasSeman = true;
        if (rawTalar === "تالار فرآورده های نفتی") hasPetroleum = true;
        if (rawTalar === "تالار کالای صادراتی کيش") hasExport = true;
        if (rawTalar === "تالار پتروشیمی" && petrochemKeywords.some(kw => namaKalaRaw.includes(kw))) {
          hasPetrochemBlock = true;
        }
        if (rawTalar === "تالار پتروشیمی" || rawTalar === "تالار حراج باز"||rawTalar === "تالار فرعی" && petrochemKeywords.some(kw => !namaKalaRaw.includes(kw))) {
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
    
          // سایر محصولات قبلی
          if (allPPTextileOff) cell.value = cell.value.replace("X", "های OFF پلی پروپیلن نساجی");
          else if (allPPTextile) cell.value = cell.value.replace("X", "پلی پروپیلن نساجی");
          else if (allPPChemicalOff) cell.value = cell.value.replace("X", "های OFF پلی پروپیلن شیمیایی");
          else if (allPPChemical) cell.value = cell.value.replace("X", "پلی پروپیلن شیمیایی");
          else if (allPVC_S) cell.value = cell.value.replace("X", "پلی وینیل کلراید PVC(S)");
          else if (allPVC_E) cell.value = cell.value.replace("X", "پلی وینیل کلراید PVC(E)");
          else if (allDEGandMEG) cell.value = cell.value.replace("X", "DEG & MEG");
          else if (allDEG) cell.value = cell.value.replace("X", "DEG");
          else if (allMEG) cell.value = cell.value.replace("X", "MEG");
          else if (allPET_Bottle) cell.value = cell.value.replace("X", "پلی اتیلن ترفتالات بطری");
          else if (allPolystyrene) cell.value = cell.value.replace("X", "پلی استایرن");
          else if (allteg) cell.value = cell.value.replace("X", "تری اتیلن گلایکول");
          else if (asid) cell.value = cell.value.replace("X", "اسید ترفتالیک");
          else if (hasIndustrial) cell.value = cell.value.replace("X", "محصولات صنعتی");
          else if (hasSeman) cell.value = cell.value.replace("X", "سیمان");
          else if (hasPetroleum) cell.value = cell.value.replace("X", "محصولات");
          else if (hasExport) cell.value = cell.value.replace("X", "تالار صادراتی");
          else if (hasPetrochemBlock) cell.value = cell.value.replace("X", "محصولات پلیمری");
          else if (notPetrochemBlock) cell.value = cell.value.replace("X", "محصولات شیمیایی");

        }
      });
    });
   
    // === AutoFit عرض ستون‌ها بر اساس محتوا ===
workbook.eachSheet((ws) => {
  // --- حذف ستون "تالار" از همه شیت‌ها بعد از جایگزینی X و اضافه کردن داده‌ها ---

 


  ws.columns.forEach((column) => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
      // فقط محتویات داده‌ها و هدر را حساب کن، ردیف عنوان (1 و 2) را می‌توان نادیده گرفت
      if (rowNumber >= 3) {
        const cellValue = cell.value ? cell.value.toString() : "";
        maxLength = Math.max(maxLength, cellValue.length);
      }
    });
    // عرض ستون حداقل 10، حداکثر 50 و کمی فاصله اضافه
    column.width = Math.min(Math.max(maxLength + 2, 10), 50);
  });
  
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