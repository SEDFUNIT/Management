// ฟังก์ชันสำหรับหน้า HTML
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle("ระบบเบิกจ่ายวัสดุ-อุปกรณ์ กทศ.")
    .setFaviconUrl("https://img.a4h6.c18.e2-4.dev/school-material.png")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** @Get URL */
function getURL() {
  return ScriptApp.getService().getUrl();
}

// ฟังก์ชันดึงไฟล์ HTML
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ฟังก์ชันดึงข้อมูลจาก Sheet
function getData(sh) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sh).getDataRange().getDisplayValues().slice(2);
}

function getDataApp(sh) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sh).getDataRange().getDisplayValues().slice(1).filter(r => r[1] == "รออนุมัติ");
}

// ฟังก์ชันบันทึกข้อมูลลงใน Google Sheets
function saveData(cart, fname, dpm) {
  if (!Array.isArray(cart) || cart.length === 0) {
    Logger.log('Invalid or empty cart data: ' + cart);
    return;
  }

  var bookno = uuid(); 
  var pad = "00000"; 
  var runid = pad.substring(0, pad.length - bookno.length) + bookno;

  const now = new Date();
  const resultDateTime = `${now.toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' })} ${now.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit', second: '2-digit' })}`;

  var msg = `แจ้งผลรายการขอเบิกวัสดุสำนักงาน
             \n 🔀เลขที่ : SEDF68/${runid}
             \n 📌ชื่อ-สกุล : ${fname}
             \n 🏢กลุ่มงาน : ${dpm}
             \n 📅วันที่เบิกวัสดุ : ${resultDateTime}`;

  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LogProduct');

  cart.forEach(r => {
    ss.appendRow([ 
      "",
      "รออนุมัติ",
      resultDateTime,
      `SEDF68/${runid}`,
      fname,
      dpm,
      `'${r.id}`,
      r.name,
      r.unix,
      r.count
    ]);

    msg += `\n📍 รหัสวัสดุ : ${r.id} 
            \n📝 รายการ : ${r.name} จำนวน ${r.count} ${r.unix}`;
  });

  var token = tokenID;
  sendNotify(msg, token);
}

// ฟังก์ชันอัปเดตข้อมูลใน Google Sheets
function toGoogleSheets(obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Product');
  var data = ss.getDataRange().getValues();
  var indexName = data.map(d => d[0]);
  var position = indexName.indexOf(obj.data1);
  if (position > -1) {
    ss.getRange(position + 1, 5).setValue(obj.data2);
  }
  return true;
}

// ฟังก์ชันบันทึกการอนุมัติหรือไม่อนุมัติรายการเดียว
function saveEditApp(id, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LogProduct');
  const data = sheet.getDataRange().getDisplayValues();
  const rowID = data.findIndex(row => row[0] == id) + 1;
  
  if (rowID > 1) {
    const now = new Date();
    const approvalDateTime = `${now.toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' })} ${now.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit', second: '2-digit' })}`;

    sheet.getRange(rowID, 2).setValue(status);
    sheet.getRange(rowID, 11).setValue(approvalDateTime);
    if (status === 'ไม่อนุมัติ') {
      sheet.getRange(rowID, 10).setValue('0');
    }
  }
}

// ฟังก์ชันบันทึกการอนุมัติหรือไม่อนุมัติหลายรายการ
function saveMultipleEntries(ids, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LogProduct');
  const data = sheet.getDataRange().getValues();

  ids.forEach(id => {
    const rowID = data.findIndex(row => row[0] == id) + 1;
    
    if (rowID > 1) {
      const now = new Date();
      const approvalDateTime = `${now.toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' })} ${now.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit', second: '2-digit' })}`;

      sheet.getRange(rowID, 2).setValue(status);
      sheet.getRange(rowID, 11).setValue(approvalDateTime);
      if (status === 'ไม่อนุมัติ') {
        sheet.getRange(rowID, 10).setValue('0');
      }
    }
  });
}

// ฟังก์ชันลบรายการที่เลือกหลายรายการ
function deleteMultipleEntries(ids) {
  if (!Array.isArray(ids) || ids.length === 0) {
    Logger.log('ไม่มีรายการที่เลือก');
    return; // ออกจากฟังก์ชันหากไม่มีรายการที่จะลบ
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LogProduct');
  const data = sheet.getDataRange().getDisplayValues();

  // ตรวจสอบข้อมูลที่ดึงมา
  Logger.log('ข้อมูลจากชีต:', data);

  // เก็บแถวที่ต้องการลบ
  const rowsToDelete = ids
    .map(id => {
      const rowID = data.findIndex(row => row[0] == id) + 1;
      Logger.log('ID:', id, 'ค้นพบที่แถว:', rowID);
      return rowID;
    })
    .filter(rowID => rowID > 0);
  
  if (rowsToDelete.length === 0) {
    Logger.log('ไม่พบรายการที่ตรงกับ id ที่ระบุ');
    return;
  }

  // เรียงลำดับแถวที่ต้องการลบจากมากไปน้อย
  rowsToDelete.sort((a, b) => b - a);
  Logger.log('แถวที่ต้องลบ:', rowsToDelete);

  // สร้างหรือเลือกชีต "DeleteLog"
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('DeleteLog');
  if (!logSheet) {
    logSheet = ss.insertSheet('DeleteLog');
    logSheet.appendRow(['ID', 'ที่', 'ผู้เบิก', 'กลุ่มงาน', 'เลขที่', 'รหัส', 'รายการ', 'จำนวน', 'หน่วยนับ', 'วันที่เบิก', 'File FDF']); // Header
  }

  // เก็บข้อมูลที่ถูกลบ
  const deletedRows = data.filter((row, index) => rowsToDelete.includes(index + 1));
  
  // เพิ่มข้อมูลที่ถูกลบลงในชีต "DeleteLog"
  deletedRows.forEach(row => logSheet.appendRow(row));
  
  // ลบแถวเริ่มจากแถวที่มีดัชนีสูงสุดเพื่อลดปัญหาการเลื่อนแถว
  rowsToDelete.forEach(rowID => {
    Logger.log('ลบแถวที่:', rowID);
    sheet.deleteRow(rowID);
  });

  Logger.log('ลบรายการเรียบร้อย');
}


// ฟังก์ชันดึงรายละเอียดตามหมายเลขใบสั่งเบิก
function getDetailsByInvoice(invoiceNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LogProduct');
  const data = sheet.getDataRange().getValues();  
  
  // ตรวจสอบข้อมูลที่ดึงมา
  Logger.log('Data: ' + JSON.stringify(data));
  
  // กรองข้อมูลตามหมายเลขใบสั่งซื้อ
  const entries = data.filter(row => row[3] === invoiceNumber) // คอลัมน์ 4 คือหมายเลขใบสั่งเบิก
                      .map(row => ({
                        detail: row[7],  // ตรวจสอบคอลัมน์ 8 คือรายละเอียด
                        quantity: row[9] // ตรวจสอบคอลัมน์ 10 คือจำนวน
                      }));

  // ตรวจสอบรายละเอียดที่ดึงมา
  Logger.log('Entries: ' + JSON.stringify(entries));
  
  return entries;
}
function getmyUser(row, lastColumn) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LogProduct');
  const range = sheet.getRange(row, 1, 1, lastColumn);
  const values = range.getValues()[0];

  if (values.length > 0) {
    return {
      id: values[0],          // ID
      status: values[1],      // สถานะ
      date: values[2],        // วันที่
      number: values[3],      // หมายเลขใบสั่งซื้อ
      name: values[4],        // ชื่อ
      dep: values[5],         // กลุ่ม
      materialCode: values[6],// รหัสวัสดุ
      detail: values[7],      // รายการ
      quantity: values[9],    // จำนวน
      unit: values[8]         // หน่วย
    };
  }
  return null;
}


function createPDFForNewEntries() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LogProduct');
  const values = sheet.getDataRange().getValues();
  
  let processedNumbers = {};  // ติดตามเลขที่ที่สร้าง PDF แล้ว
  const existingPDFs = values.reduce((acc, row) => {
    const number = row[3]; // เลขที่ใบสั่งซื้อในคอลัมน์ที่ 4
    const pdfUrl = row[13]; // URL ของ PDF ในคอลัมน์ที่ 14
    if (pdfUrl) {
      acc[number] = true;  // ถ้ามี URL แสดงว่าได้สร้าง PDF แล้ว
    }
    return acc;
  }, {});

  for (let index = 1; index < values.length; index++) {  // เริ่มที่ index 1 เพื่อข้าม header
    const row = values[index];
    const number = row[3];  // เลขที่ใบสั่งซื้อในคอลัมน์ที่ 4
    
    // ตรวจสอบว่าเป็นรายการ "รออนุมัติ" และยังไม่มีการสร้าง PDF
    if (row[1] === 'รออนุมัติ' && !existingPDFs[number]) {
      if (!processedNumbers[number]) {
        processedNumbers[number] = true;  // ทำเครื่องหมายว่าได้สร้าง PDF สำหรับเลขนี้แล้ว
        
        // หาแถวแรกที่มีเลขที่ตรงกัน
        const firstRowIndex = findFirstRowIndexForNumber(number, values);
        
        if (firstRowIndex !== -1) {
          const user = getmyUser(firstRowIndex + 1, sheet.getLastColumn());  // ดึงข้อมูลผู้ใช้จากแถวแรกที่มีเลขที่ตรงกัน
          if (user && user.number) {
            try {
              const pdf = createmyPDF(user);  // สร้าง PDF
              if (pdf) {
                sheet.getRange(firstRowIndex + 1, 13).setValue(pdf.getUrl());  // บันทึก URL ของ PDF ในคอลัมน์ M
                sheet.getRange(firstRowIndex + 1, 14).setValue('สร้างรายงานแล้ว');  // อัปเดตสถานะ
              }
            } catch (e) {
              Logger.log('Error creating PDF for row ' + (firstRowIndex + 1) + ': ' + e.message);
            }
          }
        }
      }
    }
  }

  if (Object.keys(processedNumbers).length === 0) {
    Logger.log('No new entries to create PDF for.');
  }
}

function createmyPDF(user) {
  const folderID = '1HWkj9yYl0-JtTRuMsyOqURWkngXXazyR'; 
  const slidesID = '1HEmpRnWf1dlZsxhK5yJS4i5MSnFscB_MO1En03Zm7Zw'; 
  const slidesTemp = DriveApp.getFileById(slidesID);
  const mainfolder = DriveApp.getFolderById(folderID);
  
  let slidesNew, editNew;
  try {
    slidesNew = slidesTemp.makeCopy(mainfolder);
    editNew = SlidesApp.openById(slidesNew.getId());
  } catch (error) {
    Logger.log('Error copying or opening slides template: ' + error.message);
    return null;
  }
  
  const slides = editNew.getSlides();
  const entries = getDetailsByInvoice(user.number);

  slides.forEach(slide => {
    slide.replaceAllText('{name}', user.name);
    slide.replaceAllText('{number}', user.number);
    slide.replaceAllText('{dep}', user.dep);
    
    // ตรวจสอบค่า user.date ก่อน
    Logger.log('User date before formatting: ' + user.date);
    slide.replaceAllText('{date}', formatDateTime(user.date));

    let sequenceNumber = 1;
    entries.forEach((entry, index) => {
      let detailPlaceholder = `{detail_${index + 1}}`;
      let quantityPlaceholder = `{quantity_${index + 1}}`;
      let numPlaceholder = `{num_${index + 1}}`;
      
      slide.replaceAllText(detailPlaceholder, entry.detail || '');
      slide.replaceAllText(quantityPlaceholder, entry.quantity || '');
      slide.replaceAllText(numPlaceholder, sequenceNumber);
      
      sequenceNumber++;  // เพิ่มลำดับที่
    });

    // ตั้งค่า placeholders ที่เกินจากจำนวนรายการ
    const maxEntries = 10;
    for (let i = entries.length + 1; i <= maxEntries; i++) {
      slide.replaceAllText(`{detail_${i}}`, '');
      slide.replaceAllText(`{quantity_${i}}`, '');
      slide.replaceAllText(`{num_${i}}`, '');
    }
  });

  editNew.saveAndClose();

  try {
    const myBlob = slidesNew.getAs(MimeType.PDF);
    const pdfName = `${user.number}_${user.name}`;  // กำหนดชื่อไฟล์ PDF
    const newPDF = mainfolder.createFile(myBlob).setName(pdfName);  // ใช้ชื่อไฟล์ที่กำหนด
    slidesNew.setTrashed(true);  
    return newPDF;
  } catch (error) {
    Logger.log('Error creating PDF: ' + error.message);
    slidesNew.setTrashed(true);  
    return null;
  }
}

function formatDateTime(dateTime) {
  // ตรวจสอบว่า dateTime ไม่เป็น null หรือ undefined
  if (!dateTime) {
    Logger.log('Invalid date provided: ' + dateTime);
    return 'Invalid Date';  // คืนค่าข้อความถ้าวันที่ไม่ถูกต้อง
  }

  const date = new Date(dateTime);

  // ตรวจสอบว่าค่าที่ได้เป็นวันที่ถูกต้อง
  if (isNaN(date.getTime())) {
    Logger.log('Invalid date provided: ' + dateTime);
    return 'Invalid Date';  // คืนค่าข้อความถ้าวันที่ไม่ถูกต้อง
  }

  // ใช้ฟังก์ชัน convertToBuddhistYear
  const buddhistYear = convertToBuddhistYear(date.getFullYear());

  const day = String(date.getDate()).padStart(2, '0');
  const monthNames = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
  ];
  const month = monthNames[date.getMonth()];
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');

  // ส่งคืนวันที่ในรูปแบบที่ต้องการ
  return `${day} ${month} ${buddhistYear} ${hours}:${minutes}:${seconds} น.`;
}

function convertToBuddhistYear(year) {
  return year ;
}

function findFirstRowIndexForNumber(number, values) {
  if (!Array.isArray(values)) {
    Logger.log('Error: values is not an array or undefined');
    return -1;
  }

  Logger.log('Values: ' + JSON.stringify(values));  // ล็อกค่า values เพื่อการตรวจสอบ

  for (let index = 1; index < values.length; index++) {  // เริ่มที่ index 1 เพื่อข้าม header
    if (values[index][3] === number) {  // ค้นหาเลขที่ในคอลัมน์ที่ 4
      return index;  // คืนค่าดัชนีของแถวที่พบ
    }
  }
  return -1;  // หากไม่พบเลขที่ตรงกัน
}
function testFormatDateTime() {
  Logger.log(formatDateTime('2024-09-18T11:34:19'));  // ควรแสดงว่า "18 กันยายน 2567 11:34:19"
}

function testFormatDateTime() {
  const dateStr = '2024-09-18T11:34:19';  // ใช้วันที่ที่คาดหวัง
  const formattedDate = formatDateTime(dateStr);
  Logger.log('Formatted date: ' + formattedDate);
}

function saveUnitsToSheet(unitValues) {
    var ssl = SpreadsheetApp.getActive();
    var dataSheet = ssl.getSheetByName("Select");
    var getLastRow = dataSheet.getLastRow();
    
    // สร้าง ID ใหม่เริ่มต้นจาก 01
    var newIDNumber = getLastRow > 1 ? getLastRow - 1 : 0; // เริ่มต้นที่ 0 หากไม่มีแถวใด
    for (var i = 0; i < unitValues.length; i++) {
        newIDNumber++; // เพิ่ม ID ใหม่
        var newUnitID = ('0' + newIDNumber).slice(-2); // เติมศูนย์ให้เป็นสองหลัก
        
        // บันทึกหน่วยนับลงในแผ่นงาน
        dataSheet.getRange(getLastRow + 1, 1).setValue(newUnitID); // เซลล์ ID
        dataSheet.getRange(getLastRow + 1, 2).setValue(unitValues[i]); // เซลล์หน่วยนับ
        getLastRow++; // อัปเดตแถวล่าสุด
    }

    // อัปเดตหมายเลข ID หลังจากเพิ่มหน่วยนับใหม่
    updateUnitIDs();
    
    return 'success'; // คืนค่าความสำเร็จ
}
// ฟังก์ชันสำหรับอัปเดต ID
function updateUnitIDs() {
    var ssl = SpreadsheetApp.getActive();
    var dataSheet = ssl.getSheetByName("Select");
    var getLastRow = dataSheet.getLastRow();

    // ดึงหมายเลข ID ทั้งหมดจากแถวที่ 2
    var unitid_values = dataSheet.getRange(2, 1, getLastRow - 1, 1).getValues();
    var currentIDs = [];

    for (var i = 0; i < unitid_values.length; i++) {
        if (unitid_values[i][0]) {
            currentIDs.push(unitid_values[i][0]);
        }
    }

    // ลบช่องว่างและอัปเดตหมายเลข ID
    for (var j = 0; j < currentIDs.length; j++) {
        var newUnitID = ('0' + (j + 1)).slice(-2); // สร้าง ID ใหม่จาก 01 ขึ้นไป
        if (currentIDs[j] !== newUnitID) {
            dataSheet.getRange(j + 2, 1).setValue(newUnitID); // อัปเดตหมายเลข ID ในแผ่นงาน
        }
    }
}

function updateDepartment(id, department) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Select 2');
    var range = sheet.getDataRange();
    var values = range.getValues();

    for (var i = 0; i < values.length; i++) {
        if (values[i][0] == id) { // เปรียบเทียบ ID
            sheet.getRange(i + 1, 2).setValue(department); // อัปเดตหน่วยงานในคอลัมน์ B
            break;
        }
    }
}

function onEdit(e) {
  // ตรวจสอบว่า e มีค่าหรือไม่
  if (!e) {
    Logger.log("Event object is undefined");
    return;
  }

  var sheetName = e.source.getActiveSheet().getName();
  
  // ตรวจสอบว่าเป็นชีต "Select 2" หรือไม่
  if (sheetName === "Select 2") {
    Logger.log("Creating dropdown...");
    createDropdown(); // เรียกใช้ฟังก์ชันสร้าง Drop Down
  }
}


function createDropdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheetByName("Users"); // ชื่อชีตที่ต้องการ
  var sourceSheet = ss.getSheetByName("Select 2"); // ชื่อชีตต้นทาง
  
  var range = targetSheet.getRange("E2:E"); // ช่วงที่ต้องการสร้าง Drop Down
  var values = sourceSheet.getRange("B1:B").getValues(); // ดึงค่าจากชีตต้นทาง
  
  // ลบค่าว่างและค่าซ้ำ
  var uniqueValues = [];
  for (var i = 0; i < values.length; i++) {
    var value = values[i][0];
    if (typeof value === 'string' && value.trim() !== '' && uniqueValues.indexOf(value.trim().toLowerCase()) === -1) {
      uniqueValues.push(value.trim().toLowerCase()); // แปลงเป็นตัวเล็ก
    }
  }

  // ดึงค่าปัจจุบันใน dropdown
  var currentValidation = range.getDataValidation();
  var currentValues = currentValidation ? currentValidation.getCriteriaValues()[0] : [];

  // รวมค่าที่มีอยู่กับค่าที่ใหม่
  var allValues = currentValues.concat(uniqueValues.filter(v => currentValues.indexOf(v) === -1));

  // สร้างรายการใหม่ที่รวมเฉพาะค่าที่มีอยู่ใน uniqueValues
  var finalValues = uniqueValues.filter(value => allValues.includes(value));

  // ตั้งค่าการตรวจสอบข้อมูล
  var validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(finalValues)
    .setAllowInvalid(false)
    .build();
  
  // ตั้งค่าการตรวจสอบข้อมูลในช่วงที่กำหนด
  range.setDataValidation(validation);
}


function testFormatDateTime() {
  // กำหนดวันที่ทดสอบ
  const testDateTime = new Date('2024-09-26T10:25:14');  // วันที่และเวลาที่ต้องการทดสอบ

  // เรียกใช้ฟังก์ชัน formatDateTime
  const formattedDate = formatDateTime(testDateTime);

  // แสดงผลลัพธ์ใน Logger
  Logger.log('Formatted Date: ' + formattedDate);
}

function passwordSheet() {
  const pwd = '12345';
  const ui = SpreadsheetApp.getUi();

  while (true) {
    const msgBox = ui.prompt('กรุณากรอกรหัสผ่าน', ui.ButtonSet.OK);
    const button = msgBox.getSelectedButton();
    const input = msgBox.getResponseText();

    if (input === '') {
      continue; // กลับไปเริ่มลูปใหม่ถ้ายังไม่ได้กรอก
    }

    if (button === ui.Button.OK && input === pwd) {
      return; // ออกจากฟังก์ชันถ้ารหัสผ่านถูกต้อง
    }
  }
}


function onOpen() {
  passwordSheet(); // เรียกใช้งานฟังก์ชัน passwordSheet เมื่อเปิดแผ่นงาน
}

