function doGet() {
  var template = HtmlService.createTemplateFromFile('index');
  var html = template.evaluate()
    .setTitle('ลงทะเบียนเข้าร่วมงาน "ร้อยร้าน ล้านรัก"');

  // เพิ่มส่วนนี้เพื่อกำหนดสิทธิ์การเข้าถึงเว็บแอป
  var output = HtmlService.createHtmlOutput(html);
  output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1');


  return output;
}


// ฟังก์ชันสำหรับค้นหาข้อมูลพนักงานโดยใช้รหัสพนักงาน
function searchEmployee(employeeCode) {
  var spreadsheetId = '1ZjZFBpHEdpogPyKamXLW8BpYhW3rGzUHrznzoRNwNqI'; // เปลี่ยนเป็น spreadsheet_ID ของคุณ
  var sheetName = 'listdata'; // เปลี่ยนเป็นชื่อชีทที่คุณใช้



  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();



  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === employeeCode) { // หาข้อมูลโดยใช้รหัสพนักงาน
      var name = data[i][1];
      var department = data[i][5];
      var phoneNumber = data[i][6];
      var yearsOfWork = data[i][9];
     
      return {
        name: name,
        department: department,
        phoneNumber: phoneNumber,
        yearsOfWork: yearsOfWork
      };
    }
  }
 
  // หากไม่พบข้อมูล
  return null;
}



// ฟังก์ชันสำหรับเรียกใช้จากเว็บแอป HTML
function getDataMedicamentos(employeeCode) {
  var employeeInfo = searchEmployee(employeeCode);
  return employeeInfo; // ส่งข้อมูลพนักงานกลับไปในรูปแบบ JSON
}



// ฟังก์ชันสำหรับบันทึกข้อมูลลงใน Google Sheets
function registerEmployee(employeeCode, name, department, phoneNumber, yearsOfWork, seatNumber) {
  var spreadsheetId = '1ZjZFBpHEdpogPyKamXLW8BpYhW3rGzUHrznzoRNwNqI'; // เปลี่ยนเป็น ID ของ Google Sheets ของคุณ
  var sheetName = 'register'; // ชื่อชีทที่คุณต้องการบันทึกข้อมูล

  // รวม '00' และ employeeCode เข้าด้วยกัน
  var formattedEmployeeCode = "'"  + employeeCode;

  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);

  // ตรวจสอบว่ารหัสพนักงานซ้ำหรือไม่
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] === formattedEmployeeCode) {
      return 'รหัสพนักงานซ้ำ ไม่สามารถบันทึกซ้ำได้';
    }
  }

  // เพิ่มข้อมูลลงในแถวใหม่โดยไม่สร้างลำดับที่นั่งใหม่
  var newRow = [new Date(), formattedEmployeeCode, name, department, phoneNumber, yearsOfWork, seatNumber]; // เพิ่ม timestamp ลงไปก่อนข้อมูล
  sheet.appendRow(newRow);

  return 'บันทึกข้อมูลลงทะเบียนสำเร็จ';
}




// ฟังก์ชันสำหรับตรวจสอบรายชื่อซ้ำ
function checkDuplicate(employeeCode) {
  var sheet = SpreadsheetApp.openById('1ZjZFBpHEdpogPyKamXLW8BpYhW3rGzUHrznzoRNwNqI').getSheetByName('register'); // เปลี่ยนเป็น ID และชื่อชีตของคุณ
  var data = sheet.getDataRange().getValues();
  var isDuplicate = false;
  var employeeData = null;

  // ตรวจสอบว่ารหัสพนักงานซ้ำหรือไม่
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] === employeeCode) { // 1 คือคอลัมน์ที่มีรหัสพนักงาน
      isDuplicate = true;
      employeeData = {
        name: data[i][2],
        department: data[i][3],
        phoneNumber: data[i][4],
        yearsOfWork: data[i][5],
        seatNumber: data[i][6]
      };
      break;
    }
  }

  return { isDuplicate: isDuplicate, employeeData: employeeData };
}





// ตรวจสอบว่ารหัสพนักงานมีอยู่ในชีท List data
function checkEmployeeCodeInListData(employeeCode) {
  var dataSheet = SpreadsheetApp.openById('1ZjZFBpHEdpogPyKamXLW8BpYhW3rGzUHrznzoRNwNqI').getSheetByName('listdata');
  var employeeCodes = dataSheet.getRange('A:A').getValues().flat(); // ดึงข้อมูลรหัสพนักงานจากชีท List data


  return employeeCodes.includes(employeeCode);
}


function checkRegistrationStatus(employeeCode) {
  // ทำการตรวจสอบสถานะการลงทะเบียนของรหัสพนักงานในคอลัมน์ B ของชีท "register" และคืนค่าสถานะ
  var ss = SpreadsheetApp.openById('1ZjZFBpHEdpogPyKamXLW8BpYhW3rGzUHrznzoRNwNqI'); // เปลี่ยน YOUR_SPREADSHEET_ID เป็น ID ของสเปรดชีทของคุณ
  var registerSheet = ss.getSheetByName('register');
  var registerData = registerSheet.getRange(1, 1, registerSheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < registerData.length; i++) {
    if (registerData[i][0] === employeeCode) {
      return 'registered';
    }
  }
  return 'notRegistered';
}


function getAvailableSeat() {
  var spreadsheetId = '1ZjZFBpHEdpogPyKamXLW8BpYhW3rGzUHrznzoRNwNqI'; // แทนที่ด้วย ID ของ Google Sheets ของคุณ
  var sheetName = 'que'; // ชื่อชีท "que"

  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === "") { // ถ้าคอลัมน์ A ว่าง
      var seatNumber = data[i][1]; // ดึงเลขที่นั่งที่ว่างจากคอลัมน์ B
      // หลังจากดึงเลขที่นั่งแล้ว คุณสามารถทำการอัพเดตคอลัมน์ A ให้เป็น "ไม่ว่าง" หรือตามที่คุณต้องการ

      // อัพเดตคอลัมน์ A เป็น "ไม่ว่าง" เมื่อมีคนลงทะเบียน
      sheet.getRange(i + 1, 1).setValue("ไม่ว่าง");

      return seatNumber; // คืนค่าเลขที่นั่งที่ว่าง
    }
  }

  return null; // ถ้าไม่มีที่นั่งว่าง
}

function deleteDuplicateData(employeeCode) {
  var spreadsheetId = '1ZjZFBpHEdpogPyKamXLW8BpYhW3rGzUHrznzoRNwNqI'; // เปลี่ยนเป็น ID ของ Google Sheets ของคุณ
  var sheetName = 'register'; // ชื่อชีท "register" ที่ต้องการใช้งาน

  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === employeeCode) { // เช็ครหัสพนักงาน
      sheet.deleteRow(i + 1); // +1 เพราะแถวแรกเป็นหัวข้อ
    }
  }
}




