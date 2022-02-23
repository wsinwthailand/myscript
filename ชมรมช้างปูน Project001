function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
  .addMetaTag('viewport', 'width=device-width, initial-scale=1')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

var url = 'https://docs.google.com/spreadsheets/d/1i_v9wClwfjDkX-s7zrrBFOEMe_1d1PG6ArrVj44xp8I/edit#gid=973286585'
var sheetname = 'ชีต1'
var folderId = '1tg380en42qarSUwkrNhED1UF-vid8L4C'

function processForm(formdata){
  var ss= SpreadsheetApp.openById('1i_v9wClwfjDkX-s7zrrBFOEMe_1d1PG6ArrVj44xp8I');    // ระบุ id ของ Google Sheet ที่ใช้เก็บข้อมูล
  var sh=ss.getSheetByName('ชีต1')    //ระบุชื่อชีทที่ใช้เก็บข้อมูล  
  var lastrow = sh.getLastRow() + 1
  var no='=IF(B' + lastrow + '<>"",ROW()-1,"")' //สูตรสำหรับใส่ลำดับที่ 
  var superscript = SuperScript.initSuper(url,sheetname)
  var formObject = {}
  formdata.forEach(element => formObject[element.name] = element.value)
  var timestamp = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm:ss')
    var file1 ,file2 ,file3
        if(formObject.myfile1.data){
            file1 = superscript.uploadFile(folderId,formObject.myfile1.data,formObject.myfile1.name)
            pictureURL1 = 'http://drive.google.com/uc?id=' + file1.getId()            
        }
        if(formObject.myfile2.data){
            file2 = superscript.uploadFile(folderId,formObject.myfile2.data,formObject.myfile2.name)
            pictureURL2 = 'http://drive.google.com/uc?id=' + file2.getId()            
        }
        if(formObject.myfile3.data){
            file3 = superscript.uploadFile(folderId,formObject.myfile3.data,formObject.myfile3.name)
            pictureURL3 = 'http://drive.google.com/uc?id=' + file3.getId()            
        }

  var newrow = [
    no,
    timestamp,
    formObject.Mname,
    formObject.Group, 
    "'" + formObject.Mphone,
    formObject.Mvaccine,
    formObject.Mright,
    formObject.Mroom,
    formObject.Fname,
    formObject.Fvaccine,
    formObject.Fright,
    pictureURL1,
    pictureURL2,
    pictureURL3
  ] 
  sh.appendRow(newrow);
  
  //----ตรวจสอบข้อมูลใหม่ ซ้ำกับของเดิมหรือไม่ ถ้าซ้ำให้ลบข้อมูลเดิมทิ้ง
  var name = sh.getRange(sh.getLastRow(),3).getValue();
  var lastRow = sh.getLastRow();
  var i=2;
  var nameChk =sh.getRange(2,3).getValue();     // กำหนดให้ไปเริ่มต้นตรวจสอบที่บรรทัดที่ 2 คอลัมน์ 3 ก่อนเสมอ (คือชื่อและนามสกุล)
    while(name!==nameChk){
      i++;
    var nameChk=sh.getRange(i,3).getValue();  
  }
  if(i<lastRow){
    sh.deleteRow(i)
  }

//-----ส่ง LINE Notify และ อีเมล-----
//--ต้องใช้ Library ชื่อ NotifyApp ---
//--libraryId: 1vXbZfRP-7AqwqV7k0fGAnVjCe34pYyI2WdZBJw1Y8U0_DuEbo5fN32P9
  
  //var token = 'lZvhJSF8TMDz7HAMiKMRVZL7ZxWFX51VUb1FRWjKVYm'     // Token One-on-One พรเทพ
  //var token = '0eP81PeoVoY2tjlGDd7scpWhAKNyXXyujR2jpx435Fg'       // Token กลุ่ม หาดใหญ่-สตูล-ตรัง
  var token = 'lZKGaHYzhsXMKtIWOxs6F9J5424FxCrasWsBEU9FJfl'       // Token กลุ่มเฉพาะกิจข้อมูลทริปหาดใหญ่
  var msgLine = '\nลำดับเหตุการณ์ที่ ' + (lastRow-1) +
      '\nมีผู้อัปโหลดไฟล์ 3 ไฟล์' +
      '\nวัน/เวลา: '+ timestamp +
      '\n\n◼ชื่อ: '+ formObject.Mname +
      '\n◼ผู้ติดตาม: '+ formObject.Fname + 
      '\n◼กลุ่มที่: '+ formObject.Group +
      '\n◼URL สลิป: '+ 
      '\n' + pictureURL1 + 
      '\n◼URL บัตรปชช: '+
      '\n' + pictureURL2 +
      '\n◼URL ฉีดวัคซีน: '+
      '\n' + pictureURL3
      NotifyApp.sendNotify(token,msgLine);    
}

