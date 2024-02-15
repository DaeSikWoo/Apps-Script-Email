// 메일 발송 팝업창
function popUp(){
  var htmlOutput = HtmlService.createHtmlOutputFromFile('popup')
      .setWidth(300)
      .setHeight(180);
  SpreadsheetApp.getActiveSpreadsheet().show(htmlOutput);
}

// 이메일 정보
function emailInfo(emails, header){
  // 각 이메일 주소로 이메일을 보냄
  for (var email in emails) {
    // 이메일 제목
    var subject = "세미나 정보 확인 요청의 건"; 
    
    // HTML 형식으로 이메일 본문을 작성
    var body = '<html><body>';
    
    // 이메일 내용
    body += '<br><p>이번 세미나에 등록된 정보입니다. 확인을 부탁드리며, 정보 수정이 필요한 경우 아래 테이블에 수정 후 회신해주시기 바랍니다</p></br>';


    // 정보가 있는 테이블을 추가합니다.
    body += '<table style="border-collapse: collapse;">';
    
    // 상단에 헤더 정보를 추가합니다.
    body += '<tr style="background-color: #D3D3D3;">';
    for (var i = 0; i < header.length; i++) {
      body += '<th style="border: 1px solid black;">' + header[i] + '</th>';
    }
    body += '</tr>';

    for (var i = 0; i < emails[email].length; i++) {
      body += '<tr>';
      for (var j = 0; j < emails[email][i].length; j++) {
        body += '<td style="border: 1px solid black;">' + emails[email][i][j] + '</td>';
      }
      body += '</tr>';
    }
    body += '</table>';

    body += '</body></html>';
    
    // 메일 발송
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: "Your email client doesn't support HTML.",
      htmlBody: body
    });
  }
}

// 메일 발송 시간 
var now = new Date();
var timezone = "GMT+9";
var format = "yyyy-MM-dd HH:mm:ss";
var dateTime = Utilities.formatDate(now, timezone, format);

// 전체 이메일 발송
function sendEmailsAll() {
  // 시트 정보
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // 첫 번째 시트를 가져옴
  var data = sheet.getRange("A1:D100").getValues(); // A1부터 D100까지의 데이터를 가져옴
  var header = data[0]; // 헤더 정보
  var emails = {}; // 이메일 정보

  // 각 행을 순회하면서 이메일 주소별로 데이터를 모음
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var email = row[3];
  
    // 이메일 주소가 비어있는 행을 무시
    if (!email) {
      continue;
    }
    
    // 이메일 주소가 이전에 나온 적이 없다면 새로운 배열을 생성
    if (!emails[email]) {
      emails[email] = [];
    }
    
    // 이메일 주소에 해당하는 배열에 데이터를 추가
    emails[email].push(row);
  }
  
  // 이메일 전송
  emailInfo(emails, header);

  // 해당 이메일 주소에 대한 데이터를 순회하며 날짜를 기록합니다.
  for (var i = 1; i < data.length; i++) {
    if (data[i].some(cell => cell !== "")) { // 해당 행에 실제 데이터가 있는 경우만 날짜
      console.log(data[i])
      sheet.getRange("E" + (i + 1)).setValue(dateTime);
    }
  }
}

// 신규 이메일 발송
function sendEmailsNew(){
  // 시트 정보
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // 첫 번째 시트를 가져옴
  var data = sheet.getRange("A1:D100").getValues(); // A1부터 D100까지의 데이터를 가져옴
  var header = data[0]; // 헤더 정보
  var emails = {}; // 이메일 정보

  var data2 = sheet.getRange("A1:E100").getValues(); // A1부터 E100까지의 데이터를 가져옴
  var rowNums = {}; // 이메일 주소가 동일한 행 번호를 저장합니다.

  
  // 각 행을 순회하면서 이메일 주소별로 데이터를 모음
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var row2 = data2[i];
    var email = row[3]; // 수정 가능
    var emailSent = row2[4]; // 수정 가능
    
    // 이메일 주소가 비어있는 행을 무시
    if (!email) {
      continue;
    }
    
    // 이메일 주소가 이전에 나온 적이 없다면 새로운 배열을 생성하고 행 번호를 저장
    if (!emails[email]) {
      emails[email] = [];
      rowNums[email] = [];
    }

    // 이메일이 보내져 있지 않은 경우에만 데이터를 추가
    if (!emailSent) {
      emails[email].push(row);
      rowNums[email].push(i + 1); // 스프레드시트에서의 행 번호를 저장
    }
  }

  // 이메일 전송은 데이터가 있는 경우에만 수행
  for (var email in emails) {
    if (emails[email].length > 0) {
      // 이메일 전송
      emailInfo({[email]:emails[email]}, header);
    }
  }

  // 해당 이메일 주소에 대한 데이터를 순회하며 날짜를 기록
  for (var email in rowNums) {
    for (var i = 0; i < rowNums[email].length; i++) {
      var rowIndex = rowNums[email][i]; // 저장해둔 행 번호를 사용
      sheet.getRange("E" + rowIndex).setValue(dateTime);
    }
  }
}
