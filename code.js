// [설정] 
var MASTER_SPREADSHEET_ID = "1GEotzTdn2mR2LxbM3lRMa4ij176wXfn6MKWcITWcC2Y"; 
var CONFIRM_SPREADSHEET_ID = "1cZBca0XODl4Ajdm2yXuIJAI7lXP3UHeR0CtnUDQ8iMw"; 
var CONFIRM_SHEET_NAME = "강사확인서"; 
var INSTRUCTOR_SHEET_NAME = "강사정보"; 

// 💡 서명 이미지를 저장할 구글 드라이브 폴더 ID (본인 폴더 ID로 수정해주세요)
var SIGNATURE_FOLDER_ID = "1Vb7T6lfcX-TIyN6sFqRd6TE6EKqr1n30";

function doGet() {
  // return HtmlService.createTemplateFromFile('Index')
  //   .evaluate()
  //   .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // 기존에 있던 것 유지
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0'); // 💡 구글 프레임을 강제로 뚫고 브라우저에 모바일 뷰포트 신호를 직접 주입하는 치트키입니다.
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function searchInstructorData(name, phone) {
  if (!name || !phone) {
    return { success: false, message: "이름과 연락처를 입력해 주세요." };
  }

  var confirmSs;
  try {
    confirmSs = SpreadsheetApp.openById(CONFIRM_SPREADSHEET_ID);
  } catch (e) {
    return { success: false, message: "강사 확인서 파일을 열 수 없습니다." };
  }
  
  var confirmSheet = confirmSs.getSheetByName(CONFIRM_SHEET_NAME);
  if (!confirmSheet) {
    return { success: false, message: "'" + CONFIRM_SHEET_NAME + "' 시트를 찾을 수 없습니다." };
  }

  var targetMonth = "";
  var rangeValues = confirmSheet.getRange("H4:J4").getValues()[0];
  for (var k = 0; k < rangeValues.length; k++) {
    var val = String(rangeValues[k]).replace(/[^0-9]/g, "");
    if (val !== "") { targetMonth = val; break; }
  }
  
  if (!targetMonth) {
    targetMonth = String(confirmSheet.getRange("I4").getValue()).replace(/[^0-9]/g, "");
  }

  var masterSs;
  try {
    masterSs = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  } catch (e) {
    return { success: false, message: "원본 파일을 열 수 없습니다. 권한을 확인해 주세요." };
  }

  var infoSheet = masterSs.getSheetByName(INSTRUCTOR_SHEET_NAME);
  if (!infoSheet) {
    return { success: false, message: "원본 파일에서 '" + INSTRUCTOR_SHEET_NAME + "' 시트를 찾을 수 없습니다." };
  }

  var infoData = infoSheet.getDataRange().getValues();
  var isValid = false;
  var accountInfo = "정보 없음";
  
  var cleanInputPhone = String(phone).replace(/[^0-9]/g, "");
  // 💡 8자리가 아니면 엄격하게 차단
  if (cleanInputPhone.length !== 8) {
    logAccess(name, phone, false, targetMonth, "실패 (전화번호 자릿수 오류)"); 
    return { success: false, message: "연락처는 뒤 8자리로 정확히 입력해 주세요." };
  }
  
  var inputLast8 = cleanInputPhone; // 8자리가 확실하므로 그대로 비교

  for (var i = 1; i < infoData.length; i++) {
    var cVal = String(infoData[i][2] || "").trim(); // C열: 이름
    var dVal = String(infoData[i][3] || "").trim(); // D열: 연락처
    
    // 시트에 저장된 번호도 숫자만 남겨서 뒤 8자리와 정확히 일치하는지 비교
    var cleanStoredPhone = dVal.replace(/[^0-9]/g, "");
    var storedLast8 = cleanStoredPhone.slice(-8);
    
    if (cVal === name && storedLast8 === inputLast8) {
      isValid = true;
      var bankName = infoData[i][6] ? String(infoData[i][6]).trim() : ""; 
      var bankAcc = infoData[i][7] ? String(infoData[i][7]).trim() : "";  
      if (bankName && bankAcc) {
        accountInfo = bankName + " " + bankAcc;
      } else {
        accountInfo = bankAcc || bankName || "정보 없음";
      }
      break;
    }
  }

  for (var i = 1; i < infoData.length; i++) {
    var cVal = String(infoData[i][2] || "").trim(); // C열: 이름
    var dVal = String(infoData[i][3] || "").trim(); // D열: 연락처
    
    if (cVal === name && dVal.replace(/[^0-9]/g, "").slice(-8) === inputLast8) {
      isValid = true;
      var bankName = infoData[i][6] ? String(infoData[i][6]).trim() : ""; 
      var bankAcc = infoData[i][7] ? String(infoData[i][7]).trim() : "";  
      if (bankName && bankAcc) {
        accountInfo = bankName + " " + bankAcc;
      } else {
        accountInfo = bankAcc || bankName || "정보 없음";
      }
      break;
    }
  }
  
  // ❌ 정보 불일치 시 (실패 로그 기록)
  if (!isValid) {
    logAccess(name, phone, false, targetMonth, "실패 (정보불일치)"); 
    return { success: false, message: "입력하신 정보가 일치하지 않습니다." };
  }
  
  var scheduleSheetName = targetMonth + "월 프로그램 일정표";
  var scheduleSheet = masterSs.getSheetByName(scheduleSheetName);
  
  if (!scheduleSheet) {
    return { success: false, message: "원본 파일에서 '" + scheduleSheetName + "' 시트를 찾을 수 없습니다." };
  }
  
  var scheduleData = scheduleSheet.getDataRange().getValues();
  var mainSchedules = []; 
  var subSchedules = [];  
  
  var totalMainHours = 0;
  var totalMainPay = 0;
  var totalSubHours = 0;
  var totalSubOtherPay = 0;
  var totalSubOtherPay = 0;
  var totalSubPay = 0;
  
  // B열부터 시작하는 인덱스 기준 (B=0, C=1, D=2, E=3, F=4, G=5, H=6, I=7)
  for (var j = 1; j < scheduleData.length; j++) {
    var row = scheduleData[j];
    var teacherD = String(row[3] || "").trim(); // D열: 주강사
    var teacherE = String(row[4] || "").trim(); // E열: 보조강사
    
    var valB = row[1];
    if (valB) {
      valB = String(valB).replace(/\((월|화|수|목|금|토|일)\)\s*(오전|오후)/g, "($1)\n$2");
    }
    
    // 1. 주강사 데이터 처리
    if (teacherD === name) {
      var hF = Number(row[5]) || 0; // F열 (주강사 시수)
      var gG = Number(row[7]) || 0; // H열 (주강사 일급)
      
      totalMainHours += hF;
      totalMainPay += gG;
      
      mainSchedules.push({ 
        colB: valB || "", 
        colC: row[2] || "", 
        colF: hF, 
        colG: gG 
      });
    }
    
    // 2. 보조강사 데이터 처리
    if (teacherE === name) {
      var hH = Number(row[5]) || 0; 
      var gG_other = Number(row[6]) || 0; // 💡 G열 (기타시급) 추가!
      var iI = Number(row[8]) || 0; // I열 (보조강사 일급)
      
      
      totalSubHours += hH;
      totalSubOtherPay += gG_other; // 💡 기타시급 합계 누적
      totalSubPay += iI;
      
      subSchedules.push({ 
        colB: valB || "", 
        colC: row[2] || "", 
        colH: hH, 
        colOther: gG_other, // 💡 0 대신 G열 값 연결!
        colI: iI 
      });
    }
  }
  
  // ✅ 이미 제출된 건이 있는지 체크 함수 호출
  var isAlreadySubmitted = checkAlreadySubmitted(confirmSs, name, inputLast8, targetMonth);
  if (isAlreadySubmitted) {
    return { 
      success: false, 
      alreadySubmitted: true, 
      message: targetMonth + "월 강사확인서는 이미 제출이 완료되었습니다. \n수정이 필요한 경우 운영진에게 문의해 주세요." 
    };
  }
  
  // ✅ 단순 조회 성공 시 (성공 로그 기록)
  logAccess(name, phone, true, targetMonth, "성공 (조회완료)");

  return { 
    success: true, 
    month: targetMonth, 
    account: accountInfo,
    mainData: mainSchedules, 
    subData: subSchedules,
    summary: {
      mainHours: totalMainHours,
      mainPay: totalMainPay,
      subHours: totalSubHours,
      subOtherPay: totalSubOtherPay,
      subPay: totalSubPay,
      grandTotal: totalMainPay + totalSubPay
    }
  };
}

// 🆕 완본 PDF 저장 및 시트 업데이트 함수 (saveSignaturePDF)
function saveSignaturePDF(data) {
  try {
    var name = data.name;
    var phone = data.phone;
    var month = data.month;
    var pdfBase64 = data.pdfData;

    // 🛑 [여기에 추가] PDF를 만들거나 드라이브에 저장하기 전 중복 검사!
    var confirmSs = SpreadsheetApp.openById(CONFIRM_SPREADSHEET_ID);
    var last8 = String(phone).replace(/[^0-9]/g, "").slice(-8);
    
    if (checkAlreadySubmitted(confirmSs, name, last8, month)) {
      return { 
        success: false, 
        message: "이미 해당 월의 강사확인서가 제출(서명완료)되었습니다. 중복 제출할 수 없습니다." 
      };
    }

    if (!pdfBase64) {
      return { success: false, message: "PDF 데이터가 전달되지 않았습니다." };
    }

    // 💡 [추가] 파일명에 쓸 폰 뒷 4자리 및 강사번호 가져오기 로직
    var phoneStr = String(phone);
    var phoneLast4 = phoneStr.replace(/[^0-9]/g, "").slice(-4);


    // 1. Base64 PDF 데이터를 구글 드라이브 Blob 파일로 변환
    var base64Content = pdfBase64.split(",")[1];
    var bytes = Utilities.base64Decode(base64Content);
    var blob = Utilities.newBlob(bytes, "application/pdf", month + "월_강사확인서_" + name + phoneLast4 + ".pdf");

    // 2. 구글 드라이브 지정 폴더에 완본 PDF 파일 생성
    var folder = DriveApp.getFolderById(SIGNATURE_FOLDER_ID);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = file.getUrl();

    // 3. '운영진_조회현황' 시트에 완본 PDF 링크 업데이트
    updateOrLogSignature(name, phone, month, fileUrl);

    return { success: true, fileUrl: fileUrl };
  } catch (e) {
    console.error("PDF 저장 오류: " + e.message);
    return { success: false, message: e.message };
  }
}

// 시트 업데이트 (기존 동일 함수 - PDF 링크로 삽입됨)
function updateOrLogSignature(name, phone, month, pdfUrl) {
  try {
    var ss = SpreadsheetApp.openById(CONFIRM_SPREADSHEET_ID);
    var logSheet = ss.getSheetByName("운영진_조회현황");
    if (!logSheet) return;

    var data = logSheet.getDataRange().getValues();
    var targetMonth = month ? month + "월" : "-";
    var linkFormula = '=HYPERLINK("' + pdfUrl + '", "PDF 보기")'; // 💡 시트에 PDF 바로가기로 표시
    var updated = false;

    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][2] == name && data[i][4] == phone && data[i][5] == targetMonth) {
        logSheet.getRange(i + 1, 7).setValue("성공 (서명완료)");
        logSheet.getRange(i + 1, 8).setValue(linkFormula);
        updated = true;
        break;
      }
    }

    if (!updated) {
      var now = new Date();
      var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

      // 💡 핵심: 핸드폰 번호 뒷 4자리 추출 후 이름과 합쳐서 강사번호 생성 (예: 김이름5678)
      var phoneStr = String(phone);
      var phoneLast4 = phoneStr.replace(/[^0-9]/g, "").slice(-4);
      var instructorNo = name + phoneLast4;

      logSheet.appendRow(["", formattedDate, name, instructorNo, phone, targetMonth, "성공 (서명완료)", linkFormula]);
    }
  } catch (e) {
    console.error("시트 업데이트 실패: " + e.message);
  }
}

// -------------------------------------------------------------
// 📊 [신규 추가] 강사 조회 로그 기록 함수 (logAccess)
// -------------------------------------------------------------
function logAccess(name, phone, success, month, customStatus) {
  try {
    var ss = SpreadsheetApp.openById(CONFIRM_SPREADSHEET_ID);
    var logSheet = ss.getSheetByName("운영진_조회현황");
    
    // '운영진_조회현황' 시트가 없으면 헤더와 함께 새로 생성
    if (!logSheet) {
      logSheet = ss.insertSheet("운영진_조회현황");
      logSheet.appendRow(["NO", "조회 일시", "강사명", "강사번호", "연락처(입력)", "조회 월", "조회 결과", "강사확인서"]);
    }
    
    var now = new Date();
    var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var status = customStatus || (success ? "성공 (조회완료)" : "실패 (정보불일치)");

    // 💡 여기도 동일하게 생성
    var phoneStr = String(phone);
    var phoneLast4 = phoneStr.replace(/[^0-9]/g, "").slice(-4);
    var instructorNo = name + phoneLast4;
    
    // 조회 현황 로그 행 추가
    logSheet.appendRow(["", formattedDate, name, instructorNo, phone, (month ? month + "월" : "-"), status, ""]);
  } catch (e) {
    console.error("로그 기록 실패: " + e.message);
  }
}

// 💡 중복 제출 여부를 검사하는 독립 함수
function checkAlreadySubmitted(confirmSs, name, inputLast8, targetMonth) {
  var logSheet = confirmSs.getSheetByName("운영진_조회현황");
  if (!logSheet) return false;

  var logData = logSheet.getDataRange().getValues();
  var targetMonthStr = targetMonth + "월";
  
  for (var l = 1; l < logData.length; l++) {
    var lName = logData[l][2];
    var lPhone = String(logData[l][4]);
    var lMonth = logData[l][5];
    var lStatus = logData[l][6];
    
    if (lName === name && lPhone.replace(/[^0-9]/g, "").slice(-8) === inputLast8 && lMonth === targetMonthStr && lStatus === "성공 (서명완료)") {
      return true; // 이미 제출됨
    }
  }
  return false;
}

// 외부(깃허브 페이지)의 fetch 요청을 받아주는 유일한 진입점 (현관문)
function doPost(e) {
  // var result = {};
  
  // try {
  //   // text/plain으로 넘어온 JSON 문자열 파싱
  //   var data = JSON.parse(e.postData.contents);
  //   var action = data.action;

  //   if (action === 'searchInstructorData') {
  //     result = searchInstructorData(data.name, data.phone); 
  //   } else if (action === 'saveSignaturePDF') {
  //     result = saveSignaturePDF(data);
  //   } else {
  //     result = { success: false, message: "알 수 없는 요청입니다." };
  //   }

  // } catch (err) {
  //   result = { success: false, message: "서버 처리 에러: " + err.toString() };
  // }

  // // ★ 핵심: TextOutput으로 출력하되, CORS 요청을 완벽하게 수용하도록 리턴합니다.
  // return ContentService.createTextOutput(JSON.stringify(result))
  //                      .setMimeType(ContentService.MimeType.JSON);
  var result = { success: true, message: "접속 성공" };
  return ContentService.createTextOutput(JSON.stringify(result))
                       .setMimeType(ContentService.MimeType.JSON);
  
}