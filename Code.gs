var DB_URL = "https://docs.google.com/spreadsheets/d/1vpdmZGmI0Bkb1dy6Ud1mfGOaQlHg5L-AXKvupUYmLvI/edit";
var ADMIN_EMAIL = "le.nguyen.bh91@gmail.com"; // Email quản trị viên nhận thông báo

function getSS() {
  return SpreadsheetApp.openByUrl(DB_URL);
}

// --- KHỞI TẠO CẤU TRÚC SHEETS ---
function setup() {
  var ss = getSS();
  
  var dataSheet = ss.getSheetByName("Data");
  if (!dataSheet) {
    dataSheet = ss.insertSheet("Data");
    dataSheet.appendRow(["Thời gian đăng ký", "Họ và tên", "Số điện thoại", "Email", "Nền tảng", "Tiêu đề", "Mô tả", "Trạng thái", "Mã Yêu Cầu"]);
    dataSheet.getRange("A1:I1").setFontWeight("bold").setBackground("#da251d").setFontColor("#ffffff");
  } else {
    if(dataSheet.getLastColumn() < 9) {
      dataSheet.getRange(1, 9).setValue("Mã Yêu Cầu").setFontWeight("bold").setBackground("#da251d").setFontColor("#ffffff");
      var data = dataSheet.getDataRange().getValues();
      for(var i=1; i<data.length; i++) {
        dataSheet.getRange(i+1, 9).setValue("YC-" + Math.floor(1000 + Math.random() * 9000));
      }
    }
  }
  dataSheet.getRange("C:C").setNumberFormat('@'); 
  
  var chatSheet = ss.getSheetByName("LichSuChat");
  if (!chatSheet) {
    chatSheet = ss.insertSheet("LichSuChat");
    chatSheet.appendRow(["Mã Yêu Cầu", "Thời gian", "Người gửi", "Vai trò", "Nội dung"]);
    chatSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#da251d").setFontColor("#ffffff");
  }

  var nenTangSheet = ss.getSheetByName("NenTang");
  if (!nenTangSheet) {
    nenTangSheet = ss.insertSheet("NenTang");
    nenTangSheet.appendRow(["Tên nền tảng"]);
    nenTangSheet.getRange("A1").setFontWeight("bold").setBackground("#da251d").setFontColor("#ffffff");
    var defaults = [["Hệ thống Quản lý Văn bản"], ["Kênh Mặt trận số 24/7"], ["Danh bạ điện tử"], ["Phần mềm Kế toán/Hành chính"], ["Hệ thống Bản đồ Tương tác"], ["Khác"]];
    nenTangSheet.getRange(2, 1, defaults.length, 1).setValues(defaults);
  }

  var taiKhoanSheet = ss.getSheetByName("TaiKhoan");
  if (!taiKhoanSheet) {
    taiKhoanSheet = ss.insertSheet("TaiKhoan");
    taiKhoanSheet.appendRow(["Tài khoản", "Mật khẩu"]);
    taiKhoanSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#da251d").setFontColor("#ffffff");
    taiKhoanSheet.appendRow(["admin", "123456"]);
  }

  var counterSheet = ss.getSheetByName("Counter");
  if (!counterSheet) {
    counterSheet = ss.insertSheet("Counter");
    counterSheet.appendRow(["Ngày", "Lượt truy cập"]);
    counterSheet.getRange("A1:B1").setFontWeight("bold");
    counterSheet.appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"), 0]);
  }
}

function updateCounter() {
    var ss = getSS();
    var counterSheet = ss.getSheetByName("Counter");
    if (!counterSheet) return {today: 0, total: 0};
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var data = counterSheet.getDataRange().getDisplayValues();
    var isNewDay = true;
    var totalVisits = 0;
    var todayVisits = 1;
    for (var i = 1; i < data.length; i++) {
        totalVisits += parseInt(data[i][1]) || 0;
        if (data[i][0] == today) {
            isNewDay = false;
            todayVisits = (parseInt(data[i][1]) || 0) + 1;
            counterSheet.getRange(i + 1, 2).setValue(todayVisits);
        }
    }
    if (isNewDay) counterSheet.appendRow([today, 1]);
    totalVisits += isNewDay ? 1 : 1;
    return { today: todayVisits, total: totalVisits };
}

function doGet() {
  setup(); 
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Hệ thống tiếp nhận, hỗ trợ kỹ thuật các nền tảng số - MTTQ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getPublicData() {
  setup();
  var ss = getSS();
  var pSheet = ss.getSheetByName("NenTang");
  var platforms = pSheet.getLastRow() > 1 ? pSheet.getRange(2, 1, pSheet.getLastRow() - 1, 1).getDisplayValues().map(function(r){return r[0];}) : [];
  
  var dataSheet = ss.getSheetByName("Data");
  var sentData = dataSheet.getLastRow() > 1 ? dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 9).getDisplayValues() : [];
  
  var totalYeuCau = sentData.length;
  var daXuLy = sentData.filter(function(row) { return row[7] === "Đã xử lý"; }).length;
  var chuaXuLy = totalYeuCau - daXuLy;
  
  var formattedData = sentData.map(function(row) {
    return {
      thoigian: row[0], hoten: row[1], sdt: row[2], email: row[3],
      nentang: row[4], tieude: row[5], mota: row[6], trangthai: row[7], maYC: row[8]
    };
  }).reverse();
  
  return {
    success: true, platforms: platforms.filter(String), list: formattedData,
    stats: { tongSo: totalYeuCau, daXuLy: daXuLy, chuaXuLy: chuaXuLy }, dbUrl: DB_URL
  };
}

function submitData(formData) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("Data");
    var chatSheet = ss.getSheetByName("LichSuChat");
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var formattedPhone = "'" + formData.sdt;
    var maYC = "YC" + Math.floor(1000 + Math.random() * 9000); 
    
    // Lưu vào Data
    sheet.appendRow([timestamp, formData.hoten, formattedPhone, formData.email, formData.nentang, formData.tieude, formData.mota, "Chờ xử lý", maYC]);
    // Lưu mô tả vào Lịch sử chat
    chatSheet.appendRow([maYC, timestamp, formData.hoten, "user", formData.mota]);
    
    // 1. GỬI EMAIL THÔNG BÁO CHO QUẢN TRỊ VIÊN
    try {
      MailApp.sendEmail({
        to: ADMIN_EMAIL,
        subject: "CÓ YÊU CẦU MỚI: [" + maYC + "] " + formData.hoten,
        htmlBody: "<h3 style='color:#a81912;'>Hệ thống vừa nhận một yêu cầu!</h3><p><b>Mã:</b> " + maYC + "</p><p><b>Người yêu cầu:</b> " + formData.hoten + "</p><p><b>SĐT:</b> " + formData.sdt + "</p><p><b>Nền tảng:</b> " + formData.nentang + "</p><p><b>Tiêu đề:</b> " + formData.tieude + "</p><p><b>Mô tả:</b> " + formData.mota + "</p>"
      });
    } catch(e){}

    // 2. GỬI EMAIL XÁC NHẬN CHO NGƯỜI DÙNG
    if(formData.email && formData.email.trim() !== "") {
      try {
        MailApp.sendEmail({
          to: formData.email,
          subject: "[MTTQ Đồng Nai] Xác nhận gửi yêu cầu thành công - " + maYC,
          htmlBody: "<h3>Kính gửi " + formData.hoten + ",</h3>" +
                    "<p>Hệ thống đã ghi nhận yêu cầu hỗ trợ kỹ thuật của bạn.</p>" +
                    "<p>Mã tra cứu của bạn là: <b style='color:#da251d; font-size:18px;'>" + maYC + "</b></p>" +
                    "<p>Vui lòng sử dụng mã này cùng số điện thoại của bạn tại mục <b>Tra cứu</b> trên trang web để theo dõi tiến độ và nhận phản hồi từ kỹ thuật viên.</p>" +
                    "<hr><p><i>Ủy ban Mặt trận Tổ quốc Việt Nam thành phố Đồng Nai</i></p>"
        });
      } catch(e){}
    }

    return { success: true, message: "Yêu cầu đã gửi! Mã Tra Cứu của bạn là: " + maYC, maYC: maYC };
  } catch (e) {
    return { success: false, message: "Lỗi hệ thống: " + e.message };
  }
}

function verifyAndGetChat(maYC, sdtInput, isAdmin) {
  var ss = getSS();
  var dataSheet = ss.getSheetByName("Data");
  var chatSheet = ss.getSheetByName("LichSuChat");
  
  var data = dataSheet.getDataRange().getDisplayValues();
  var requestInfo = null;
  
  for(var i=1; i<data.length; i++) {
    if(data[i][8] && data[i][8].toString().toUpperCase() === maYC.toUpperCase()) {
      if(!isAdmin && data[i][2].replace(/'/g, "") !== sdtInput) {
        return { success: false, message: "Số điện thoại không khớp với Mã Yêu Cầu này!" };
      }
      requestInfo = {
        thoigian: data[i][0], hoten: data[i][1], sdt: data[i][2], email: data[i][3],
        nentang: data[i][4], tieude: data[i][5], trangthai: data[i][7], maYC: data[i][8]
      };
      break;
    }
  }
  
  if(!requestInfo) return { success: false, message: "Không tìm thấy Mã Yêu Cầu!" };

  var chats = [];
  if(chatSheet.getLastRow() > 1) {
    var cData = chatSheet.getRange(2, 1, chatSheet.getLastRow() - 1, 5).getDisplayValues();
    for(var j=0; j<cData.length; j++) {
      if(cData[j][0] && cData[j][0].toString().toUpperCase() === maYC.toUpperCase()) {
        chats.push({ thoigian: cData[j][1], nguoigui: cData[j][2], vaitro: cData[j][3], noidung: cData[j][4] });
      }
    }
  }
  return { success: true, info: requestInfo, chats: chats };
}

function sendChatMessage(maYC, nguoigui, vaitro, noidung) {
  try {
    var ss = getSS();
    var chatSheet = ss.getSheetByName("LichSuChat");
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    chatSheet.appendRow([maYC, timestamp, nguoigui, vaitro, noidung]);
    
    // Đổi trạng thái nếu user gửi tin
    if(vaitro === "user") {
      changeRequestStatus(maYC, "Chờ xử lý", true);
    } 
    // Nếu Admin gửi tin (Trả kết quả) -> Báo Mail cho User
    else if(vaitro === "admin") {
      var dataSheet = ss.getSheetByName("Data");
      var data = dataSheet.getDataRange().getDisplayValues();
      for(var i=1; i<data.length; i++) {
        if(data[i][8] && data[i][8].toString().toUpperCase() === maYC.toUpperCase()) {
          var userEmail = data[i][3];
          var userName = data[i][1];
          if(userEmail && userEmail.trim() !== "") {
            try {
              MailApp.sendEmail({
                to: userEmail,
                subject: "[MTTQ Đồng Nai] Có phản hồi mới cho yêu cầu " + maYC,
                htmlBody: "<h3>Chào " + userName + ",</h3>" +
                          "<p>Quản trị viên vừa trả lời yêu cầu hỗ trợ mã <b>" + maYC + "</b> của bạn với nội dung:</p>" +
                          "<blockquote style='border-left: 4px solid #da251d; padding-left: 15px; margin-left: 0; color: #333; font-style: italic;'>" + noidung.replace(/\n/g, '<br>') + "</blockquote>" +
                          "<p>Vui lòng truy cập lại trang Tra Cứu trên hệ thống, nhập mã <b>" + maYC + "</b> để xem chi tiết và tiếp tục trao đổi.</p>" +
                          "<hr><p><i>Ủy ban Mặt trận Tổ quốc Việt Nam thành phố Đồng Nai</i></p>"
              });
            } catch(mailErr){}
          }
          break;
        }
      }
    }
    
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function changeRequestStatus(maYC, newStatus, skipMail) {
  var sheet = getSS().getSheetByName("Data");
  var data = sheet.getDataRange().getDisplayValues();
  for(var i=1; i<data.length; i++) {
    if(data[i][8] == maYC || data[i][0] == maYC) { 
      sheet.getRange(i+1, 8).setValue(newStatus); 
      
      // 3. GỬI EMAIL KHI ĐỔI TRẠNG THÁI (Nếu được yêu cầu)
      if(!skipMail) {
        try {
          var statusColor = newStatus === "Đã xử lý" ? "green" : "red";
          MailApp.sendEmail({
            to: data[i][3],
            subject: "[MTTQ] Cập nhật trạng thái Yêu cầu " + data[i][8],
            htmlBody: "<h3>Chào " + data[i][1] + ",</h3>" +
                      "<p>Yêu cầu <b>" + data[i][8] + "</b> của bạn đã được quản trị viên cập nhật sang trạng thái: <b style='color:" + statusColor + ";'>" + newStatus + "</b>.</p>" +
                      "<p>Vui lòng truy cập hệ thống để xem chi tiết.</p>"
          });
        } catch(e){}
      }
      return true; 
    }
  }
  return false;
}

// --- Các hàm bảo trì & quản trị giữ nguyên ---
function checkLogin(username, password) {
  var data = getSS().getSheetByName("TaiKhoan").getDataRange().getDisplayValues();
  for(var i=1; i<data.length; i++) {
    if(data[i][0] == username && data[i][1] == password) return true;
  }
  return false;
}
function getAdminAccounts() {
  var data = getSS().getSheetByName("TaiKhoan").getDataRange().getDisplayValues();
  var accs = [];
  for(var i=1; i<data.length; i++) { accs.push({user: data[i][0], pass: data[i][1]}); }
  return accs;
}
function deleteRequest(ts) {
  var sheet = getSS().getSheetByName("Data");
  var data = sheet.getDataRange().getDisplayValues();
  for(var i=1; i<data.length; i++) {
    if(data[i][0] == ts) { sheet.deleteRow(i+1); return true; }
  }
  return false;
}
function managePlatform(action, oldVal, newVal) {
  var sheet = getSS().getSheetByName("NenTang");
  if(action === 'add') { sheet.appendRow([newVal]); return true; }
  var data = sheet.getDataRange().getDisplayValues();
  for(var i=1; i<data.length; i++) {
    if(data[i][0] == oldVal) {
      if(action === 'edit') { sheet.getRange(i+1, 1).setValue(newVal); return true; }
      if(action === 'delete') { sheet.deleteRow(i+1); return true; }
    }
  }
  return false;
}
function manageAccount(action, oldUser, newUser, newPass) {
  var sheet = getSS().getSheetByName("TaiKhoan");
  if(action === 'add') { sheet.appendRow([newUser, newPass]); return true; }
  var data = sheet.getDataRange().getDisplayValues();
  for(var i=1; i<data.length; i++) {
    if(data[i][0] == oldUser) {
      if(action === 'edit') { sheet.getRange(i+1, 1).setValue(newUser); if(newPass) sheet.getRange(i+1, 2).setValue(newPass); return true; }
      if(action === 'delete') { sheet.deleteRow(i+1); return true; }
    }
  }
  return false;
}