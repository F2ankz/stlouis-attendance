var STATUS_OPTIONS = ['present', 'late', 'absent', 'leave'];

function toDateStr(val) {
  if (!val) return '';
  var d = val;
  if (!(val instanceof Date)) {
    var s = String(val).trim();
    if (s.match(/^\d{4}-\d{2}-\d{2}/)) return s.substring(0, 10);
    
    var thMatch = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if(thMatch) {
       var year = parseInt(thMatch[3], 10);
       var month = parseInt(thMatch[2], 10);
       var day = parseInt(thMatch[1], 10);
       if (year > 2500) year -= 543;
       return year + '-' + (month < 10 ? '0'+month : month) + '-' + (day < 10 ? '0'+day : day);
    }
    
    d = new Date(val);
  }
  if (d instanceof Date && !isNaN(d.getTime())) {
    var m = d.getMonth() + 1;
    var day = d.getDate();
    return d.getFullYear() + '-' + (m < 10 ? '0'+m : m) + '-' + (day < 10 ? '0'+day : day);
  }
  return String(val);
}

function toTimeStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var h = val.getHours();
    var m = val.getMinutes();
    return (h < 10 ? '0'+h : h) + ':' + (m < 10 ? '0'+m : m);
  }
  var s = String(val).trim();
  var match = s.match(/(\d{1,2}):(\d{2})/);
  if (match) {
    var hStr = match[1];
    if (hStr.length === 1) hStr = '0' + hStr;
    return hStr + ':' + match[2];
  }
  return s;
}

function doGet(e) {
  var action = e.parameter.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (action === 'readAll') {
    var out = { students: [], schedule: [], attendance: {}, leaveRequests: [], users: [] };
    
    // 1. อ่านข้อมูลนักเรียน
    var shS = ss.getSheetByName('ชื่อนักเรียน');
    if (shS) {
      var dataS = shS.getDataRange().getValues();
      for(var i=1; i<dataS.length; i++) {
        if(dataS[i][0]) {
          out.students.push({
            id: String(dataS[i][0]),
            firstName: String(dataS[i][1] || ''),
            lastName: String(dataS[i][2] || ''),
            nickname: String(dataS[i][3] || ''),
            saint: String(dataS[i][4] || ''),
            qrCode: String(dataS[i][5] || ''),
            status: dataS[i][6] !== undefined && dataS[i][6] !== '' ? Number(dataS[i][6]) : 1
          });
        }
      }
    }
    
    // 2. อ่านตารางเรียน
    var shT = ss.getSheetByName('ตารางเรียน');
    if (shT) {
      var dataT = shT.getDataRange().getValues();
      for(var i=1; i<dataT.length; i++) {
        var dt = toDateStr(dataT[i][1]);
        if(dt !== '') {
          out.schedule.push({
            id: 'p_' + i,
            no: dataT[i][0] || i,
            date: dt,
            topic: String(dataT[i][2] || ''),
            special: String(dataT[i][3] || ''),
            start: String(dataT[i][4] || ''),
            end: String(dataT[i][5] || ''),
            makeup: String(dataT[i][6] || '')
          });
        }
      }
    }
    
    // 3. อ่านเวลามาเรียน
    var shA = ss.getSheetByName('เวลามาเรียน');
    if (shA) {
      var dataA = shA.getDataRange().getValues();
      for(var i=1; i<dataA.length; i++) {
        var sId = String(dataA[i][1] || '').replace(/^'/, '');
        var dt = toDateStr(dataA[i][2]);
        var time = toTimeStr(dataA[i][3]);
        var status = String(dataA[i][4] || '');
        
        var scannedBy = String(dataA[i][5] || '');
        
        if (dt !== '' && sId !== '') {
          if(!out.attendance[dt]) out.attendance[dt] = {};
          out.attendance[dt][sId] = { time: time, status: status, scannedBy: scannedBy };
        }
      }
    }

    // 4. อ่านข้อมูลใบลา (Leave Requests)
    var shL = ss.getSheetByName('ใบลา');
    if (shL) {
      var dataL = shL.getDataRange().getValues();
      for(var i=1; i<dataL.length; i++) {
        var rawId = dataL[i][1];
        var sId = '';
        if (rawId instanceof Date) {
          var y = rawId.getFullYear();
          var m = rawId.getMonth() + 1;
          sId = y + '-' + (m < 10 ? '0'+m : m);
        } else {
          sId = String(rawId || '').replace(/^'/, '');
        }
        var lId = String(dataL[i][0] || ''); // ID ใบลา
        var lDate = toDateStr(dataL[i][2]);  // วันที่ขอลา
        var lReason = String(dataL[i][3] || ''); // เหตุผล
        var lStatus = String(dataL[i][4] || ''); // สถานะ (pending, approved, rejected)
        var lTimestamp = String(dataL[i][5] || ''); // เวลาที่ยื่น
        var lApprovedBy = String(dataL[i][6] || ''); // ผู้อนุมัติ
        var lApprovedAt = String(dataL[i][7] || ''); // เวลาที่อนุมัติ
        if (lId !== '' && i > 0) { // skip header
          out.leaveRequests.push({ id: lId, studentId: sId, date: lDate, reason: lReason, status: lStatus, timestamp: lTimestamp, approvedBy: lApprovedBy, approvedAt: lApprovedAt });
        }
      }
    }
    
    // 5. อ่านข้อมูลผู้ใช้งาน (Users)
    var shU = ss.getSheetByName('ผู้ใช้งาน');
    if (shU) {
      var dataU = shU.getDataRange().getValues();
      for(var i=1; i<dataU.length; i++) {
        var uUsername = String(dataU[i][0] || '').trim();
        var uPassword = String(dataU[i][1] || '').trim();
        var uRole = String(dataU[i][2] || '').trim();
        var uRefId = String(dataU[i][3] || '').trim();
        var uFName = String(dataU[i][4] || '').trim();
        var uLName = String(dataU[i][5] || '').trim();
        var uNName = String(dataU[i][6] || '').trim();
        if (uUsername !== '') {
          out.users.push({
            username: uUsername,
            password: uPassword,
            role: uRole,
            refId: uRefId,
            firstName: uFName,
            lastName: uLName,
            nickname: uNName
          });
        }
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: out }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var postData = JSON.parse(e.postData.contents);
  var action = postData.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (action === 'saveStudents') {
    var shS = ss.getSheetByName('ชื่อนักเรียน');
    if(!shS) shS = ss.insertSheet('ชื่อนักเรียน');
    shS.clear();
    shS.appendRow(['StudentID', 'ชื่อ', 'นามสกุล', 'ชื่อเล่น', 'นักบุญ', 'QR Code', 'Status']);
    var rows = postData.students.map(function(s) {
      return [s.id, s.firstName, s.lastName, s.nickname, s.saint, s.qrCode, s.status !== undefined ? s.status : 1];
    });
    if(rows.length > 0) shS.getRange(2, 1, rows.length, 7).setValues(rows);
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'saveSchedule') {
    var shT = ss.getSheetByName('ตารางเรียน');
    if(!shT) shT = ss.insertSheet('ตารางเรียน');
    shT.clear();
    shT.appendRow(['No', 'วันที่', 'หัวข้อ', 'กิจกรรมพิเศษ', 'เวลาเริ่ม', 'เวลาจบ', 'Makeup']);
    var rows = postData.periods.map(function(p, idx) {
      return [
        idx + 1, 
        p.date ? "'" + p.date : '', 
        p.topic, 
        p.special, 
        p.start ? "'" + p.start : '', 
        p.end ? "'" + p.end : '',
        p.makeup || ''
      ];
    });
    if(rows.length > 0) shT.getRange(2, 1, rows.length, 7).setValues(rows);
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // ----------------------------------------------------
  // บันทึกเวลาเรียน + เก็บข้อมูลผู้ที่ทำการสแกน QR
  // ----------------------------------------------------
  if (action === 'saveRoll') {
    var date = postData.date;
    var roll = postData.roll; 
    var shA = ss.getSheetByName('เวลามาเรียน');
    if(!shA) shA = ss.insertSheet('เวลามาเรียน');
    
    var data = shA.getDataRange().getValues();
    var newRows = [];
    var header = data.length > 0 ? data[0] : ['No', 'StudentID', 'วันที่เข้าเรียน', 'เวลาที่เข้าเรียน', 'สถานะ', 'ผู้เช็คชื่อ'];
    if(header.length < 5) header[4] = 'สถานะ'; 
    if(header.length < 6) header[5] = 'ผู้เช็คชื่อ'; 
    newRows.push(header);
    
    var noCount = 1;
    var existingLeaves = {};
    for(var i=1; i<data.length; i++) {
      var d = toDateStr(data[i][2]);
      if(d === date) {
        var st = String(data[i][4] || '').trim();
        if (st === 'leave') {
          var sid = String(data[i][1] || '').replace(/^'/, '');
          existingLeaves[sid] = { time: toTimeStr(data[i][3]), user: String(data[i][5] || '') };
        }
      } else if(d !== '') { 
        data[i][0] = noCount++;
        data[i][2] = "'" + d;
        if(data[i][3]) data[i][3] = "'" + toTimeStr(data[i][3]);
        if (data[i].length < 6) data[i][5] = '';
        newRows.push(data[i]);
      }
    }
    
    for(var sid in existingLeaves) {
      if (!roll[sid] || !roll[sid].status || roll[sid].status === 'absent') {
        roll[sid] = { status: 'leave', time: existingLeaves[sid].time, scannedBy: existingLeaves[sid].user };
      }
    }
    
    for(var sId in roll) {
      newRows.push([
        noCount++, 
        "'" + sId, 
        "'" + date, 
        roll[sId].time ? "'" + roll[sId].time : '', 
        roll[sId].status,
        roll[sId].scannedBy || '' 
      ]);
    }
    
    shA.clear();
    shA.getRange(1, 1, newRows.length, newRows[0].length).setValues(newRows);
    
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // ----------------------------------------------------
  // ยื่นขอลาเรียน (Leave Request)
  // ----------------------------------------------------
  if (action === 'saveLeaveRequest') {
    var shL = ss.getSheetByName('ใบลา');
    if(!shL) {
      shL = ss.insertSheet('ใบลา');
      shL.appendRow(['LeaveID', 'StudentID', 'วันที่ลา', 'เหตุผล', 'สถานะ', 'เวลาที่ยื่น', 'ผู้อนุมัติ', 'เวลาที่อนุมัติ']);
    }
    var leaveId = 'L_' + new Date().getTime();
    var timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    var date = postData.date; 
    
    shL.appendRow([leaveId, "'" + postData.studentId, "'" + date, postData.reason, 'pending', "'" + timestamp, '', '']);
    
    return ContentService.createTextOutput(JSON.stringify({ success: true, leaveId: leaveId }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'deleteLeaveRequest') {
    var shL = ss.getSheetByName('ใบลา');
    if(!shL) return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'ไม่พบชีตใบลา' })).setMimeType(ContentService.MimeType.JSON);
    
    var data = shL.getDataRange().getValues();
    var success = false;
    for(var i=1; i<data.length; i++) {
      if(String(data[i][0]) === postData.leaveId) {
        shL.deleteRow(i + 1);
        success = true;
        break;
      }
    }
    
    if (success) {
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'ไม่พบข้อมูลใบลาดังกล่าว' })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === 'editLeaveRequest') {
    var shL = ss.getSheetByName('ใบลา');
    if(!shL) return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'ไม่พบชีตใบลา' })).setMimeType(ContentService.MimeType.JSON);
    
    var data = shL.getDataRange().getValues();
    var success = false;
    for(var i=1; i<data.length; i++) {
      if(String(data[i][0]) === postData.leaveId) {
        shL.getRange(i + 1, 3).setValue("'" + postData.date);
        shL.getRange(i + 1, 4).setValue(postData.reason);
        success = true;
        break;
      }
    }
    
    if (success) {
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'ไม่พบข้อมูลใบลาดังกล่าว' })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ----------------------------------------------------
  // อนุมัติการลา (Approve/Reject Leave)
  // ----------------------------------------------------
  if (action === 'updateLeaveStatus') {
    var shL = ss.getSheetByName('ใบลา');
    if(!shL) return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'ไม่พบชีตใบลา' })).setMimeType(ContentService.MimeType.JSON);
    
    var data = shL.getDataRange().getValues();
    var success = false;
    for(var i=1; i<data.length; i++) {
      if(String(data[i][0]) === postData.leaveId) {
        shL.getRange(i+1, 5).setValue(postData.status); // Update status in column E
        if (postData.approvedBy) {
          shL.getRange(i+1, 7).setValue(postData.approvedBy); // Column G
          shL.getRange(i+1, 8).setValue("'" + Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss")); // Column H
        }
        
        // ถ้าอนุมัติ (approved) ให้ไปบันทึกลงชีตเวลามาเรียนอัตโนมัติ
        if (postData.status === 'approved') {
          var rawId = data[i][1];
          var sId = '';
          if (rawId instanceof Date) {
            var y = rawId.getFullYear();
            var m = rawId.getMonth() + 1;
            sId = y + '-' + (m < 10 ? '0'+m : m);
          } else {
            sId = String(rawId || '').replace(/^'/, '');
          }
          var lDate = toDateStr(data[i][2]);
          
          var shA = ss.getSheetByName('เวลามาเรียน');
          if(!shA) {
             shA = ss.insertSheet('เวลามาเรียน');
             shA.appendRow(['No', 'StudentID', 'วันที่เข้าเรียน', 'เวลาที่เข้าเรียน', 'สถานะ', 'ผู้เช็คชื่อ']);
          }
          var dataA = shA.getDataRange().getValues();
          var found = false;
          for(var j=1; j<dataA.length; j++) {
            var existingId = String(dataA[j][1] || '').replace(/^'/, '');
            var targetId = sId.replace(/^'/, '');
            if(existingId === targetId && toDateStr(dataA[j][2]) === lDate) {
              shA.getRange(j+1, 5).setValue('leave'); // update status to leave
              found = true;
              break;
            }
          }
          if(!found) {
             var newNo = dataA.length > 1 ? dataA.length : 1;
             shA.appendRow([newNo, "'" + sId, "'" + lDate, '', 'leave', 'System (Leave)']);
          }
        }
        success = true;
        break;
      }
    }
    
    if(success) {
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'ไม่พบคำร้องใบลา' })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ----------------------------------------------------
  // เปลี่ยนรหัสผ่าน
  // ----------------------------------------------------
  if (action === 'changePassword') {
    var username = postData.username;
    var oldPass = postData.oldPass;
    var newPass = postData.newPass;
    
    var shU = ss.getSheetByName('ผู้ใช้งาน');
    if (!shU) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'ไม่พบฐานข้อมูลผู้ใช้งาน' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var data = shU.getDataRange().getValues();
    var success = false;
    var errorMsg = "ไม่พบผู้ใช้งานนี้";
    
    for (var i = 1; i < data.length; i++) {
      var sheetUser = String(data[i][0]).trim();
      var sheetPass = String(data[i][1]).trim();
      
      var rawNumUser = parseInt(data[i][0], 10);
      var rawNumPass = parseInt(data[i][1], 10);
      
      var userMatch = (sheetUser === username || sheetUser.toLowerCase() === String(username).toLowerCase() || (!isNaN(rawNumUser) && rawNumUser.toString() === username));
      
      if (userMatch) {
        var passMatch = (sheetPass === oldPass || sheetPass.toLowerCase() === String(oldPass).toLowerCase() || (!isNaN(rawNumPass) && rawNumPass.toString() === oldPass));
        
        if (passMatch) {
          shU.getRange(i + 1, 2).setValue(newPass);
          success = true;
        } else {
          errorMsg = "รหัสผ่านเดิมไม่ถูกต้อง";
        }
        break;
      }
    }
    
    if (success) {
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: errorMsg }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ----------------------------------------------------
  // จัดการข้อมูลผู้ใช้งาน
  // ----------------------------------------------------
  if (action === 'saveUser') {
    var uUsername = String(postData.username || '').trim();
    var originalUsername = String(postData.originalUsername || '').trim();
    var searchUsername = originalUsername !== '' ? originalUsername : uUsername;
    var uPassword = String(postData.password || '').trim();
    var uRole = String(postData.role || '').trim();
    var uRefId = String(postData.refId || '').trim();
    var uFName = String(postData.firstName || '').trim();
    var uLName = String(postData.lastName || '').trim();
    var uNName = String(postData.nickname || '').trim();
    
    var shU = ss.getSheetByName('ผู้ใช้งาน');
    if (!shU) {
      shU = ss.insertSheet('ผู้ใช้งาน');
      shU.appendRow(['Username', 'Password', 'Role', 'RefID', 'ชื่อ (E)', 'นามสกุล (F)', 'ชื่อเล่น (G)']);
    }
    
    var dataU = shU.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < dataU.length; i++) {
      if (String(dataU[i][0]).trim() === searchUsername) {
        shU.getRange(i + 1, 1).setValue(uUsername);
        if (uPassword !== '') shU.getRange(i + 1, 2).setValue(uPassword);
        shU.getRange(i + 1, 3).setValue(uRole);
        shU.getRange(i + 1, 4).setValue(uRefId);
        shU.getRange(i + 1, 5).setValue(uFName);
        shU.getRange(i + 1, 6).setValue(uLName);
        shU.getRange(i + 1, 7).setValue(uNName);
        found = true;
        break;
      }
    }
    
    if (!found) {
      shU.appendRow([uUsername, uPassword, uRole, uRefId, uFName, uLName, uNName]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'deleteUser') {
    var uUsername = String(postData.username || '').trim();
    var shU = ss.getSheetByName('ผู้ใช้งาน');
    if (shU) {
      var dataU = shU.getDataRange().getValues();
      for (var i = 1; i < dataU.length; i++) {
        if (String(dataU[i][0]).trim() === uUsername) {
          shU.deleteRow(i + 1);
          break;
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ----------------------------------------------------
  // ล็อกอิน
  // ----------------------------------------------------
  if (action === 'login') {
    var user = String(postData.username || '').trim();
    var pass = String(postData.password || '').trim();
    
    var shU = ss.getSheetByName('ผู้ใช้งาน');
    if(!shU) {
      shU = ss.insertSheet('ผู้ใช้งาน');
      shU.appendRow(['Username', 'Password', 'Role', 'RefID']);
      shU.appendRow(['admin', 'admin123', 'admin', '']);
    }
    
    var data = shU.getDataRange().getValues();
    for(var i=1; i<data.length; i++) {
      var sheetUser = String(data[i][0]).trim();
      var sheetPass = String(data[i][1]).trim();
      
      var rawUser = String(data[i][0]);
      var rawNumUser = parseInt(rawUser, 10);
      var rawPass = String(data[i][1]);
      var rawNumPass = parseInt(rawPass, 10);

      var userMatch = (sheetUser === user || sheetUser.toLowerCase() === user.toLowerCase() || (!isNaN(rawNumUser) && rawNumUser.toString() === user));
      var passMatch = (sheetPass === pass || sheetPass.toLowerCase() === pass.toLowerCase() || (!isNaN(rawNumPass) && rawNumPass.toString() === pass));

      if(userMatch && passMatch) {
        var fName = data[i][4] ? String(data[i][4]).trim() : ''; // คอลัมน์ E (ชื่อ)
        var lName = data[i][5] ? String(data[i][5]).trim() : ''; // คอลัมน์ F (สกุล)
        var nName = data[i][6] ? String(data[i][6]).trim() : ''; // คอลัมน์ G (ชื่อเล่น)
        
        var fullName = fName;
        if (lName) fullName += ' ' + lName;
        if (nName) fullName += ' (' + nName + ')';

        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          role: String(data[i][2]).trim(),
          refId: String(data[i][3]).trim(),
          name: fullName // ส่งชื่อเต็มกลับไปให้ระบบแสดงผล
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'User/Password ไม่ถูกต้อง' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}
