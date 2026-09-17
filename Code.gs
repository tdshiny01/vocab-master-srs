function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Spaced Repetition Vocab')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

const ACCOUNTS_SHEET = "Accounts";

function getDatabase() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Hash mật khẩu bằng SHA-256 + salt riêng cho từng tài khoản.
// Định dạng lưu trong Sheet: salt$hash
function hashPassword(password, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + String(password),
    Utilities.Charset.UTF_8
  );

  return bytes
    .map(function(byte) {
      const value = byte < 0 ? byte + 256 : byte;
      return ('0' + value.toString(16)).slice(-2);
    })
    .join('');
}

function createPasswordHash(password) {
  const salt = Utilities.getUuid();
  return salt + '$' + hashPassword(password, salt);
}

function verifyPassword(password, storedValue) {
  const stored = String(storedValue || '').trim();
  const separator = stored.indexOf('$');

  if (separator > 0) {
    const salt = stored.substring(0, separator);
    const expectedHash = stored.substring(separator + 1);
    const actualHash = hashPassword(password, salt);
    return actualHash === expectedHash;
  }

  // Tương thích với tài khoản cũ đang lưu mật khẩu dạng plaintext.
  // Khi đăng nhập thành công, login() sẽ tự động nâng cấp sang hash.
  return stored.replace(/^'/, '') === String(password);
}

// Hàm tự động vá lỗi nếu Sheet bị thiếu cột
function ensureColumns(sheet) {
  if (!sheet) return;
  const maxCols = sheet.getMaxColumns();
  if (maxCols < 13) {
    sheet.insertColumnsAfter(maxCols, 13 - maxCols);
  }
}

function register(username, password) {
  try {
    const db = getDatabase();
    if (!db) return { success: false, message: "Lỗi: Không tìm thấy cơ sở dữ liệu." };
    
    let accSheet = db.getSheetByName(ACCOUNTS_SHEET);
    if (!accSheet) {
      accSheet = db.insertSheet(ACCOUNTS_SHEET);
      accSheet.appendRow(["Username", "Password", "Created At"]);
    }
    
    let inputUser = String(username || '').trim();
    let inputPass = String(password || '').trim();

    if (inputUser === "" || inputPass === "") {
      return { success: false, message: "Không được để trống." };
    }

    const data = accSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (
        String(data[i][0]).trim().toLowerCase() ===
        inputUser.toLowerCase()
      ) {
        return {
          success: false,
          message: "Tài khoản đã tồn tại!"
        };
      }
    }
    
    accSheet.appendRow([
      inputUser,
      createPasswordHash(inputPass),
      new Date()
    ]);
    
    if (!db.getSheetByName(inputUser)) {
      let userSheet = db.insertSheet(inputUser);

      ensureColumns(userSheet);

      userSheet.getRange("A1:D1").setValues([
        ["'1970-01-01", 0, 0, ""]
      ]);

      let headers = [
        "STT",
        "Tiếng Anh",
        "Phiên âm",
        "Tiếng Việt",
        "Bộ từ vựng",
        "Ví dụ",
        "Ghi chú",
        "Ngày ôn tập tiếp theo",
        "Level",
        "Số lần ôn tập",
        "Flashcard",
        "MCQ",
        "Written"
      ];

      userSheet
        .getRange("A2:M2")
        .setValues([headers]);

      userSheet
        .getRange("A2:M2")
        .setFontWeight("bold")
        .setBackground("#e8f4f8");
    }

    return {
      success: true,
      message: "Tạo tài khoản thành công! Vui lòng đăng nhập."
    };

  } catch (err) {
    return {
      success: false,
      message: "Lỗi Server: " + err.message
    };
  }
}

function login(username, password) {
  try {
    const db = getDatabase();

    if (!db) {
      return {
        success: false,
        message: "Lỗi: Không kết nối được Database."
      };
    }
    
    const accSheet = db.getSheetByName(ACCOUNTS_SHEET);

    if (!accSheet) {
      return {
        success: false,
        message: "Chưa có tài khoản nào được tạo."
      };
    }
    
    const data = accSheet.getDataRange().getValues();

    let inputUser = String(username || '')
      .trim()
      .toLowerCase();

    let inputPass = String(password || '')
      .trim();

    for (let i = 1; i < data.length; i++) {

      let sheetUser = String(data[i][0] || '')
        .trim()
        .toLowerCase();

      let sheetPass = String(data[i][1] || '')
        .trim();

      if (
        sheetUser === inputUser &&
        verifyPassword(inputPass, sheetPass)
      ) {

        // Tự động nâng cấp tài khoản cũ từ plaintext sang hash sau khi
        // đăng nhập thành công, để không cần bắt người dùng đăng ký lại.
        if (sheetPass.indexOf('$') <= 0) {
          accSheet.getRange(i + 1, 2).setValue(createPasswordHash(inputPass));
        }

        let originalName = String(data[i][0]).trim();

        let uSheet = db.getSheetByName(originalName);

        if (uSheet) {
          ensureColumns(uSheet);
        }

        return {
          success: true,
          username: originalName
        };
      }
    }

    return {
      success: false,
      message: "Sai tên đăng nhập hoặc mật khẩu."
    };

  } catch (err) {
    return {
      success: false,
      message: "Lỗi máy chủ: " + err.message
    };
  }
}

function logStudyActivity(username, addSeconds, didStudyWord) {
  try {
    let user = String(username).trim();

    const sheet = getDatabase().getSheetByName(user);

    if (!sheet) return null;
    
    let stats = sheet.getRange("A1:D1").getValues()[0];

    let rawDate = stats[0];

    let lastStudyDateStr = "1970-01-01";

    if (rawDate instanceof Date) {
      lastStudyDateStr = Utilities.formatDate(
        rawDate,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );

    } else if (rawDate) {

      lastStudyDateStr = String(rawDate)
        .replace(/'/g, '')
        .substring(0, 10);
    }
    
    let streak = parseInt(stats[1]) || 0;

    let totalSeconds = parseFloat(stats[2]) || 0;

    if (
      totalSeconds > 0 &&
      totalSeconds < 10000 &&
      !Number.isInteger(totalSeconds)
    ) {
      totalSeconds = Math.floor(totalSeconds * 3600);
    }
    
    let historyStr = String(stats[3] || "");

    let todayD = new Date();

    let todayStr = Utilities.formatDate(
      todayD,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
    
    totalSeconds += (addSeconds || 0);
    
    let historyMap = {};

    if (historyStr) {

      historyStr.split(",").forEach(item => {

        let parts = item.split(":");

        if (parts.length === 2) {
          historyMap[parts[0]] = parseInt(parts[1]);
        }

        else if (parts.length === 1) {
          historyMap[parts[0]] = 0;
        }

      });
    }

    if (!historyMap[todayStr]) {
      historyMap[todayStr] = 0;
    }

    historyMap[todayStr] += (addSeconds || 0);

    if (
      lastStudyDateStr !== "1970-01-01" &&
      lastStudyDateStr !== todayStr
    ) {

      let partsLast = lastStudyDateStr.split('-');
      let partsToday = todayStr.split('-');

      let dLast = new Date(
        partsLast[0],
        partsLast[1] - 1,
        partsLast[2]
      );

      let dToday = new Date(
        partsToday[0],
        partsToday[1] - 1,
        partsToday[2]
      );

      let diffDays = Math.round(
        Math.abs(dToday - dLast) /
        (1000 * 60 * 60 * 24)
      );

      if (diffDays > 1) {
        streak = 0;
      }
    }

    if (didStudyWord) {

      if (lastStudyDateStr !== todayStr) {

        let partsLast = lastStudyDateStr.split('-');
        let partsToday = todayStr.split('-');

        let dLast = new Date(
          partsLast[0],
          partsLast[1] - 1,
          partsLast[2]
        );

        let dToday = new Date(
          partsToday[0],
          partsToday[1] - 1,
          partsToday[2]
        );

        let diffDays = Math.round(
          Math.abs(dToday - dLast) /
          (1000 * 60 * 60 * 24)
        );
        
        if (diffDays === 1 || streak === 0) {
          streak += 1;
        }

        lastStudyDateStr = todayStr;
      }
    }
    
    let newHistoryStr = Object.keys(historyMap)
      .map(k => `${k}:${historyMap[k]}`)
      .join(",");

    sheet.getRange("A1:D1").setValues([
      [
        "'" + lastStudyDateStr,
        streak,
        totalSeconds,
        newHistoryStr
      ]
    ]);
    
    return {
      streak: streak,
      totalSeconds: totalSeconds,
      historyMap: historyMap
    };

  } catch(e) {
    return null;
  }
}

function getUserStats(username) {
  try {

    let user = String(username).trim();

    const sheet = getDatabase().getSheetByName(user);

    if (!sheet) return null;
    
    const stats = sheet.getRange("A1:D1").getValues()[0];

    let rawDate = stats[0];

    let lastStudyStr = "1970-01-01";

    if (rawDate instanceof Date) {

      lastStudyStr = Utilities.formatDate(
        rawDate,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );

    } else if (rawDate) {

      lastStudyStr = String(rawDate)
        .replace(/'/g, '')
        .substring(0, 10);
    }
    
    let streak = parseInt(stats[1]) || 0;

    let todayStr = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
    
    if (
      lastStudyStr !== "1970-01-01" &&
      lastStudyStr !== todayStr
    ) {

      let partsLast = lastStudyStr.split('-');
      let partsToday = todayStr.split('-');

      let dLast = new Date(
        partsLast[0],
        partsLast[1] - 1,
        partsLast[2]
      );

      let dToday = new Date(
        partsToday[0],
        partsToday[1] - 1,
        partsToday[2]
      );

      let diffDays = Math.round(
        Math.abs(dToday - dLast) /
        (1000 * 60 * 60 * 24)
      );

      if (diffDays > 1) {
        streak = 0;
      }
    }
    
    let totalSeconds = parseFloat(stats[2]) || 0;

    if (
      totalSeconds > 0 &&
      totalSeconds < 10000 &&
      !Number.isInteger(totalSeconds)
    ) {
      totalSeconds = Math.floor(totalSeconds * 3600);
    }
    
    let historyStr = String(stats[3] || "");

    let historyMap = {};

    if (historyStr) {

      historyStr.split(",").forEach(item => {

        let parts = item.split(":");

        if (parts.length === 2) {
          historyMap[parts[0]] = parseInt(parts[1]);
        }

        else if (parts.length === 1) {
          historyMap[parts[0]] = 0;
        }

      });
    }
    
    return {
      lastStudy: lastStudyStr,
      streak: streak,
      totalSeconds: totalSeconds,
      historyMap: historyMap
    };

  } catch(e) {
    return null;
  }
}

function addVocabulary(username, wordData) {
  try {

    let user = String(username).trim();

    const sheet = getDatabase().getSheetByName(user);

    ensureColumns(sheet);

    const lastRow = Math.max(sheet.getLastRow(), 2);

    const stt = lastRow - 1;

    const todayStr = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
    
    let rowData = [

      stt,

      String(wordData[0] || ''),

      String(wordData[1] || ''),

      String(wordData[2] || ''),

      String(wordData[3] || ''),

      String(wordData[4] || ''),

      String(wordData[5] || ''),

      "'" + todayStr,

      0,

      0,

      0,

      0,

      0

    ];

    sheet.appendRow(rowData);

    return true;

  } catch(e) {

    throw new Error(e.message);

  }
}

function getVocabList(username) {
  try {

    let user = String(username).trim();

    const sheet = getDatabase().getSheetByName(user);

    if (!sheet) return [];

    ensureColumns(sheet);

    const lastRow = sheet.getLastRow();

    if (lastRow < 3) return [];
    
    const data = sheet
      .getRange(3, 1, lastRow - 2, 13)
      .getValues();

    return data.map((row, index) => {

      let rawDate = row[7];

      let dStr = "";

      if (rawDate instanceof Date) {

        dStr = Utilities.formatDate(
          rawDate,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd"
        );

      } else {

        dStr = String(rawDate)
          .replace(/'/g, '')
          .trim()
          .substring(0, 10);
      }

      return {

        rowIndex: index + 3,

        stt: row[0],

        eng: String(row[1] || ''),

        phon: String(row[2] || ''),

        vie: String(row[3] || ''),

        set: String(row[4] || ''),

        ex: String(row[5] || ''),

        note: String(row[6] || ''),

        nextDate: dStr,

        level: row[8],

        reviews: row[9],

        flashcard: row[10] || 0,

        mcq: row[11] || 0,

        written: row[12] || 0

      };

    });

  } catch(e) {

    return [];

  }
}

function updateSRS(username, rowIndex, isCorrect) {
  try {

    let user = String(username).trim();

    const sheet = getDatabase().getSheetByName(user);

    ensureColumns(sheet);

    let currentData = sheet
      .getRange(rowIndex, 8, 1, 3)
      .getValues()[0];

    let currentLevel =
      parseInt(currentData[1]) || 0;

    let reviews =
      (parseInt(currentData[2]) || 0) + 1;
    
    let newLevel =
      isCorrect
        ? currentLevel + 1
        : 0;

    let daysToAdd = 0;

    switch (newLevel) {

      case 0:
        daysToAdd = 0;
        break;

      case 1:
        daysToAdd = 1;
        break;

      case 2:
        daysToAdd = 3;
        break;

      case 3:
        daysToAdd = 7;
        break;

      case 4:
        daysToAdd = 14;
        break;

      case 5:
        daysToAdd = 21;
        break;

      case 6:
        daysToAdd = 60;
        break;

      default:
        daysToAdd = 60;
    }
    
    let nextDate = new Date();

    nextDate.setDate(
      nextDate.getDate() + daysToAdd
    );

    let nextDateStr = Utilities.formatDate(
      nextDate,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
    
    sheet
      .getRange(rowIndex, 8, 1, 3)
      .setValues([
        [
          "'" + nextDateStr,
          newLevel,
          reviews
        ]
      ]);

    return newLevel;

  } catch(e) {

    return 0;

  }
}

function markWordDone(username, rowIndex, mode) {
  try {

    let user = String(username).trim();

    const sheet = getDatabase().getSheetByName(user);

    ensureColumns(sheet);

    let col = 11;

    if (mode === 'mcq') {
      col = 12;
    }

    else if (mode === 'written') {
      col = 13;
    }

    sheet
      .getRange(rowIndex, col)
      .setValue(1);

    return true;

  } catch(e) {

    return false;

  }
}

function resetProgress(username, sets, mode) {
  try {

    let user = String(username).trim();

    const sheet = getDatabase().getSheetByName(user);

    ensureColumns(sheet);

    const lastRow = sheet.getLastRow();

    if (lastRow < 3) return true;
    
    let col = 11;

    if (mode === 'mcq') {
      col = 12;
    }

    else if (mode === 'written') {
      col = 13;
    }
    
    const data = sheet
      .getRange(3, 5, lastRow - 2, 1)
      .getValues();
    
    for (let i = 0; i < data.length; i++) {

      let setName = String(data[i][0]);

      if (sets.includes(setName)) {

        sheet
          .getRange(i + 3, col)
          .setValue(0);

      }
    }

    return true;

  } catch(e) {

    return false;

  }
}