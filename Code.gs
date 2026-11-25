// ==========================================
// PYTHON CODE ACADEMY
// ระบบเชื่อมต่อฐานข้อมูล Google Apps Script (Code.gs)
// รองรับ: Login (รหัสนักเรียน + รหัสผ่าน), Pre-test,
// บทเรียน If-Else, While Loop, For Loop และกิจกรรมแอนิเมชันยานสำรวจ
// ==========================================

const SPREADSHEET_ID = "10Stk8GJvW2k2un2Ymivh7kJy75-pgi9GPZRLBW_x9ks";

const STUDENTS_SHEET = "Students";
const PROGRESS_SHEET = "Progress";
const QUIZ_SHEET = "QuizResults";
const LOGS_SHEET = "ActivityLogs";

const SESSION_TIMEOUT = 21600; // 6 ชั่วโมง

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("PYTHON CODE ACADEMY: Pro Learning Studio")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDatabase() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function hashPassword(password, salt) {
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + password,
    Utilities.Charset.UTF_8
  );

  return rawHash.map(byte => {
    const value = (byte + 256) % 256;
    return value.toString(16).padStart(2, "0");
  }).join("");
}

function safeCompare(a, b) {
  if (typeof a !== "string" || typeof b !== "string") return false;
  if (a.length !== b.length) return false;
  let result = 0;
  for (let i = 0; i < a.length; i++) {
    result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return result === 0;
}

function login(studentId, password, studentName) {
  try {
    studentId = String(studentId || "").trim();
    if (!studentId) return { success: false, message: "กรุณากรอกรหัสผู้ใช้หรือรหัสนักเรียน" };

    // Unified login: ตรวจบัญชีแอดมินก่อน แล้วจึงตรวจบัญชีนักเรียน
    ensureAcademySheets_();
    const adminSheet = getDatabase().getSheetByName(ADMINS_SHEET);
    if (adminSheet) {
      const adminRows = adminSheet.getDataRange().getValues();
      const admin = adminRows.slice(1).find(row => String(row[0]).trim().toLowerCase() === studentId.toLowerCase());
      if (admin) {
        if (!password || studentId.toLowerCase() === "admin") {
          return getAdminDirectSession(admin[0]);
        }
        if (String(admin[3] || "").toUpperCase() !== "ACTIVE") return { success: false, message: "บัญชีผู้ดูแลถูกปิดใช้งาน" };
        const parts = String(admin[2] || "").split(":");
        if (parts.length !== 2 || !safeCompare(hashPassword(password, parts[0]), parts[1])) return { success: false, message: "รหัสผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" };
        const adminToken = Utilities.getUuid() + Utilities.getUuid();
        const session = { role: "admin", adminId: String(admin[0]), studentId: String(admin[0]), studentName: String(admin[1] || "ผู้ดูแลระบบ") };
        CacheService.getScriptCache().put("session_" + adminToken, JSON.stringify(session), SESSION_TIMEOUT);
        CacheService.getScriptCache().put("admin_" + adminToken, JSON.stringify(session), SESSION_TIMEOUT);
        return { success: true, token: adminToken, role: "admin", studentId: String(admin[0]), studentName: String(admin[1] || "ผู้ดูแลระบบ"), message: "เข้าสู่ระบบผู้ดูแลสำเร็จ" };
      }
    }

    // หากเป็นนักเรียน (ส่งชื่อมา หรือไม่ได้ส่งรหัสผ่าน)
    if (studentName || !password) {
      return studentLogin(studentId, studentName || studentId);
    }

    // กรณีมีรหัสผ่านส่งมา (ตรวจสอบบัญชีนักเรียนเดิม)
    const sheet = getDatabase().getSheetByName(STUDENTS_SHEET);
    if (!sheet) throw new Error("ไม่พบชีต Students");

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
      const student = data.find(row => String(row[0]).trim().toLowerCase() === studentId.toLowerCase());
      if (student) {
        const canonicalId = String(student[0]).trim();
        const name = String(student[1] || canonicalId);
        const storedHash = String(student[2] || "");
        if (storedHash) {
          if (storedHash.includes(":")) {
            const parts = storedHash.split(":");
            if (!safeCompare(hashPassword(password, parts[0]), parts[1])) {
              return { success: false, message: "รหัสผ่านไม่ถูกต้อง" };
            }
          } else if (storedHash !== password) {
            return { success: false, message: "รหัสผ่านไม่ถูกต้อง" };
          }
        }
        return studentLogin(canonicalId, name);
      }
    }

    return studentLogin(studentId, studentName || studentId);
  } catch (error) {
    return { success: false, message: "ระบบขัดข้อง: " + error.message };
  }
}

function studentLogin(studentId, studentName) {
  try {
    ensureAcademySheets_();
    studentId = String(studentId || "").trim();
    studentName = String(studentName || "").trim();

    if (!studentId) {
      return { success: false, message: "กรุณากรอกรหัสนักเรียน" };
    }

    // หากพิมพ์คำว่า admin หรือรหัสผู้ดูแลระบบ ให้เข้าสู่ระบบผู้ดูแลทันที (ไม่จำเป็นต้องกรอกชื่อ)
    if (studentId.toLowerCase() === "admin") {
      return getAdminDirectSession("admin");
    }
    const adminSheet = getDatabase().getSheetByName(ADMINS_SHEET);
    if (adminSheet) {
      const adminRows = adminSheet.getDataRange().getValues();
      const adminMatch = adminRows.slice(1).find(r => String(r[0]).trim().toLowerCase() === studentId.toLowerCase());
      if (adminMatch) {
        return getAdminDirectSession(adminMatch[0]);
      }
    }

    if (!studentId || studentId.length < 1 || studentId.length > 40) {
      return { success: false, message: "กรุณากรอกรหัสนักเรียนให้ถูกต้อง (1–40 ตัวอักษร)" };
    }
    if (!studentName || studentName.length < 2 || studentName.length > 80) {
      return { success: false, message: "กรุณากรอกชื่อ-นามสกุลให้ถูกต้อง (อย่างน้อย 2 ตัวอักษร)" };
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const ss = getDatabase();
      const studentSheet = ss.getSheetByName(STUDENTS_SHEET);
      if (!studentSheet) throw new Error("ไม่พบชีต Students");

      const studentVals = studentSheet.getDataRange().getValues();
      const sIdx = studentVals.findIndex((r, i) => i > 0 && String(r[0]).trim().toLowerCase() === studentId.toLowerCase());

      let resolvedName = studentName;
      if (sIdx > 0) {
        const row = studentVals[sIdx];
        const status = String(row[3] || "ACTIVE").trim().toUpperCase();
        if (status !== "ACTIVE" && status !== "") {
          return { success: false, message: "บัญชีนี้ถูกระงับการใช้งาน กรุณาติดต่อผู้สอน" };
        }
        const existingName = String(row[1] || "").trim();
        if (studentName && studentName !== existingName) {
          studentSheet.getRange(sIdx + 1, 2).setValue(studentName);
          resolvedName = studentName;
        } else if (existingName) {
          resolvedName = existingName;
        }
      } else {
        // นักเรียนใหม่: ลงทะเบียนอัตโนมัติ ทุกคนเข้าเรียนได้ทันที
        studentSheet.appendRow([studentId, studentName, "", "ACTIVE"]);
        resolvedName = studentName;
      }

      // ตรวจสอบความก้าวหน้าในชีต Progress
      const progressSheet = ss.getSheetByName(PROGRESS_SHEET);
      let pretestStatus = "NOT_STARTED";
      if (progressSheet) {
        const lastRow = progressSheet.getLastRow();
        let pRow = -1;
        if (lastRow >= 2) {
          const pData = progressSheet.getRange(2, 1, lastRow - 1, 3).getValues();
          const pIdx = pData.findIndex(r => String(r[0]).trim().toLowerCase() === studentId.toLowerCase());
          if (pIdx !== -1) {
            pRow = pIdx + 2;
            pretestStatus = String(pData[pIdx][1] || "NOT_STARTED").trim().toUpperCase();
            const curIfElse = String(pData[pIdx][2] || "").trim().toUpperCase();
            if (!curIfElse || curIfElse === "LOCKED") {
              progressSheet.getRange(pRow, 3).setValue("UNLOCKED");
            }
          }
        }
        if (pRow === -1) {
          progressSheet.appendRow([studentId, "NOT_STARTED", "UNLOCKED", "LOCKED", "LOCKED", "LOCKED", "LOCKED", "LOCKED", "LOCKED", "LOCKED", "LOCKED", "IN_PROGRESS", "LOCKED"]);
          pretestStatus = "NOT_STARTED";
        }
      }

      SpreadsheetApp.flush();

      const token = Utilities.getUuid() + Utilities.getUuid();
      const session = {
        role: "student",
        studentId: studentId,
        studentName: resolvedName
      };
      CacheService.getScriptCache().put("session_" + token, JSON.stringify(session), SESSION_TIMEOUT);

      return {
        success: true,
        token: token,
        studentId: studentId,
        studentName: resolvedName,
        role: "student",
        pretestStatus: pretestStatus,
        message: "เข้าสู่ห้องเรียนสำเร็จ"
      };
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return { success: false, message: "ระบบขัดข้อง: " + error.message };
  }
}

function getMyProfile(token) {
  if (!token || typeof token !== "string") {
    return { success: false, message: "ไม่พบ Session กรุณาเข้าสู่ระบบใหม่" };
  }
  const cache = CacheService.getScriptCache();
  const sessionData = cache.get("session_" + token);
  if (!sessionData) {
    return { success: false, message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่" };
  }
  const session = JSON.parse(sessionData);
  return {
    success: true,
    studentId: session.studentId,
    studentName: session.studentName,
    role: session.role || "student"
  };
}

function logout(token) {
  if (token && typeof token === "string") {
    CacheService.getScriptCache().remove("session_" + token);
    CacheService.getScriptCache().remove("admin_" + token);
  }
  return { success: true, message: "ออกจากระบบสำเร็จ" };
}

function upsertPreTestStatus_(ss, studentId, score, maxScore) {
  ensureAcademySheets_();
  const sh = ss.getSheetByName("PreTestStatus");
  if (!sh) return;
  const rows = sh.getDataRange().getValues();
  const idx = rows.findIndex((r, i) => i > 0 && String(r[0]).trim().toLowerCase() === String(studentId).trim().toLowerCase());
  const oldBest = idx > 0 ? (Number(rows[idx][2]) || 0) : 0;
  const bestScore = idx > 0 ? Math.max(oldBest, score) : score;
  const rec = [studentId, true, bestScore, maxScore, score, new Date()];
  if (idx > 0) sh.getRange(idx + 1, 1, 1, 6).setValues([rec]);
  else sh.appendRow(rec);
}

function getPreTestRecord_(ss, studentId) {
  const sid = String(studentId || "").trim().toLowerCase();
  if (!sid) return { score: null, latestScore: null, maxScore: PRETEST_QUESTIONS.length, completedAt: "" };
  try {
    const sh = ss.getSheetByName("PreTestStatus");
    if (sh && sh.getLastRow() >= 2) {
      const row = sh.getDataRange().getValues().slice(1).find(r => String(r[0]).trim().toLowerCase() === sid);
      if (row) {
        return {
          score: Number(row[2]),
          latestScore: row[4] !== "" && row[4] !== undefined ? Number(row[4]) : Number(row[2]),
          maxScore: Number(row[3]) || PRETEST_QUESTIONS.length,
          completedAt: row[5] ? String(row[5]) : ""
        };
      }
    }
  } catch (e) {}
  try {
    const qs = ss.getSheetByName(QUIZ_SHEET);
    if (qs && qs.getLastRow() >= 2) {
      const matches = qs.getDataRange().getValues().slice(1).filter(r =>
        String(r[1]).trim().toLowerCase() === sid && String(r[2] || "").toLowerCase().includes("pre-test")
      );
      if (matches.length) {
        const last = matches[matches.length - 1];
        const rawScore = Number(last[3]);
        const best = Math.max(...matches.map(m => Number(m[3]) || 0));
        return {
          score: best,
          latestScore: rawScore,
          maxScore: String(last[2] || "").includes("20") ? 20 : 100,
          completedAt: last[4] ? String(last[4]) : ""
        };
      }
    }
  } catch (e) {}
  return { score: null, latestScore: null, maxScore: PRETEST_QUESTIONS.length, completedAt: "" };
}

function getDashboardData(token) {
  const profile = getMyProfile(token);
  if (!profile.success) return profile;

  const ss = getDatabase();
  const progressSheet = ss.getSheetByName(PROGRESS_SHEET);
  if (!progressSheet) throw new Error("ไม่พบชีต Progress");

  const studentId = profile.studentId;
  const lastRow = progressSheet.getLastRow();
  let progress = null;

  if (lastRow >= 2) {
    const data = progressSheet.getRange(2, 1, lastRow - 1, 13).getValues();
    progress = data.find(row => String(row[0]).trim().toLowerCase() === String(studentId).trim().toLowerCase());
  }

  if (!progress) {
    progress = [studentId, "NOT_STARTED", "UNLOCKED", "LOCKED", "LOCKED", "LOCKED", "LOCKED", "LOCKED", "LOCKED", "LOCKED", "LOCKED", "IN_PROGRESS", "LOCKED"];
    progressSheet.appendRow(progress);
  } else {
    const curIfElse = String(progress[2] || "").trim().toUpperCase();
    if (!curIfElse || curIfElse === "LOCKED") {
      progress[2] = "UNLOCKED";
      const rowNum = findStudentProgressRow(progressSheet, studentId);
      if (rowNum > 1) progressSheet.getRange(rowNum, 3).setValue("UNLOCKED");
    }
  }

  const preRec = getPreTestRecord_(ss, studentId);

  return {
    success: true,
    student: {
      studentId: profile.studentId,
      studentName: profile.studentName
    },
    pretestScore: preRec.score,
    pretestLatestScore: preRec.latestScore,
    pretestMaxScore: preRec.maxScore,
    progress: {
      pretestStatus: progress[1],
      pretestScore: preRec.score,
      pretestLatestScore: preRec.latestScore,
      pretestMaxScore: preRec.maxScore,
      ifElseStatus: progress[2],
      whileStatus: progress[3],
      forLoopStatus: progress[4],
      activityStatus: progress[9] || progress[5],
      arenaStatus: progress[9] || "LOCKED",
      finalQuizStatus: progress[12] || "LOCKED"
    }
  };
}

// ==========================================
// PRE-TEST (20 ข้อ เต็ม 20 คะแนน ตามตารางจุดประสงค์การเรียนรู้ 10 ข้อ)
// ==========================================
const PRETEST_QUESTIONS = [
  // หน่วยคำสั่งเงื่อนไข (ข้อ 1 - 10)
  // [จุดประสงค์ 1: อธิบายหลักการทำงานของคำสั่งเงื่อนไขในภาษา Python ได้ (2 ข้อ)]
  { question: "1. หลักการทำงานของคำสั่งเงื่อนไข (Conditional Statement) ในภาษา Python คือข้อใด?", choices: ["วนซ้ำคำสั่งเดิมตามจำนวนรอบที่กำหนด", "ตรวจสอบนิพจน์เงื่อนไขที่เป็น True หรือ False เพื่อเลือกทำงานเฉพาะบล็อกคำสั่งที่ต้องการ", "ใช้สำหรับประกาศตัวแปรเก็บข้อความ", "แปลงข้อมูลตัวเลขให้เป็นทศนิยม"], answer: 1 },
  { question: "2. ในภาษา Python สิ่งใดใช้กำหนดขอบเขตของชุดคำสั่งที่อยู่ภายในบล็อกของคำสั่ง if?", choices: ["เครื่องหมายวงเล็บปีกกา { }", "คำสั่ง begin และ end", "การย่อหน้า (Indentation) หลังเครื่องหมายโคลอน (:)", "เครื่องหมายอัฒภาค (;)"], answer: 2 },

  // [จุดประสงค์ 2: อธิบายความแตกต่างระหว่าง if, if-else และ if-elif-else ได้ (2 ข้อ)]
  { question: "3. ข้อใดอธิบายความแตกต่างระหว่างคำสั่ง if กับ if-else ได้ถูกต้องที่สุด?", choices: ["if ทำงานเฉพาะเมื่อเงื่อนไขเป็นจริง ส่วน if-else มีทางเลือกสำรองให้ทำเมื่อเงื่อนไขเป็นเท็จด้วย", "if ใช้กับตัวเลขเท่านั้น ส่วน if-else ใช้กับข้อความเท่านั้น", "if-else สามารถวนซ้ำคำสั่งได้โดยอัตโนมัติ", "ไม่มีความแตกต่างกัน ใช้แทนกันได้ทุกกรณี"], answer: 0 },
  { question: "4. หากต้องการตรวจสอบเงื่อนไขหลายกรณีต่อเนื่องกันและเลือกทำงานเพียง 1 กรณีที่เป็นจริงกรณีแรก ควรใช้โครงสร้างใด?", choices: ["if อย่างเดียว", "if-else ซ้อน 1 ชั้น", "if-elif-else", "while"], answer: 2 },

  // [จุดประสงค์ 3: เขียนโปรแกรมโดยใช้ if และ if-else ได้ (2 ข้อ)]
  { question: "5. พิจารณาโค้ดต่อไปนี้:\nx = 12\nif x > 10:\n    print('ผ่าน')\nผลลัพธ์ที่ได้คือข้อใด?", choices: ["ผ่าน", "ไม่ผ่าน", "12", "ไม่แสดงผลลัพธ์ใด ๆ"], answer: 0 },
  { question: "6. พิจารณาโค้ดต่อไปนี้:\nscore = 45\nif score >= 50:\n    print('Pass')\nelse:\n    print('Fail')\nผลลัพธ์ที่แสดงออกทางหน้าจอคือข้อใด?", choices: ["Pass", "Fail", "45", "Pass Fail"], answer: 1 },

  // [จุดประสงค์ 4: เขียนโปรแกรมโดยใช้ if-elif-else ได้ (2 ข้อ)]
  { question: "7. พิจารณาโค้ดต่อไปนี้:\ntemp = 28\nif temp >= 35:\n    print('ร้อนมาก')\nelif temp >= 25:\n    print('อุ่น')\nelse:\n    print('เย็น')\nผลลัพธ์ที่ได้คือข้อใด?", choices: ["ร้อนมาก", "อุ่น", "เย็น", "ร้อนมาก อุ่น"], answer: 1 },
  { question: "8. ในโครงสร้าง if-elif-else ข้อใดกล่าวถูกต้องเกี่ยวกับลำดับการประเมินเงื่อนไข?", choices: ["ตรวจสอบจากล่างขึ้นบนเสมอ", "ทำงานทุกบล็อกที่เป็นจริงพร้อมกัน", "ข้ามเงื่อนไขแรกไปตรวจสอบ else ก่อน", "ตรวจสอบจากบนลงล่าง เมื่อพบเงื่อนไขแรกที่เป็นจริงจะทำงานบล็อกนั้นแล้วข้ามบล็อกที่เหลือ"], answer: 3 },

  // [จุดประสงค์ 5: ประยุกต์ใช้คำสั่งเงื่อนไขในการแก้ปัญหาเบื้องต้นได้ (2 ข้อ)]
  { question: "9. ร้านค้ามีส่วนลดเมื่อยอดซื้อ total >= 500 และเป็นสมาชิก (member == True) เงื่อนไขใดเขียนใน Python ได้ถูกต้อง?", choices: ["if total >= 500 or member == True:", "if total >= 500 and member == True:", "if total > 500 not member:", "if total = 500 and member = True:"], answer: 1 },
  { question: "10. หากต้องการเขียนโปรแกรมตรวจสอบว่าจำนวนเต็ม n เป็นเลขคู่หรือไม่ เงื่อนไข if ที่ถูกต้องคือข้อใด?", choices: ["if n % 2 == 0:", "if n / 2 == 0:", "if n % 2 == 1:", "if n == 2:"], answer: 0 },

  // หน่วยคำสั่งวนซ้ำ (ข้อ 11 - 20)
  // [จุดประสงค์ 6: อธิบายหลักการทำงานของการวนซ้ำ (Loop) ได้ (2 ข้อ)]
  { question: "11. ข้อใดคือหลักการทำงานของการวนซ้ำ (Loop) ในภาษา Python?", choices: ["ตรวจสอบชนิดข้อมูลของตัวแปร", "หยุดการทำงานของโปรแกรมทันที", "ทำงานชุดคำสั่งเดิมซ้ำ ๆ อัตโนมัติตามเงื่อนไขหรือจำนวนรอบที่กำหนด", "เรียงลำดับตัวอักษรในข้อความ"], answer: 2 },
  { question: "12. คำสั่งใดใช้สำหรับหยุดออกจาก Loop ทันทีแม้เงื่อนไขหรือรอบการทำงานยังไม่สิ้นสุด?", choices: ["continue", "break", "pass", "stop"], answer: 1 },

  // [จุดประสงค์ 7: อธิบายความแตกต่างระหว่าง While Loop และ For Loop ได้ (2 ข้อ)]
  { question: "13. ความแตกต่างหลักระหว่าง While Loop และ For Loop ในภาษา Python คือข้อใด?", choices: ["While Loop ใช้กับข้อความเท่านั้น ส่วน For Loop ใช้กับตัวเลขเท่านั้น", "While Loop วนซ้ำตราบที่เงื่อนไขเป็นจริง (เหมาะเมื่อไม่ทราบจำนวนรอบแน่นอน) ส่วน For Loop วนซ้ำตามลำดับข้อมูลหรือจำนวนรอบที่ทราบ", "For Loop ไม่สามารถใช้ร่วมกับ if ได้", "While Loop ทำงานเพียงรอบเดียวเสมอ"], answer: 1 },
  { question: "14. สถานการณ์ใดเหมาะสมกับการเลือกใช้ For Loop มากกว่า While Loop?", choices: ["วนรับรหัสผ่านจนกว่าผู้ใช้จะพิมพ์ถูกต้อง", "รออ่านค่าเซนเซอร์จนกว่าอุณหภูมิจะเกิน 40 องศา", "ต้องการแสดงผลตัวเลข 1 ถึง 10 หรือวนอ่านสมาชิกทุกตัวในลิสต์", "รันระบบเกมตลอดเวลาจนกว่าผู้ใช้จะกดออก"], answer: 2 },

  // [จุดประสงค์ 8: เขียนโปรแกรมโดยใช้ While Loop ได้ (2 ข้อ)]
  { question: "15. พิจารณาโค้ดต่อไปนี้:\ni = 1\nwhile i <= 3:\n    print(i, end=' ')\n    i += 1\nผลลัพธ์ที่ได้คือข้อใด?", choices: ["1 2 3", "0 1 2 3", "1 2", "1 2 3 4"], answer: 0 },
  { question: "16. หากเขียนโค้ด:\nx = 1\nwhile x < 5:\n    print(x)\nจะเกิดปัญหาใดขึ้น?", choices: ["โปรแกรมพิมพ์ 1 ถึง 4 แล้วหยุด", "เกิด SyntaxError", "โปรแกรมไม่ทำงานเลย", "เกิด Infinite Loop เพราะไม่มีการปรับค่า x ในลูป"], answer: 3 },

  // [จุดประสงค์ 9: เขียนโปรแกรมโดยใช้ For Loop ได้ (2 ข้อ)]
  { question: "17. คำสั่ง for i in range(4): จะสร้างลำดับค่า i ในแต่ละรอบเป็นตัวเลขใดบ้าง?", choices: ["1, 2, 3, 4", "0, 1, 2, 3", "0, 1, 2, 3, 4", "1, 2, 3"], answer: 1 },
  { question: "18. พิจารณาโค้ดต่อไปนี้:\nfor n in [2, 4, 6]:\n    print(n + 1, end=' ')\nผลลัพธ์ที่ได้คือข้อใด?", choices: ["2 4 6", "3 4 5", "3 5 7", "1 3 5"], answer: 2 },

  // [จุดประสงค์ 10: ประยุกต์ใช้ Loop ในการแก้ปัญหาเบื้องต้นได้ (2 ข้อ)]
  { question: "19. พิจารณาโค้ดต่อไปนี้:\ntotal = 0\nfor i in range(1, 5):\n    total += i\nprint(total)\nค่าที่พิมพ์ออกมาคือเท่าใด?", choices: ["6", "10", "15", "4"], answer: 1 },
  { question: "20. หากต้องการนับจำนวนคะแนนที่ผ่านเกณฑ์ (>= 60) จาก scores = [55, 70, 82, 40] ควรวางคำสั่ง if s >= 60: ไว้ที่ใด?", choices: ["ภายในบล็อกของคำสั่ง for s in scores:", "ก่อนเริ่มต้นคำสั่ง for", "หลังจบคำสั่ง for", "ไว้ในคำสั่ง print()"], answer: 0 }
];

function getPretestQuestions(token) {
  const profile = getMyProfile(token);
  if (!profile || !profile.success) return { success: false, message: "Session หมดอายุ" };
  return {
    success: true,
    total: PRETEST_QUESTIONS.length,
    maxScore: PRETEST_QUESTIONS.length,
    questions: PRETEST_QUESTIONS.map((item, index) => ({ id: index, question: item.question, choices: item.choices }))
  };
}

function submitPretest(token, answers) {
  const profile = getMyProfile(token);
  if (!profile || !profile.success) return { success: false, message: "Session หมดอายุ" };
  if (!Array.isArray(answers) || answers.length !== PRETEST_QUESTIONS.length) {
    return { success: false, message: "กรุณาตอบให้ครบทั้ง " + PRETEST_QUESTIONS.length + " ข้อ" };
  }

  let correct = 0;
  PRETEST_QUESTIONS.forEach((item, index) => {
    if (answers[index] === item.answer) correct++;
  });
  const total = PRETEST_QUESTIONS.length;
  const score = correct; // คะแนนเต็ม 20 (ข้อละ 1 คะแนน)
  const percent = Math.round((correct / total) * 100);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getDatabase();
    const quizSheet = ss.getSheetByName(QUIZ_SHEET);
    const progressSheet = ss.getSheetByName(PROGRESS_SHEET);

    quizSheet.appendRow([Utilities.getUuid(), profile.studentId, "Pre-test (เต็ม 20)", score, new Date()]);
    upsertPreTestStatus_(ss, profile.studentId, score, total);

    const lastRow = progressSheet.getLastRow();
    let studentRow = -1;
    if (lastRow >= 2) {
      const data = progressSheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim().toLowerCase() === String(profile.studentId).trim().toLowerCase()) {
          studentRow = i + 2;
          break;
        }
      }
    }

    if (studentRow === -1) {
      progressSheet.appendRow([profile.studentId, "COMPLETED", "UNLOCKED", "LOCKED", "LOCKED", "LOCKED"]);
    } else {
      progressSheet.getRange(studentRow, 2).setValue("COMPLETED");
      const currentIfElse = String(progressSheet.getRange(studentRow, 3).getValue() || "").toUpperCase();
      if (currentIfElse !== "COMPLETED") progressSheet.getRange(studentRow, 3).setValue("UNLOCKED");
    }
    SpreadsheetApp.flush();

    return {
      success: true,
      score: score,
      maxScore: total,
      percent: percent,
      correct: correct,
      total: total,
      message: "บันทึกคะแนน Pre-test เรียบร้อยแล้ว (" + score + "/" + total + " คะแนน)"
    };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// LESSONS & QUIZZES (ตามตารางจุดประสงค์การเรียนรู้: If-Else 15 ข้อ, While Loop 15 ข้อ, For Loop 15 ข้อ)
// ==========================================
const IFELSE_QUESTIONS = [
  // [จุดประสงค์ 1: อธิบายหลักการทำงานของคำสั่งเงื่อนไขในภาษา Python ได้ (ความเข้าใจ - 2 ข้อ)]
  {
    question: "1. ข้อใดอธิบายหลักการทำงานพื้นฐานของคำสั่งเงื่อนไข (Conditional Statement) ในการเขียนโปรแกรมได้ถูกต้องที่สุด?",
    choices: [
      "ใช้สำหรับบันทึกผลลัพธ์ของโปรแกรมลงฐานข้อมูลโดยอัตโนมัติ",
      "ใช้ตรวจสอบนิพจน์เงื่อนไขทางตรรกะ เพื่อตัดสินใจเลือกเส้นทางการทำงานของชุดคำสั่งตามผลลัพธ์ที่เป็นจริงหรือเท็จ",
      "ใช้สำหรับวนซ้ำการทำงานของชุดคำสั่งเดิมตามจำนวนรอบที่ระบุไว้",
      "ใช้สำหรับแปลงข้อมูลชนิดตัวเลขให้กลายเป็นข้อความก่อนประมวลผล"
    ],
    answer: 1
  },
  {
    question: "2. ในภาษา Python บล็อกคำสั่งภายใต้เงื่อนไข if จะถูกประมวลผลเมื่อใด?",
    choices: [
      "เมื่อผลลัพธ์ของการตรวจสอบเงื่อนไขมีค่าความจริงเป็น True",
      "เมื่อผลลัพธ์ของการตรวจสอบเงื่อนไขมีค่าความจริงเป็น False",
      "เมื่อโปรแกรมทำงานจนจบลูปสุดท้ายแล้วเท่านั้น",
      "โปรแกรมจะประมวลผลเสมอไม่ว่าเงื่อนไขจะเป็นจริงหรือเท็จ"
    ],
    answer: 0
  },

  // [จุดประสงค์ 2: อธิบายความแตกต่างระหว่างคำสั่ง if, if-else และ if-else if ได้ (ความเข้าใจ / วิเคราะห์ - 3 ข้อ)]
  {
    question: "3. ข้อใดอธิบายความแตกต่างระหว่างคำสั่ง if เดี่ยว และคำสั่ง if-else ได้อย่างถูกต้อง?",
    choices: [
      "if เดี่ยวทำงานได้เฉพาะกับตัวเลข ส่วน if-else ทำงานได้เฉพาะกับข้อความ",
      "if เดี่ยวจะทำงานเมื่อเงื่อนไขเป็นจริงเท่านั้น (หากเป็นเท็จจะข้ามไป) ส่วน if-else มีบล็อก else รองรับการทำงานเมื่อเงื่อนไขเป็นเท็จ",
      "if-else ไม่จำเป็นต้องมีการตรวจสอบเงื่อนไขทางตรรกะ",
      "if เดี่ยวสามารถตรวจสอบเงื่อนไขพร้อมกันได้หลายเงื่อนไขมากกว่า if-else"
    ],
    answer: 1
  },
  {
    question: "4. คำสั่ง elif (else if) ในภาษา Python ถูกนำมาใช้เมื่อใดและมีข้อดีอย่างไร?",
    choices: [
      "ใช้เพื่อวนซ้ำคำสั่งเหมือนคำสั่ง while",
      "ใช้เมื่อมีหลายทางเลือก โดยจะตรวจสอบเงื่อนไขถัดไปทีละลำดับเมื่อเงื่อนไขก่อนหน้าเป็นเท็จ",
      "ใช้แทนฟังก์ชัน print() เมื่อต้องการแสดงผลข้อความหลายบรรทัด",
      "ใช้เพื่อบังคับให้ทุกเงื่อนไขทำงานพร้อมกันในเวลาเดียวกัน"
    ],
    answer: 1
  },
  {
    question: "5. หากต้องการเขียนโปรแกรมตัดเกรดเป็น A, B, C, D และ F โครงสร้างคำสั่งเงื่อนไขรูปแบบใดเหมาะสมและมีประสิทธิภาพมากที่สุด?",
    choices: [
      "ใช้ if เดี่ยวแยกกัน 5 คำสั่งโดยไม่มี else",
      "ใช้โครงสร้างแบบหลายทางเลือก if - elif - else",
      "ใช้คำสั่ง while loop วนซ้ำ 5 รอบ",
      "ใช้เฉพาะคำสั่ง else ซ้อนกันโดยไม่ต้องมี if"
    ],
    answer: 1
  },

  // [จุดประสงค์ 3: เขียนโปรแกรมโดยใช้คำสั่ง if และ if-else ได้ (การประยุกต์ใช้ - 4 ข้อ)]
  {
    question: "6. พิจารณาโค้ดต่อไปนี้:\nx = 15\nif x > 10:\n    print('A')\nif x > 20:\n    print('B')\nผลลัพธ์ที่แสดงบนหน้าจอคือข้อใด?",
    choices: [
      "A",
      "B",
      "A และ B ในบรรทัดเดียวกัน",
      "ไม่แสดงข้อความใด ๆ"
    ],
    answer: 0
  },
  {
    question: "7. พิจารณาโค้ดตรวจสอบคะแนนต่อไปนี้:\nscore = 45\nif score >= 50:\n    result = 'Pass'\nelse:\n    result = 'Fail'\nprint(result)\nผลลัพธ์จากการรันโปรแกรมคือข้อใด?",
    choices: [
      "Pass",
      "Fail",
      "Error",
      "score >= 50"
    ],
    answer: 1
  },
  {
    question: "8. หากต้องการตรวจสอบว่าตัวแปร num เป็นจำนวนคู่หรือจำนวนคี่ ข้อใดเขียนคำสั่ง if-else ได้ถูกต้องตามหลักไวยากรณ์ Python?",
    choices: [
      "if num / 2 == 0: print('คู่') else: print('คี่')",
      "if num % 2 == 0:\n    print('คู่')\nelse:\n    print('คี่')",
      "if num % 2 = 0 then print('คู่') else print('คี่')",
      "if (num % 2 == 0) { print('คู่'); } else { print('คี่'); }"
    ],
    answer: 1
  },
  {
    question: "9. พิจารณาโค้ดต่อไปนี้:\nage = 18\nis_member = True\nif age >= 18 and is_member:\n    print('Welcome VIP')\nelse:\n    print('General')\nผลลัพธ์ที่พิมพ์ออกมาคือข้อใด?",
    choices: [
      "General",
      "Welcome VIP",
      "True",
      "เกิดข้อผิดพลาด SyntaxError"
    ],
    answer: 1
  },

  // [จุดประสงค์ 4: เขียนโปรแกรมโดยใช้คำสั่ง if-else if ได้ (การประยุกต์ใช้ - 3 ข้อ)]
  {
    question: "10. พิจารณาโค้ดจำแนกอุณหภูมิต่อไปนี้:\ntemp = 28\nif temp > 35:\n    print('Hot')\nelif temp > 25:\n    print('Warm')\nelif temp > 15:\n    print('Cool')\nelse:\n    print('Cold')\nผลลัพธ์ของโปรแกรมคือข้อใด?",
    choices: [
      "Hot",
      "Warm",
      "Cool",
      "Cold"
    ],
    answer: 1
  },
  {
    question: "11. พิจารณาโค้ดต่อไปนี้:\nnum = 0\nif num > 0:\n    ans = 'Positive'\nelif num < 0:\n    ans = 'Negative'\nelse:\n    ans = 'Zero'\nprint(ans)\nผลลัพธ์ที่ได้คือข้อใด?",
    choices: [
      "Positive",
      "Negative",
      "Zero",
      "None"
    ],
    answer: 2
  },
  {
    question: "12. พิจารณาโค้ดตัดเกรดต่อไปนี้:\nscore = 75\nif score >= 80:\n    grade = 'A'\nelif score >= 70:\n    grade = 'B'\nelif score >= 60:\n    grade = 'C'\nelse:\n    grade = 'F'\nprint(grade)\nตัวแปร grade จะมีค่าตรงกับข้อใด?",
    choices: [
      "A",
      "B",
      "C",
      "F"
    ],
    answer: 1
  },

  // [จุดประสงค์ 5: ประยุกต์ใช้คำสั่งเงื่อนไขในการแก้ปัญหาเบื้องต้นได้ (การประยุกต์ใช้ / วิเคราะห์ - 3 ข้อ)]
  {
    question: "13. โจทย์: ห้างสรรพสินค้าจัดโปรโมชั่น หากยอดซื้อ (total) ตั้งแต่ 1,000 บาทขึ้นไป จะได้รับส่วนลด 10% นอกนั้นไม่ได้รับส่วนลด ข้อใดเขียนโค้ดคำนวณยอดเงินสุทธิ (net) ที่ต้องจ่ายได้ถูกต้อง?",
    choices: [
      "if total >= 1000:\n    net = total * 0.9\nelse:\n    net = total",
      "if total > 1000:\n    net = total * 0.1\nelse:\n    net = 0",
      "if total == 1000:\n    net = total - 100\nelse:\n    net = total",
      "if total <= 1000:\n    net = total * 0.9\nelse:\n    net = total * 1.1"
    ],
    answer: 0
  },
  {
    question: "14. พิจารณาโค้ดคำนวณค่าธรรมเนียมจอดรถ:\nhours = 3\nfee = 0\nif hours <= 1:\n    fee = 0\nelif hours <= 4:\n    fee = (hours - 1) * 20\nelse:\n    fee = 60 + (hours - 4) * 30\nprint(fee)\nค่าธรรมเนียมจอดรถที่แสดงผลออกมาคือเท่าใด?",
    choices: [
      "0",
      "40",
      "60",
      "90"
    ],
    answer: 1
  },
  {
    question: "15. พิจารณาโค้ดตรวจสอบเงื่อนไขความถูกต้องของการเข้าสู่ระบบ (Login Validation):\nusername = 'admin'\npassword = '123'\nif username == 'admin':\n    if password == '1234':\n        status = 'Success'\n    else:\n        status = 'Wrong Password'\nelse:\n    status = 'User Not Found'\nprint(status)\nผลลัพธ์การรันโค้ดดังกล่าวคือข้อใด?",
    choices: [
      "Success",
      "Wrong Password",
      "User Not Found",
      "SyntaxError"
    ],
    answer: 1
  }
];

const WHILE_QUESTIONS = [
  // [จุดประสงค์ 1: เข้าใจหลักการวนซ้ำ (2 ข้อ)]
  {
    question: "1. ข้อใดอธิบายความหมายและประโยชน์หลักของการวนซ้ำ (Loop) ในการเขียนโปรแกรมได้ถูกต้องที่สุด?",
    choices: [
      "ช่วยให้โปรแกรมทำงานเร็วขึ้นโดยไม่ต้องใช้หน่วยความจำ",
      "ช่วยให้คำสั่งทำงานซ้ำ ๆ ตามเงื่อนไขหรือจำนวนรอบที่ต้องการโดยไม่ต้องเขียนโค้ดเดิมซ้ำ",
      "ใช้สำหรับประกาศตัวแปรหลายตัวพร้อมกันในการทำงาน",
      "ใช้สำหรับแสดงผลลัพธ์ออกทางหน้าจอในรูปแบบกราฟิก"
    ],
    answer: 1
  },
  {
    question: "2. ในการทำงานของ While Loop ในภาษา Python โปรแกรมจะตรวจสอบสิ่งใดก่อนเริ่มทำคำสั่งในบล็อกเสมอ?",
    choices: [
      "ตรวจสอบเงื่อนไข (Condition) ว่าเป็นจริง (True) หรือไม่",
      "ตรวจสอบจำนวนรอบที่กำหนดไว้ล่วงหน้าว่าครบหรือยัง",
      "ตรวจสอบว่ามีฟังก์ชัน print() อยู่ในบล็อกหรือไม่",
      "ตรวจสอบชนิดข้อมูลของตัวแปรทุกตัวในโปรแกรม"
    ],
    answer: 0
  },

  // [จุดประสงค์ 2: เข้าใจ/วิเคราะห์ ความแตกต่าง While vs For (2 ข้อ)]
  {
    question: "3. สถานการณ์ใดต่อไปนี้เหมาะสมกับการเลือกใช้ While Loop มากกว่า For Loop?",
    choices: [
      "ต้องการวนพิมพ์ตัวเลขตั้งแต่ 1 ถึง 100",
      "ต้องการวนอ่านค่าข้อมูลสมาชิกทุกคนในลิสต์รายชื่อนักเรียน",
      "การวนรับรหัสผ่านจากผู้ใช้ซ้ำ ๆ จนกว่าจะกรอกถูกต้อง (ไม่ทราบจำนวนรอบที่แน่นอน)",
      "ต้องการพิมพ์ข้อความเดิมซ้ำ 10 ครั้ง"
    ],
    answer: 2
  },
  {
    question: "4. ข้อแตกต่างสำคัญที่สุดระหว่าง While Loop กับ For Loop คือข้อใด?",
    choices: [
      "While Loop ทำงานได้เร็วกว่า For Loop เสมอในทุกกรณี",
      "While Loop ทำงานตามเงื่อนไขตรรกะ เหมาะกับกรณีไม่ทราบจำนวนรอบที่แน่นอน ส่วน For Loop เหมาะกับกรณีที่ทราบจำนวนรอบหรือวนตามลำดับข้อมูล",
      "For Loop ไม่สามารถใช้คำสั่ง break หรือ continue ได้",
      "While Loop ไม่สามารถใช้ร่วมกับตัวแปรตัวเลขจำนวนเต็มได้"
    ],
    answer: 1
  },

  // [จุดประสงค์ 3: ประยุกต์ใช้ เขียนโปรแกรม While Loop (7 ข้อ)]
  {
    question: "5. พิจารณาโค้ดต่อไปนี้:\ncount = 1\nwhile count <= 4:\n    print(count, end=' ')\n    count += 1\nผลลัพธ์ที่ได้คือข้อใด?",
    choices: ["1 2 3", "1 2 3 4", "1 2 3 4 5", "0 1 2 3 4"],
    answer: 1
  },
  {
    question: "6. หากต้องการให้ตัวแปร x เริ่มต้นจาก 10 และลดค่าลงทีละ 2 จนถึง 2 โดยใช้ While Loop ต้องเขียนคำสั่งปรับค่าตัวแปรในลูปอย่างไร?",
    choices: ["x = x + 2", "x -= 2", "x - 2", "x *= 2"],
    answer: 1
  },
  {
    question: "7. พิจารณาโค้ดต่อไปนี้:\nn = 5\nwhile n > 0:\n    n -= 1\n    if n == 2:\n        break\n    print(n, end=' ')\nผลลัพธ์ที่ได้จากการรันคือข้อใด?",
    choices: ["4 3 2 1 0", "4 3 2", "4 3", "5 4 3"],
    answer: 2
  },
  {
    question: "8. หากต้องการวนซ้ำเพื่อรับค่าคำสั่งจากผู้ใช้ไปเรื่อย ๆ จนกว่าจะพิมพ์คำว่า 'exit' ควรเขียนเงื่อนไข while อย่างไรให้ถูกต้อง?",
    choices: [
      "while command == 'exit':",
      "while command != 'exit':",
      "while command in 'exit':",
      "while not (command != 'exit'):"
    ],
    answer: 1
  },
  {
    question: "9. โค้ดข้อใดต่อไปนี้ทำให้เกิด 'ลูปไม่รู้จบ' (Infinite Loop)?",
    choices: [
      "i = 1\nwhile i <= 5:\n    print(i)\n    i += 1",
      "i = 5\nwhile i > 0:\n    print(i)\n    i -= 1",
      "i = 1\nwhile i <= 5:\n    print(i)",
      "i = 10\nwhile i < 5:\n    print(i)\n    i += 1"
    ],
    answer: 2
  },
  {
    question: "10. พิจารณาโค้ดต่อไปนี้:\ni = 0\nwhile i < 5:\n    i += 1\n    if i == 3:\n        continue\n    print(i, end=' ')\nผลลัพธ์คือข้อใด?",
    choices: ["1 2", "1 2 4 5", "1 2 3 4 5", "0 1 2 4 5"],
    answer: 1
  },
  {
    question: "11. ข้อใดเติมในช่องว่าง ______ ได้ถูกต้องเพื่อให้โปรแกรมพิมพ์ผลรวม 1 + 2 + 3 + 4 + 5 = 15?\ntotal = 0\ni = 1\nwhile i <= 5:\n    total += i\n    ______\nprint(total)",
    choices: ["total += 1", "i += 1", "i = total", "i = 5"],
    answer: 1
  },

  // [จุดประสงค์ 5: ประยุกต์ใช้ Loop ในการแก้ปัญหาเบื้องต้น (4 ข้อ)]
  {
    question: "12. (โจทย์สถานการณ์: ตู้ ATM) ระบบต้องการให้ผู้ใช้ป้อนรหัส PIN โดยมีโอกาสป้อนผิดได้ไม่เกิน 3 ครั้ง หากถูกต้องให้หยุดถามทันที โครงสร้างใดเหมาะสมที่สุด?",
    choices: [
      "ใช้ For Loop วน 3 ครั้งและไม่มีการตรวจสอบเงื่อนไข",
      "ใช้ While Loop ร่วมกับตัวนับรอบ attempts < 3 และคำสั่ง break เมื่อรหัสผ่านถูกต้อง",
      "ใช้คำสั่ง if-else ซ้อนกัน 10 ชั้น",
      "ใช้ For Loop แบบไม่กำหนดเงื่อนไขจบการทำงาน"
    ],
    answer: 1
  },
  {
    question: "13. พิจารณาโปรแกรมคำนวณเงินออม:\nsavings = 100\ntarget = 500\nmonths = 0\nwhile savings < target:\n    savings += 100\n    months += 1\nprint(months)\nโปรแกรมนี้จะแสดงค่า months เท่ากับเท่าใด?",
    choices: ["3", "4", "5", "6"],
    answer: 1
  },
  {
    question: "14. (โจทย์วิเคราะห์ข้อผิดพลาด) นักเรียนเขียนโปรแกรมหาผลรวมของเลขคู่ตั้งแต่ 2 ถึง 10 ดังนี้:\ntotal = 0\nx = 2\nwhile x <= 10:\n    total += x\nx += 2\nprint(total)\nโปรแกรมนี้เกิดข้อผิดพลาดใดและจะแก้ไขอย่างไร?",
    choices: [
      "ลูปไม่สิ้นสุด เพราะคำสั่ง x += 2 อยู่นอกบล็อก while ต้องย่อหน้าเข้าไปใน while",
      "คำตอบถูกต้องแล้ว ไม่ต้องแก้ไข",
      "ตัวแปร total ควรมีค่าเริ่มต้นเป็น 2",
      "เงื่อนไขควรเปลี่ยนเป็น x < 10"
    ],
    answer: 0
  },
  {
    question: "15. (โจทย์ประยุกต์ใช้: การทอนเงิน) หากต้องการคำนวณจำนวนเหรียญ 10 บาทที่ทอนจากเงินทอนคงเหลือ change = 48 โดยใช้ while loop:\nchange = 48\ncoin10 = 0\nwhile change >= 10:\n    change -= 10\n    coin10 += 1\nprint(coin10, change)\nผลลัพธ์ที่ได้คือข้อใด?",
    choices: ["4 8", "5 0", "4 0", "3 18"],
    answer: 0
  }
];

const FOR_QUESTIONS = [
  // [จุดประสงค์ 1: เข้าใจหลักการวนซ้ำ (2 ข้อ)]
  {
    question: "1. คำสั่ง For Loop ในภาษา Python มีหลักการทำงานพื้นฐานอย่างไร?",
    choices: [
      "วนซ้ำคำสั่งไปเรื่อย ๆ จนกว่าโปรแกรมจะหน่วยความจำเต็ม",
      "ดึงสมาชิกแต่ละตัวออกจากลำดับข้อมูล (Sequence เช่น range, list, string) มาประมวลผลทีละตัวจนครบ",
      "สุ่มเลือกคำสั่งในบล็อกขึ้นมาทำงานตามเวลาที่กำหนด",
      "แปลงคำสั่ง Python ให้กลายเป็นภาษา C เพื่อเพิ่มความเร็ว"
    ],
    answer: 1
  },
  {
    question: "2. ฟังก์ชัน range() ในภาษา Python ทำหน้าที่อะไรในการทำงานร่วมกับ For Loop?",
    choices: [
      "สุ่มตัวเลข 1 ตัวระหว่างค่าที่กำหนด",
      "สร้างลำดับของตัวเลข (Sequence of numbers) เพื่อใช้กำหนดจำนวนรอบในการวนซ้ำ",
      "ตรวจสอบความถูกต้องของไวยากรณ์ในลูป",
      "คำนวณค่าเฉลี่ยของตัวเลขในลิสต์"
    ],
    answer: 1
  },

  // [จุดประสงค์ 2: เข้าใจ/วิเคราะห์ ความแตกต่าง While vs For (3 ข้อ)]
  {
    question: "3. ในกรณีที่ทราบจำนวนรอบที่แน่นอนของการทำงาน (เช่น ต้องการทำซ้ำ 10 รอบ หรือวนตามสมาชิกในลิสต์ 5 ชิ้น) ข้อใดถูกต้องที่สุด?",
    choices: [
      "ควรใช้ For Loop เพราะเขียนกระชับ ดูแลง่าย และลดความเสี่ยงของการเกิด Infinite Loop",
      "ต้องใช้ While Loop เท่านั้นเพราะ For Loop ไม่รองรับการนับรอบ",
      "ทั้งสองคำสั่งไม่สามารถทำงานที่มีการนับรอบได้",
      "ควรใช้คำสั่ง if-else ต่อกัน 10 ครั้งแทนการวนซ้ำ"
    ],
    answer: 0
  },
  {
    question: "4. พิจารณาการวนซ้ำข้อใดต่อไปนี้ที่ For Loop ทำได้สะดวกและกระชับกว่า While Loop อย่างเห็นได้ชัด?",
    choices: [
      "การวนรอรับค่าเซนเซอร์จนกว่าอุณหภูมิจะเกิน 40 องศา",
      "การวนอ่านตัวอักษรทีละตัวในข้อความ text = 'PYTHON'",
      "การวนรับเงินหยอดเหรียญจนกว่าจะครบตามราคาสินค้า",
      "การรันโปรแกรมระบบเกมตลอดเวลาจนกว่าผู้เล่นจะกดปุ่มออก"
    ],
    answer: 1
  },
  {
    question: "5. หากแปลง For Loop: for i in range(1, 4): print(i) เป็น While Loop ข้อใดเทียบเท่าได้ถูกต้องที่สุด?",
    choices: [
      "i = 1\nwhile i <= 3:\n    print(i)\n    i += 1",
      "i = 1\nwhile i < 3:\n    print(i)",
      "i = 0\nwhile i <= 4:\n    print(i)\n    i += 1",
      "while i == 3:\n    print(i)"
    ],
    answer: 0
  },

  // [จุดประสงค์ 4: เขียนโปรแกรมโดยใช้ For Loop (7 ข้อ)]
  {
    question: "6. คำสั่ง range(2, 9, 3) จะสร้างลำดับตัวเลขชุดใด?",
    choices: [
      "[2, 3, 4, 5, 6, 7, 8, 9]",
      "[2, 5, 8]",
      "[3, 6, 9]",
      "[2, 5, 8, 11]"
    ],
    answer: 1
  },
  {
    question: "7. พิจารณาโค้ดต่อไปนี้:\nfruits = ['Apple', 'Banana', 'Cherry']\nfor f in fruits:\n    if f == 'Banana':\n        continue\n    print(f, end=' ')\nผลลัพธ์ที่ได้จากการรันคือข้อใด?",
    choices: ["Apple Banana Cherry", "Apple", "Apple Cherry", "Banana Cherry"],
    answer: 2
  },
  {
    question: "8. หากต้องการพิมพ์ตัวเลขถอยหลัง: 5 4 3 2 1 ด้วยคำสั่ง For Loop ต้องระบุ range() อย่างไร?",
    choices: ["range(5, 0, -1)", "range(5, 1, -1)", "range(1, 5, -1)", "range(5, -1, 0)"],
    answer: 0
  },
  {
    question: "9. พิจารณาโค้ดซ้อนลูป (Nested Loop) ต่อไปนี้:\ncount = 0\nfor i in range(2):\n    for j in range(3):\n        count += 1\nprint(count)\nผลลัพธ์ของ count คือข้อใด?",
    choices: ["5", "6", "8", "9"],
    answer: 1
  },
  {
    question: "10. โค้ดข้อใดต่อไปนี้ใช้คำสั่ง For Loop นับจำนวนตัวอักษร 'a' ในข้อความ word = 'banana' ได้ถูกต้อง?",
    choices: [
      "c = 0\nfor ch in word:\n    if ch == 'a':\n        c += 1",
      "c = 0\nfor ch in range(word):\n    c += 1",
      "for ch == 'a' in word:\n    c += 1",
      "c = word.for('a')"
    ],
    answer: 0
  },
  {
    question: "11. พิจารณาโค้ดต่อไปนี้:\nsum_val = 0\nfor x in range(1, 6):\n    if x % 2 == 0:\n        sum_val += x\nprint(sum_val)\nผลลัพธ์ที่พิมพ์ออกมาคือเท่าใด?",
    choices: ["6", "9", "12", "15"],
    answer: 0
  },
  {
    question: "12. ผลลัพธ์ของโค้ดต่อไปนี้คือข้อใด?\nfor i in range(1, 6):\n    if i == 4:\n        break\n    print(i, end=' ')",
    choices: ["1 2 3 4 5", "1 2 3 4", "1 2 3", "4 5"],
    answer: 2
  },

  // [จุดประสงค์ 5: ประยุกต์ใช้ Loop ในการแก้ปัญหาเบื้องต้น (3 ข้อ)]
  {
    question: "13. (โจทย์สถานการณ์: การคิดคะแนนและนับผู้ผ่านเกณฑ์) มีคะแนนของนักเรียน scores = [75, 42, 88, 59, 91] ต้องการนับว่ามีนักเรียนที่สอบผ่าน (คะแนน >= 60) กี่คน โค้ดข้อใดทำงานได้ถูกต้อง?",
    choices: [
      "passed = 0\nfor s in scores:\n    if s >= 60:\n        passed += 1\nprint(passed)",
      "for s in scores:\n    passed = 60\nprint(passed)",
      "passed = 0\nfor s in range(60):\n    if s in scores:\n        passed += 1\nprint(passed)",
      "while s in scores:\n    passed += 1"
    ],
    answer: 0
  },
  {
    question: "14. (โจทย์วิเคราะห์: ค้นหาข้อมูลในระบบ) ต้องการค้นหาว่ามีชื่อ 'Somchai' อยู่ในลิสต์ names = ['Alice', 'Bob', 'Somchai', 'David'] หรือไม่ หากพบให้แสดงว่า 'Found' แล้วหยุดค้นหาทันที ควรเขียนโค้ดอย่างไร?",
    choices: [
      "for name in names:\n    if name == 'Somchai':\n        print('Found')\n        break",
      "for name in names:\n    print('Found')\n    continue",
      "while name == 'Somchai':\n    print('Found')",
      "for name in range(names):\n    if name == 'Somchai':\n        print('Found')"
    ],
    answer: 0
  },
  {
    question: "15. (โจทย์ประยุกต์ใช้: คำนวณยอดขายรวมพร้อมส่วนลด) ร้านค้ามีรายการราคาสินค้า prices = [100, 250, 400, 50] หากสินค้าชิ้นใดราคาเกิน 200 บาท จะลดราคาชิ้นนั้น 10% (เหลือราคา 0.9 เท่า) โค้ดข้อใดคำนวณยอดรวมสุทธิได้ถูกต้อง?",
    choices: [
      "total = 0\nfor p in prices:\n    if p > 200:\n        total += p * 0.9\n    else:\n        total += p",
      "total = 0\nfor p in prices:\n    total += p * 0.9",
      "total = 0\nfor p in prices:\n    if p <= 200:\n        total += p * 0.9",
      "while p in prices:\n    total += p"
    ],
    answer: 0
  }
];

const LESSON_CONFIG = {
  ifElse: { title: "If-Else", questions: IFELSE_QUESTIONS, statusColumn: 3, quizColumn: 7, nextColumn: 5 },
  ifelse: { title: "If-Else", questions: IFELSE_QUESTIONS, statusColumn: 3, quizColumn: 7, nextColumn: 5 },
  for: { title: "For Loop", questions: FOR_QUESTIONS, statusColumn: 5, quizColumn: 9, nextColumn: 4 },
  while: { title: "While Loop", questions: WHILE_QUESTIONS, statusColumn: 4, quizColumn: 8, nextColumn: 6 }
};

// ตรวจสอบว่านักเรียนสอบผ่านแบบทดสอบท้ายบท (>= 60%) แล้วหรือยัง
function hasPassedLessonQuiz_(studentId, lessonKey) {
  const normKey = ({ ifelse: "ifElse", ifElse: "ifElse", while: "while", for: "for" })[lessonKey] || lessonKey;
  const config = LESSON_CONFIG[normKey];
  if (!config) return false;

  // 1. ตรวจสอบจาก Progress Sheet (Columns 7, 8, 9)
  try {
    const ps = getDatabase().getSheetByName(PROGRESS_SHEET);
    const row = findStudentProgressRow(ps, studentId);
    if (row > 1 && config.quizColumn) {
      const val = String(ps.getRange(row, config.quizColumn).getValue() || "").toUpperCase();
      if (val === "COMPLETED") return true;
    }
  } catch(e) {}

  // 2. ตรวจสอบจาก QuizResults Sheet (คะแนน >= 9/15 สำหรับชุดเต็ม 15 หรือ >= 60% สำหรับผลเก่า)
  try {
    const qs = getDatabase().getSheetByName(QUIZ_SHEET);
    if (qs && qs.getLastRow() >= 2) {
      const data = qs.getDataRange().getValues().slice(1);
      const title = config.title;
      const minPass = Math.ceil(config.questions.length * 0.6);
      const passed = data.some(r => {
        const matchUser = String(r[1]).trim().toLowerCase() === String(studentId).trim().toLowerCase();
        const quizName = String(r[2] || "");
        const matchLesson = quizName.includes(title) || quizName.toLowerCase().includes(String(lessonKey).toLowerCase());
        if (!matchUser || !matchLesson) return false;
        const val = Number(r[3]) || 0;
        return quizName.includes("(เต็ม 15)") ? (val >= minPass) : (val >= 60);
      });
      if (passed) return true;
    }
  } catch(e) {}

  return false;
}

function findStudentProgressRow(sheet, studentId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const index = ids.findIndex(row => String(row[0]).trim().toLowerCase() === String(studentId).trim().toLowerCase());
  return index === -1 ? -1 : index + 2;
}

function getLesson(token, lessonKey) {
  const profile = getMyProfile(token);
  if (!profile || !profile.success) return { success: false, message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่" };
  const config = LESSON_CONFIG[lessonKey];
  if (!config) return { success: false, message: "ไม่พบบทเรียน: " + lessonKey };

  const challengeKey = String(lessonKey || "").toLowerCase();
  const chSheet = getDatabase().getSheetByName("CodingChallenges");
  let practicalCompletedCount = 0;
  if (chSheet && chSheet.getLastRow() >= 2) {
    const chRows = chSheet.getDataRange().getValues().slice(1).filter(r =>
      String(r[0]) === String(profile.studentId) &&
      String(r[1]).toLowerCase() === challengeKey &&
      (r[3] === true || String(r[3]) === "TRUE")
    );
    practicalCompletedCount = new Set(chRows.map(r => Number(r[2]))).size;
  }
  const practicalTotal = (CODING_TASKS[challengeKey] || []).length || 5;
  const practicalPassed = profile.role === "admin" || practicalCompletedCount >= practicalTotal;
  const quizPassed = profile.role === "admin" || hasPassedLessonQuiz_(profile.studentId, lessonKey);

  return {
    success: true,
    lessonKey: lessonKey,
    practicalPassed: practicalPassed,
    practicalCompletedCount: practicalCompletedCount,
    practicalTotal: practicalTotal,
    quizPassed: quizPassed,
    questions: practicalPassed ? config.questions.map((item, index) => ({ id: index, question: item.question, choices: item.choices })) : [],
    totalQuestions: config.questions.length,
    maxScore: config.questions.length,
    minPassQuestions: Math.ceil(config.questions.length * 0.6),
    passScore: Math.ceil(config.questions.length * 0.6),
    passPercent: 60
  };
}

function submitLesson(token, lessonKey, answers) {
  const profile = getMyProfile(token);
  if (!profile || !profile.success) return { success: false, message: "Session หมดอายุ" };
  const config = LESSON_CONFIG[lessonKey];
  if (!config) return { success: false, message: "ไม่พบบทเรียน" };

  const progressSheetForCheck = getDatabase().getSheetByName(PROGRESS_SHEET);
  const progressRowForCheck = findStudentProgressRow(progressSheetForCheck, profile.studentId);
  if (progressRowForCheck < 2) return { success: false, message: "ไม่พบข้อมูลความก้าวหน้า" };
  const accessStatus = String(progressSheetForCheck.getRange(progressRowForCheck, config.statusColumn).getValue() || "").toUpperCase();
  if (profile.role !== "admin" && config.statusColumn !== 3 && !["UNLOCKED", "COMPLETED"].includes(accessStatus)) return { success: false, message: "ยังไม่ได้ปลดล็อกบทเรียนนี้" };

  const challengeKey = String(lessonKey || "").toLowerCase();
  if (profile.role !== "admin" && !hasPassedAllChallenges_(profile.studentId, challengeKey)) {
    return { success: false, message: "ต้องผ่านแบบทดสอบปฏิบัติครบทั้ง 5 ข้อก่อนจึงจะทำแบบทดสอบท้ายบทเรียนได้" };
  }

  const questions = config.questions;
  if (!Array.isArray(answers) || answers.length !== questions.length || answers.some(a => !Number.isInteger(a) || a < 0 || a > 3)) {
    return { success: false, message: "รูปแบบคำตอบไม่ถูกต้อง กรุณาตอบให้ครบทุกข้อ (มีทั้งหมด " + questions.length + " ข้อ)" };
  }

  let correct = 0;
  answers.forEach((ans, idx) => {
    if (ans === questions[idx].answer) correct++;
  });
  const total = questions.length;
  const score = correct; // คะแนนเต็ม 15 (ข้อละ 1 คะแนน)
  const percent = Math.round((correct / total) * 100);
  const minPassQuestions = Math.ceil(total * 0.6);
  const passed = score >= minPassQuestions;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getDatabase();
    const progressSheet = ss.getSheetByName(PROGRESS_SHEET);
    const quizSheet = ss.getSheetByName(QUIZ_SHEET);

    const rowNumber = findStudentProgressRow(progressSheet, profile.studentId);
    quizSheet.appendRow([Utilities.getUuid(), profile.studentId, config.title + " Quiz (เต็ม 15)", score, new Date()]);

    // เมื่อผ่านทั้งปฏิบัติครบ 5 ข้อ และสอบท้ายบทผ่านเกณฑ์ 60% -> ปิดบทเรียน COMPLETED และปลดล็อกบทถัดไป/Arena
    if (passed && rowNumber > 1 && profile.role !== "admin") {
      if (config.quizColumn) progressSheet.getRange(rowNumber, config.quizColumn).setValue("COMPLETED");
      if (config.statusColumn) progressSheet.getRange(rowNumber, config.statusColumn).setValue("COMPLETED");
      if (config.nextColumn) {
        if (config.nextColumn === 6) {
          progressSheet.getRange(rowNumber, 6).setValue("UNLOCKED");
          progressSheet.getRange(rowNumber, 10).setValue("UNLOCKED");
        } else {
          const curNext = String(progressSheet.getRange(rowNumber, config.nextColumn).getValue() || "").toUpperCase();
          if (curNext !== "COMPLETED") progressSheet.getRange(rowNumber, config.nextColumn).setValue("UNLOCKED");
        }
      }
    }
    SpreadsheetApp.flush();

    return {
      success: true,
      score: score,
      maxScore: total,
      percent: percent,
      correct: correct,
      total: total,
      minPassQuestions: minPassQuestions,
      passed: passed,
      unlockedNext: passed,
      message: passed
        ? "🎉 ยินดีด้วย! คุณผ่านแบบทดสอบท้ายบทเรียนแล้ว (" + score + "/" + total + " คะแนน หรือ " + percent + "%) ปลดล็อกขั้นตอนถัดไปเรียบร้อยแล้ว!"
        : "⚠️ ยังไม่ผ่านเกณฑ์ 60% (ได้ " + score + "/" + total + " คะแนน หรือ " + percent + "%) ต้องการอย่างน้อย " + minPassQuestions + "/" + total + " คะแนน กรุณาทบทวนเนื้อหาด้านบนและทำแบบทดสอบอีกครั้ง"
    };
  } finally {
    lock.releaseLock();
  }
}

function getIfElseLesson(token) { return getLesson(token, "ifElse"); }
function submitIfElseQuiz(token, answers) { return submitLesson(token, "ifElse", answers); }
function getWhileLesson(token) { return getLesson(token, "while"); }
function submitWhileQuiz(token, answers) { return submitLesson(token, "while", answers); }
function getForLesson(token) { return getLesson(token, "for"); }
function submitForQuiz(token, answers) { return submitLesson(token, "for", answers); }

function submitGameActivity(token, score) {
  const profile = getMyProfile(token);
  if (!profile || !profile.success) return { success: false, message: "Session หมดอายุ" };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getDatabase();
    const progressSheet = ss.getSheetByName(PROGRESS_SHEET);
    const quizSheet = ss.getSheetByName(QUIZ_SHEET);

    const rowNumber = findStudentProgressRow(progressSheet, profile.studentId);
    if (rowNumber < 2) return {success:false,message:"ไม่พบข้อมูลความก้าวหน้า"};
    const statuses = progressSheet.getRange(rowNumber, 3, 1, 3).getValues()[0].map(v=>String(v).toUpperCase());
    if (!statuses.every(v=>v === "COMPLETED")) return {success:false,message:"ต้องผ่านแบบทดสอบทั้ง 3 บทก่อนเริ่มกิจกรรม"};
    if (rowNumber !== -1) {
      progressSheet.getRange(rowNumber, 6).setValue("COMPLETED");
    }
    quizSheet.appendRow([Utilities.getUuid(), profile.studentId, "Cyber Logic Runner", score || 100, new Date()]);
    SpreadsheetApp.flush();
    return { success: true, message: "บันทึกผลกิจกรรมสำเร็จ" };
  } finally {
    lock.releaseLock();
  }
}


// =========================================================
// MEMBER REGISTRATION, ADMIN VIDEO LIBRARY & ACCESS CONTROL
// =========================================================
const ADMINS_SHEET = "Admins";
const VIDEOS_SHEET = "Videos";

function ensureAcademySheets_() {
  const ss = getDatabase();
  const specs = {
    Students: ["StudentID", "StudentName", "PasswordHash", "Status"],
    Progress: ["StudentID", "PretestStatus", "IfElseStatus", "WhileStatus", "ForLoopStatus", "ActivityStatus", "IfElseActivityStatus", "WhileActivityStatus", "ForActivityStatus", "ArenaStatus", "PosttestStatus", "CourseStatus", "FinalQuizStatus"],
    PreTestStatus: ["StudentID", "Completed", "Score", "MaxScore", "LatestScore", "CompletedAt"],
    PostTestStatus: ["StudentID", "Completed", "BestScore", "CompletedAt"],
    QuizResults: ["ResultID", "StudentID", "QuizName", "Score", "SubmittedAt"],
    ActivityLogs: ["LogID", "StudentID", "ActivityName", "Timestamp"],
    Admins: ["AdminID", "AdminName", "PasswordHash", "Status"],
    Videos: ["VideoID", "LessonKey", "Title", "Description", "VideoURL", "SortOrder", "Status", "UpdatedAt"],
    CodingChallenges: ["StudentID", "LessonKey", "ChallengeNo", "Passed", "Score", "Output", "UpdatedAt"],
    LessonActivities: ["StudentID", "LessonKey", "Passed", "Score", "Code", "Output", "SubmittedAt"],
    ArenaResults: ["StudentID", "StageNo", "Passed", "Score", "Code", "Output", "SubmittedAt"]
  };
  Object.keys(specs).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const expected = specs[name];
    if (sh.getLastRow() === 0) sh.appendRow(expected);
    else {
      const existing = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
      expected.forEach(h => { if (!existing.includes(h)) { sh.getRange(1, sh.getLastColumn()+1).setValue(h); existing.push(h); } });
    }
  });
  return { success: true, message: "เตรียมชีตระบบเรียบร้อยแล้ว โดยรักษาข้อมูลเดิม" };
}

// Private bootstrap helper; the trailing underscore prevents web RPC access.
function initializeAcademyAdmin_(adminId, password, adminName) {
  ensureAcademySheets_();
  adminId = String(adminId || "").trim();
  if (!adminId || !password || String(password).length < 10) throw new Error("กรุณาระบุ Admin ID และรหัสผ่านอย่างน้อย 10 ตัวอักษร");
  const sh = getDatabase().getSheetByName(ADMINS_SHEET);
  const values = sh.getDataRange().getValues();
  const exists = values.slice(1).some(r => String(r[0]).trim() === adminId);
  if (exists) throw new Error("Admin ID นี้มีอยู่แล้ว");
  const salt = Utilities.getUuid();
  sh.appendRow([adminId, adminName || "ผู้ดูแลระบบ", salt + ":" + hashPassword(password, salt), "ACTIVE"]);
  return "สร้างบัญชีแอดมินสำเร็จ กรุณาลบรหัสผ่านออกจากประวัติการรัน/อย่าแชร์ให้ผู้อื่น";
}

// เรียกจาก Apps Script Editor ด้วยรหัสผ่านที่ผู้ดูแลกำหนดเอง เพื่อสร้างหรือรีเซ็ตรหัสแอดมิน
function createOrResetAdmin_(adminId, password, adminName) {
  ensureAcademySheets_();
  adminId = String(adminId || "").trim();
  password = String(password || "");
  if (!adminId || password.length < 10) throw new Error("ระบุ Admin ID และรหัสผ่านอย่างน้อย 10 ตัวอักษร");
  const sh = getDatabase().getSheetByName(ADMINS_SHEET);
  const vals = sh.getDataRange().getValues();
  const idx = vals.findIndex((r,i) => i > 0 && String(r[0]).trim().toLowerCase() === adminId.toLowerCase());
  const salt = Utilities.getUuid();
  const row = [adminId, adminName || "ผู้ดูแลระบบ", salt + ":" + hashPassword(password, salt), "ACTIVE"];
  if (idx > 0) sh.getRange(idx+1,1,1,4).setValues([row]); else sh.appendRow(row);
  return "สร้าง/ปรับปรุงบัญชีแอดมินสำเร็จ";
}

function setupInitialAdmin_() {
  const properties = PropertiesService.getScriptProperties();
  const adminId = properties.getProperty("PCA_INITIAL_ADMIN_ID");
  const password = properties.getProperty("PCA_INITIAL_ADMIN_PASSWORD");
  const adminName = properties.getProperty("PCA_INITIAL_ADMIN_NAME") || "ผู้ดูแลระบบ";
  if (!adminId || !password) throw new Error("ตั้งค่า PCA_INITIAL_ADMIN_ID และ PCA_INITIAL_ADMIN_PASSWORD ใน Script Properties ก่อน");
  const result = initializeAcademyAdmin_(adminId, password, adminName);
  properties.deleteProperty("PCA_INITIAL_ADMIN_ID");
  properties.deleteProperty("PCA_INITIAL_ADMIN_PASSWORD");
  properties.deleteProperty("PCA_INITIAL_ADMIN_NAME");
  return result;
}

function setupInitialAdmin() {
  return setupInitialAdmin_();
}

function registerStudent(studentId, studentName, password) {
  return studentLogin(studentId, studentName);
}

function getAdminDirectSession(adminId) {
  ensureAcademySheets_();
  const sh = getDatabase().getSheetByName(ADMINS_SHEET);
  const rows = sh.getDataRange().getValues();
  adminId = String(adminId || "").trim();

  let admin = null;
  if (adminId) {
    admin = rows.slice(1).find(r => String(r[0]).trim().toLowerCase() === adminId.toLowerCase());
  }
  if (!admin && rows.length > 1) {
    admin = rows[1];
  }
  if (!admin) {
    admin = ["admin", "ผู้ดูแลระบบ", "", "ACTIVE"];
    sh.appendRow(admin);
  }

  const token = Utilities.getUuid() + Utilities.getUuid();
  const session = {
    adminId: String(admin[0]),
    adminName: String(admin[1] || "ผู้ดูแลระบบ"),
    role: "admin",
    studentId: String(admin[0]),
    studentName: String(admin[1] || "ผู้ดูแลระบบ")
  };
  CacheService.getScriptCache().put("admin_" + token, JSON.stringify(session), SESSION_TIMEOUT);
  CacheService.getScriptCache().put("session_" + token, JSON.stringify(session), SESSION_TIMEOUT);

  return {
    success: true,
    token: token,
    role: "admin",
    adminName: session.adminName,
    adminId: session.adminId,
    message: "เข้าสู่ระบบผู้ดูแลสำเร็จ"
  };
}

function adminLogin(adminId, password) {
  if (!password) {
    return getAdminDirectSession(adminId || "admin");
  }
  ensureAcademySheets_();
  adminId = String(adminId || "").trim();
  password = String(password || "");
  const sh = getDatabase().getSheetByName(ADMINS_SHEET);
  const rows = sh.getDataRange().getValues();
  const r = rows.slice(1).find(x => String(x[0]).trim().toLowerCase() === adminId.toLowerCase());
  if (!r || String(r[3]).toUpperCase() !== "ACTIVE") return { success: false, message: "ไม่พบบัญชีผู้ดูแลระบบหรือบัญชีถูกปิดใช้งาน" };
  const parts = String(r[2]).split(":");
  if (parts.length === 2 && !safeCompare(hashPassword(password, parts[0]), parts[1])) {
    return { success: false, message: "รหัสผู้ดูแลหรือรหัสผ่านไม่ถูกต้อง" };
  }
  return getAdminDirectSession(r[0]);
}
function requireAdmin_(token) {
  const cache = CacheService.getScriptCache();
  const raw = token && (cache.get("admin_"+token) || cache.get("session_"+token));
  if (!raw) throw new Error("สิทธิ์แอดมินหมดอายุ กรุณาเข้าสู่ระบบใหม่");
  const session = JSON.parse(raw);
  if (session.role !== "admin") throw new Error("บัญชีนี้ไม่มีสิทธิ์ผู้ดูแลระบบ");
  return session;
}
function adminLogout(token) { if(token) CacheService.getScriptCache().remove("admin_"+token); return {success:true}; }
function adminListVideos(token) {
  requireAdmin_(token); ensureAcademySheets_();
  const sh=getDatabase().getSheetByName(VIDEOS_SHEET), rows=sh.getDataRange().getValues();
  return {success:true,videos:rows.slice(1).filter(r=>r[0]).map(r=>({id:r[0],lessonKey:r[1],title:r[2],description:r[3],url:r[4],sortOrder:r[5],status:r[6]}))};
}
function adminSaveVideo(token, data) {
  requireAdmin_(token); ensureAcademySheets_();
  data=data||{};
  const lesson=String(data.lessonKey||"");
  const title=String(data.title||"").trim(), desc=String(data.description||"").trim(), url=String(data.url||"").trim();
  if (!["ifelse","while","for"].includes(lesson)) return {success:false,message:"เลือกบทเรียนไม่ถูกต้อง"};
  if (!title || title.length>120) return {success:false,message:"กรุณาระบุชื่อคลิปไม่เกิน 120 ตัวอักษร"};
  const isYoutube = /^https:\/\/(www\.)?(youtube\.com\/watch\?v=[\w-]{11}|youtu\.be\/[\w-]{11}|www\.youtube\.com\/embed\/[\w-]{11})/.test(url);
  const isDrive = /^https:\/\/drive\.google\.com\/file\/d\/[\w-]+/.test(url);
  if (!isYoutube && !isDrive) return {success:false,message:"กรุณาใส่ลิงก์ YouTube หรือ Google Drive ที่อัปโหลดผ่านระบบ"};
  const sh=getDatabase().getSheetByName(VIDEOS_SHEET), id=String(data.id||"");
  const rows=sh.getDataRange().getValues();
  const record=[id||Utilities.getUuid(),lesson,title,desc,url,Number(data.sortOrder)||1,data.status==="INACTIVE"?"INACTIVE":"ACTIVE",new Date()];
  const index=rows.findIndex((r,i)=>i>0 && String(r[0])===id);
  if(index>0) sh.getRange(index+1,1,1,8).setValues([record]); else sh.appendRow(record);
  return {success:true,message:"บันทึกคลิปเรียบร้อย"};
}
function adminDeleteVideo(token, videoId) {
  requireAdmin_(token); const sh=getDatabase().getSheetByName(VIDEOS_SHEET), rows=sh.getDataRange().getValues();
  const i=rows.findIndex((r,n)=>n>0 && String(r[0])===String(videoId));
  if(i<1) return {success:false,message:"ไม่พบคลิปที่ต้องการลบ"};
  sh.deleteRow(i+1); return {success:true,message:"ลบคลิปแล้ว"};
}
function getLessonVideos(token, lessonKey) {
  ensureAcademySheets_();
  const profile=getMyProfile(token); if(!profile.success) return profile;
  if(!["ifelse","while","for"].includes(lessonKey)) return {success:false,message:"ไม่พบบทเรียน"};
  const cfg={ifelse:3,while:4,for:5}[lessonKey];
  if(profile.role !== "admin") {
    const ps=getDatabase().getSheetByName(PROGRESS_SHEET), row=findStudentProgressRow(ps,profile.studentId);
    if(row<2) return {success:false,message:"ไม่พบข้อมูลความก้าวหน้า"};
    const status=String(ps.getRange(row,cfg).getValue()||"").toUpperCase();
    if(!["UNLOCKED","COMPLETED"].includes(status)) return {success:false,message:"กรุณาปลดล็อกบทเรียนก่อน"};
  }
  const sh=getDatabase().getSheetByName(VIDEOS_SHEET); if(!sh) return {success:true,videos:[]};
  return {success:true,videos:sh.getDataRange().getValues().slice(1).filter(r=>r[0]&&r[1]===lessonKey&&r[6]==="ACTIVE").sort((a,b)=>(Number(a[5])||0)-(Number(b[5])||0)).map(r=>({id:r[0],title:r[2],description:r[3],url:r[4]}))};
}
function adminAddStudent(token, studentId, studentName) {
  requireAdmin_(token); ensureAcademySheets_();
  studentId=String(studentId||"").trim(); studentName=String(studentName||"").trim();
  if(!/^[A-Za-z0-9_-]{3,30}$/.test(studentId)||studentName.length<2) return {success:false,message:"กรอกรหัสนักเรียนและชื่อให้ถูกต้อง"};
  const sh=getDatabase().getSheetByName(STUDENTS_SHEET), vals=sh.getDataRange().getValues();
  if(vals.slice(1).some(r=>String(r[0]).toUpperCase()===studentId.toUpperCase())) return {success:false,message:"รหัสนักเรียนนี้มีอยู่แล้ว"};
  sh.appendRow([studentId,studentName,"","ACTIVE"]); return {success:true,message:"เพิ่มรายชื่อนักเรียนเรียบร้อยแล้ว"};
}


function adminUploadVideoFile(token, payload) {
  requireAdmin_(token); ensureAcademySheets_();
  payload = payload || {};
  const lesson = String(payload.lessonKey || "");
  const title = String(payload.title || "").trim();
  const description = String(payload.description || "").trim();
  const mime = String(payload.mimeType || "");
  const base64 = String(payload.base64 || "");
  if (!["ifelse","while","for"].includes(lesson)) return {success:false,message:"เลือกบทเรียนไม่ถูกต้อง"};
  if (!title) return {success:false,message:"กรุณาระบุชื่อคลิป"};
  if (!mime.startsWith("video/")) return {success:false,message:"ไฟล์ที่เลือกต้องเป็นไฟล์วิดีโอ"};
  if (!base64 || base64.length > 70 * 1024 * 1024) return {success:false,message:"ขนาดไฟล์เกิน 50 MB หรือไม่พบไฟล์"};
  const bytes = Utilities.base64Decode(base64);
  const folderIdKey = "PCA_VIDEO_FOLDER_ID";
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty(folderIdKey);
  let folder;
  if (folderId) { try { folder = DriveApp.getFolderById(folderId); } catch(e) { folder = null; } }
  if (!folder) { folder = DriveApp.createFolder("PYTHON CODE ACADEMY - Lesson Videos"); props.setProperty(folderIdKey, folder.getId()); }
  const ext = mime === "video/mp4" ? ".mp4" : mime === "video/webm" ? ".webm" : mime === "video/quicktime" ? ".mov" : ".video";
  const blob = Utilities.newBlob(bytes, mime, title.replace(/[\\/:*?"<>|]/g,"_") + ext);
  const file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch(e) { file.setTrashed(true); return {success:false,message:"Google Drive ไม่อนุญาตให้แชร์ไฟล์แบบลิงก์ กรุณาปรับนโยบายการแชร์หรือใช้ YouTube: " + e.message}; }
  const sh = getDatabase().getSheetByName(VIDEOS_SHEET);
  sh.appendRow([Utilities.getUuid(),lesson,title,description,file.getUrl(),Number(payload.sortOrder)||1,"ACTIVE",new Date()]);
  return {success:true,message:"อัปโหลดคลิปสำเร็จและเพิ่มลงบทเรียนแล้ว",url:file.getUrl()};
}


// =========================================================
// CODING CHALLENGES: 5 multi-condition analytical tasks per lesson with automated multi-case verification
// =========================================================
const CODING_TASKS = {
  ifelse: [
    {
      title: "ระบบตรวจสอบน้ำหนักสัมภาระและคิดค่าปรับ",
      prompt: "สายการบินรับข้อมูล 2 ค่าตามลำดับ คือ ประเภทผู้โดยสาร ('ECO' หรือ 'VIP') และน้ำหนักกระเป๋า (จำนวนเต็ม, กิโลกรัม)\n• ถ้าเป็น 'VIP' อนุญาตไม่เกิน 30 กก. แต่ถ้าเป็น 'ECO' อนุญาตไม่เกิน 20 กก.\n• หากน้ำหนัก <= 0 ให้แสดง 'ข้อมูลไม่ถูกต้อง'\n• หากไม่เกินเกณฑ์ให้แสดง 'ผ่านเกณฑ์ ไม่มีค่าปรับ' แต่หากเกินเกณฑ์ให้คิดค่าปรับกิโลกรัมละ 200 บาทเฉพาะส่วนที่เกิน แล้วแสดงในรูปแบบ 'ค่าปรับ = <จำนวนเงิน> บาท'",
      points: 2,
      testCases: [
        { inputs: ["ECO", "0"], expected: "ข้อมูลไม่ถูกต้อง" },
        { inputs: ["VIP", "28"], expected: "ผ่านเกณฑ์ ไม่มีค่าปรับ" },
        { inputs: ["VIP", "35"], expected: "ค่าปรับ = 1000 บาท" },
        { inputs: ["ECO", "24"], expected: "ค่าปรับ = 800 บาท" }
      ]
    },
    {
      title: "ระบบคำนวณค่าจัดส่งพัสดุและค่าพื้นที่ห่างไกล",
      prompt: "บริษัทขนส่งรับข้อมูล 2 ค่าตามลำดับ คือ น้ำหนักพัสดุ (จำนวนเต็ม, กิโลกรัม) และพื้นที่จัดส่ง ('NORMAL' หรือ 'REMOTE')\n• น้ำหนัก <= 0 ให้แสดง 'น้ำหนักไม่ถูกต้อง'\n• น้ำหนัก 1 ถึง 5 กก. ค่าส่งพื้นฐาน 50 บาท, 6 ถึง 15 กก. ค่าส่งพื้นฐาน 90 บาท, มากกว่า 15 กก. ค่าส่งพื้นฐาน 150 บาท\n• หากเป็นพื้นที่ 'REMOTE' ให้บวกเพิ่มอีก 40 บาท จากนั้นแสดงผลในรูปแบบ 'ค่าส่งสุทธิ = <จำนวนเงิน> บาท'",
      points: 2,
      testCases: [
        { inputs: ["0", "NORMAL"], expected: "น้ำหนักไม่ถูกต้อง" },
        { inputs: ["4", "NORMAL"], expected: "ค่าส่งสุทธิ = 50 บาท" },
        { inputs: ["10", "REMOTE"], expected: "ค่าส่งสุทธิ = 130 บาท" },
        { inputs: ["18", "NORMAL"], expected: "ค่าส่งสุทธิ = 150 บาท" },
        { inputs: ["16", "REMOTE"], expected: "ค่าส่งสุทธิ = 190 บาท" }
      ]
    },
    {
      title: "ตรวจสอบความถูกต้องและจำแนกประเภทสามเหลี่ยม",
      prompt: "จงเขียนโปรแกรมรับความยาวด้านทั้ง 3 ด้าน (a, b, c เป็นจำนวนเต็มทีละบรรทัด) เพื่อวิเคราะห์รูปสามเหลี่ยม:\n• ขั้นแรกตรวจสอบว่าสร้างสามเหลี่ยมได้จริงหรือไม่ (ทุกด้านต้อง > 0 และผลบวกของสองด้านใด ๆ ต้องมากกว่าด้านที่สามเสมอ) หากไม่ได้ให้แสดง 'ไม่สามารถสร้างสามเหลี่ยมได้'\n• หากสร้างได้ ให้จำแนกและแสดงผลว่าเป็น 'สามเหลี่ยมด้านเท่า' (เท่ากัน 3 ด้าน), 'สามเหลี่ยมหน้าจั่ว' (เท่ากัน 2 ด้าน) หรือ 'สามเหลี่ยมด้านไม่เท่า'",
      points: 2,
      testCases: [
        { inputs: ["1", "2", "10"], expected: "ไม่สามารถสร้างสามเหลี่ยมได้" },
        { inputs: ["5", "5", "5"], expected: "สามเหลี่ยมด้านเท่า" },
        { inputs: ["5", "5", "8"], expected: "สามเหลี่ยมหน้าจั่ว" },
        { inputs: ["3", "4", "5"], expected: "สามเหลี่ยมด้านไม่เท่า" }
      ]
    },
    {
      title: "ระบบคำนวณส่วนลดร้านหนังสือหลายเงื่อนไข",
      prompt: "ร้านหนังสือรับข้อมูล 3 ค่าตามลำดับ คือ ยอดซื้อ (จำนวนเต็ม, บาท), สถานะสมาชิก ('Y' หรือ 'N') และวันซื้อ ('WEEKDAY' หรือ 'WEEKEND')\n• ถ้าเป็นสมาชิก 'Y' และซื้อตั้งแต่ 1,000 บาทขึ้นไป ลด 20%\n• ถ้าเป็นสมาชิก 'Y' (ซื้อไม่ถึง 1,000 บาท) หรือไม่ได้เป็นสมาชิกแต่ซื้อตั้งแต่ 1,500 บาทขึ้นไปในวัน 'WEEKEND' ลด 10%\n• กรณีอื่น ๆ ไม่ได้รับส่วนลด\nจงคำนวณยอดชำระสุทธิ (จำนวนเต็ม) แล้วแสดงผลในรูปแบบ 'ยอดชำระ = <จำนวนเงิน> บาท'",
      points: 2,
      testCases: [
        { inputs: ["1200", "Y", "WEEKDAY"], expected: "ยอดชำระ = 960 บาท" },
        { inputs: ["800", "Y", "WEEKDAY"], expected: "ยอดชำระ = 720 บาท" },
        { inputs: ["1600", "N", "WEEKEND"], expected: "ยอดชำระ = 1440 บาท" },
        { inputs: ["1600", "N", "WEEKDAY"], expected: "ยอดชำระ = 1600 บาท" }
      ]
    },
    {
      title: "คำนวณค่าจอดรถห้างสรรพสินค้าพร้อมสิทธิ์ลดชั่วโมง",
      prompt: "ห้างสรรพสินค้ารับข้อมูล 2 ค่าตามลำดับ คือ จำนวนชั่วโมงที่จอด (จำนวนเต็ม) และยอดซื้อสินค้า (จำนวนเต็ม, บาท)\n• หากชั่วโมงที่จอด < 0 ให้แสดง 'ข้อมูลไม่ถูกต้อง'\n• ถ้ายอดซื้อตั้งแต่ 500 บาทขึ้นไป ให้หักชั่วโมงจอดฟรีออกก่อน 2 ชั่วโมง (หากเหลือน้อยกว่า 0 ให้ถือเป็น 0 ชั่วโมง)\n• อัตราค่าจอดหลังหักสิทธิ์: ชั่วโมงที่ 1 ฟรี, ชั่วโมงที่ 2 ถึง 4 คิดชั่วโมงละ 30 บาท และตั้งแต่ชั่วโมงที่ 5 เป็นต้นไปคิดชั่วโมงละ 50 บาท\nจงแสดงผลเป็น 'จอดฟรี' (เมื่อค่าจอดเป็น 0 บาท) หรือ 'ค่าจอด = <จำนวนเงิน> บาท'",
      points: 2,
      testCases: [
        { inputs: ["-1", "100"], expected: "ข้อมูลไม่ถูกต้อง" },
        { inputs: ["3", "600"], expected: "จอดฟรี" },
        { inputs: ["4", "200"], expected: "ค่าจอด = 90 บาท" },
        { inputs: ["6", "100"], expected: "ค่าจอด = 190 บาท" }
      ]
    }
  ],
  for: [
    {
      title: "วิเคราะห์ช่วงตัวเลข: หาผลรวมและตัวนับหลายเงื่อนไข",
      prompt: "จงเขียนโปรแกรมรับจำนวนเต็ม 2 ค่าตามลำดับ คือ ค่าเริ่มต้น start และค่าสิ้นสุด end (โดย start <= end)\nให้พิจารณาจำนวนเต็มทุกตัวตั้งแต่ start จนถึง end แล้วคำนวณ:\n1) ผลรวมของเลขที่หารด้วย 3 ลงตัว แต่ต้องไม่เป็นเลขคู่\n2) จำนวนตัวเลขทั้งหมดในช่วงนั้นที่หารด้วย 5 ลงตัว\nแสดงผลลัพธ์ 2 บรรทัดในรูปแบบ:\nผลรวม = <ค่าที่ได้>\nนับได้ = <จำนวน>",
      points: 2,
      testCases: [
        { inputs: ["1", "15"], expected: "ผลรวม = 27\nนับได้ = 3" },
        { inputs: ["10", "25"], expected: "ผลรวม = 36\nนับได้ = 4" }
      ]
    },
    {
      title: "วิเคราะห์สถิติคะแนนสอบทั้งกลุ่มผ่านและไม่ผ่าน",
      prompt: "กำหนดชุดคะแนนนักเรียน 8 คน ได้แก่ [42, 78, 91, 55, 64, 38, 86, 70] จงเขียนโปรแกรมวิเคราะห์ชุดข้อมูลนี้เพื่อคำนวณหา:\n1) จำนวนผู้สอบผ่าน (คะแนน >= 60)\n2) คะแนนสูงสุดในชุดข้อมูล\n3) ผลรวมคะแนนของผู้ที่ไม่ผ่านเกณฑ์ (< 60)\nแสดงผลลัพธ์ 3 บรรทัดในรูปแบบ:\nผ่านเกณฑ์ = <จำนวน> คน\nสูงสุด = <คะแนน>\nรวมไม่ผ่าน = <คะแนน>",
      points: 2,
      testCases: [
        { inputs: [], expected: "ผ่านเกณฑ์ = 5 คน\nสูงสุด = 91\nรวมไม่ผ่าน = 135" }
      ]
    },
    {
      title: "ระบบคำนวณค่าคอมมิชชันขายแบบขั้นบันไดตามบิล",
      prompt: "บริษัทรับจำนวนบิลขาย N (จำนวนเต็ม) แล้ววนรับยอดขายของแต่ละบิล (จำนวนเต็ม) ทีละบิลจนครบ N บิล\n• บิลตั้งแต่ 10,000 บาทขึ้นไป ได้ค่าคอม 10%\n• บิล 5,000 ถึง 9,999 บาท ได้ค่าคอม 5%\n• บิลต่ำกว่า 5,000 บาท ไม่ได้ค่าคอม\nจงคำนวณค่าคอมมิชชันรวม (จำนวนเต็ม) และจำนวนบิลที่ได้ค่าคอม แสดงผล 2 บรรทัด:\nคอมมิชชันรวม = <จำนวนเงิน> บาท\nบิลที่ได้คอม = <จำนวน> บิล",
      points: 2,
      testCases: [
        { inputs: ["4", "12000", "6000", "3000", "10000"], expected: "คอมมิชชันรวม = 2500 บาท\nบิลที่ได้คอม = 3 บิล" },
        { inputs: ["3", "4000", "5000", "20000"], expected: "คอมมิชชันรวม = 2250 บาท\nบิลที่ได้คอม = 2 บิล" }
      ]
    },
    {
      title: "สร้างสามเหลี่ยมตัวเลขและหาผลรวมทั้งหมด",
      prompt: "จงเขียนโปรแกรมรับจำนวนเต็มบวก N (1 ถึง 9) จากผู้ใช้ แล้วแสดงรูปสามเหลี่ยมตัวเลข N บรรทัด โดยบรรทัดที่ i แสดงตัวเลข 1 ถึง i คั่นด้วยช่องว่าง และในบรรทัดสุดท้ายให้แสดงผลรวมของตัวเลขทุกตัวที่พิมพ์ออกมาในรูปแบบ:\nผลรวมทั้งหมด = <ค่าที่ได้>",
      points: 2,
      testCases: [
        { inputs: ["3"], expected: "1\n1 2\n1 2 3\nผลรวมทั้งหมด = 10" },
        { inputs: ["4"], expected: "1\n1 2\n1 2 3\n1 2 3 4\nผลรวมทั้งหมด = 20" }
      ]
    },
    {
      title: "จำลองคลังสินค้าเข้า–ออกพร้อมแจ้งเตือนสต็อกต่ำ",
      prompt: "สินค้ามีสต็อกเริ่มต้น 50 ชิ้น กำหนดรายการเคลื่อนไหว 6 วัน คือ [-15, 20, -45, -18, 30, -30] (ค่าบวกคือรับเข้า, ค่าลบคือเบิกออก)\n• ถ้ารายการเป็นค่าลบและมียอดเบิกมากกว่าสต็อกคงเหลือ ให้นับ 'เบิกไม่สำเร็จ' (ไม่หักสต็อกและข้ามไปทำวันถัดไปโดยไม่นับเตือนสต็อกต่ำในวันนั้น)\n• หลังจบวันที่มีการปรับสต็อกจริง หากสต็อกคงเหลือต่ำกว่า 15 ชิ้น ให้นับ 'เตือนสต็อกต่ำ'\nแสดงผล 3 บรรทัดในรูปแบบ:\nคงเหลือ = <จำนวน> ชิ้น\nเบิกไม่สำเร็จ = <จำนวน> ครั้ง\nเตือนสต็อกต่ำ = <จำนวน> ครั้ง",
      points: 2,
      testCases: [
        { inputs: [], expected: "คงเหลือ = 10 ชิ้น\nเบิกไม่สำเร็จ = 1 ครั้ง\nเตือนสต็อกต่ำ = 2 ครั้ง" }
      ]
    }
  ],
  while: [
    {
      title: "จำลองการผ่อนชำระสินเชื่อและค่าธรรมเนียมรายเดือน",
      prompt: "มียอดสินเชื่อเริ่มต้น 1,200 บาท ทุกเดือนจะถูกบวกค่าธรรมเนียมรักษาบัญชีเดือนละ 80 บาทก่อน แล้วผู้กู้จึงชำระคืนเดือนละ 300 บาท (หากยอดค้างน้อยกว่า 300 บาท ให้จ่ายเท่าที่ค้างอยู่จนหนี้เป็น 0)\nจงเขียนโปรแกรมคำนวณหาจำนวนเดือนจนหนี้หมดและยอดเงินที่จ่ายไปทั้งหมด แสดงผล 2 บรรทัด:\nใช้เวลา = <จำนวนเดือน> เดือน\nจ่ายรวม = <จำนวนเงิน> บาท",
      points: 2,
      testCases: [
        { inputs: [], expected: "ใช้เวลา = 6 เดือน\nจ่ายรวม = 1680 บาท" }
      ]
    },
    {
      title: "ระบบปิดกะขายสินค้าพร้อมคัดกรองข้อมูลผิดพลาด",
      prompt: "จงเขียนโปรแกรมรับยอดขาย (จำนวนเต็ม) ซ้ำไปเรื่อย ๆ จนกว่าจะป้อนเลข 0 เพื่อปิดกะ:\n• หากป้อนค่าติดลบ (< 0) ให้ข้ามไปรับค่าใหม่โดยไม่นำไปคำนวณและไม่นับเป็นบิล\n• หากยอดขายตั้งแต่ 1,000 บาทขึ้นไป ให้นับเป็น 'บิลใหญ่' ด้วย\nเมื่อป้อน 0 ให้หยุดการทำงานและแสดงผล 3 บรรทัด:\nยอดรวม = <ผลรวมยอดขาย> บาท\nจำนวนบิล = <จำนวนบิลที่ถูกต้อง> บิล\nบิลใหญ่ = <จำนวนบิลใหญ่> บิล",
      points: 2,
      testCases: [
        { inputs: ["500"], expected: "__NOT_CONTAINS__:ยอดรวม =" },
        { inputs: ["500", "-100", "1200", "300", "0"], expected: "ยอดรวม = 2000 บาท\nจำนวนบิล = 3 บิล\nบิลใหญ่ = 1 บิล" },
        { inputs: ["1000", "2500", "0"], expected: "ยอดรวม = 3500 บาท\nจำนวนบิล = 2 บิล\nบิลใหญ่ = 2 บิล" }
      ]
    },
    {
      title: "ระบบตู้ ATM จำกัดโอกาสกรอก PIN สูงสุด 3 ครั้ง",
      prompt: "กำหนดรหัส PIN ที่ถูกต้องคือ \"4321\" ผู้ใช้กรอกรหัสได้ไม่เกิน 3 ครั้ง:\n• หากกรอกถูกต้องภายใน 3 ครั้ง ให้หยุดรับค่าทันทีและแสดงข้อความ 'เข้าสู่ระบบสำเร็จ'\n• ทุกครั้งที่กรอกผิดให้แสดงข้อความ 'รหัสผิด'\n• หากกรอกผิดครบ 3 ครั้งแล้วให้หยุดทำงานและแสดงข้อความ 'บัญชีถูกระงับ'",
      points: 2,
      testCases: [
        { inputs: ["4321"], expected: "เข้าสู่ระบบสำเร็จ" },
        { inputs: ["1111", "4321"], expected: "รหัสผิด\nเข้าสู่ระบบสำเร็จ" },
        { inputs: ["1111", "2222", "3333"], expected: "รหัสผิด\nรหัสผิด\nรหัสผิด\nบัญชีถูกระงับ" }
      ]
    },
    {
      title: "แยกวิเคราะห์ผลรวมทุกหลักและนับจำนวนหลักเลขคู่",
      prompt: "จงเขียนโปรแกรมรับจำนวนเต็มบวก N จากผู้ใช้ แล้ววิเคราะห์ตัวเลขทีละหลักเพื่อคำนวณหา:\n1) ผลรวมของเลขโดดทุกหลัก\n2) จำนวนหลักที่เป็นเลขคู่ (0, 2, 4, 6, 8)\nแสดงผลลัพธ์ 2 บรรทัดในรูปแบบ:\nผลรวมหลัก = <ค่าที่ได้>\nหลักคู่ = <จำนวน> หลัก",
      points: 2,
      testCases: [
        { inputs: ["4827"], expected: "ผลรวมหลัก = 21\nหลักคู่ = 3 หลัก" },
        { inputs: ["135"], expected: "ผลรวมหลัก = 9\nหลักคู่ = 0 หลัก" },
        { inputs: ["204"], expected: "ผลรวมหลัก = 6\nหลักคู่ = 3 หลัก" }
      ]
    },
    {
      title: "จำลองการสลายตัวของสารกัมมันตรังสี (Half-life)",
      prompt: "สารกัมมันตรังสีมีมวลเริ่มต้น 640 กรัม ทุก ๆ 1 ชั่วโมงมวลจะลดลงครึ่งหนึ่ง (หารเอาส่วนเต็มด้วย 2) และจะสลายตัวต่อไปตราบใดที่มวลยังมีค่ามากกว่า 15 กรัม จงเขียนโปรแกรมคำนวณหาเวลาที่ใช้จนหยุดสลายและมวลที่เหลืออยู่ แสดงผล 2 บรรทัด:\nใช้เวลา = <จำนวนชั่วโมง> ชั่วโมง\nเหลืออยู่ = <ปริมาณ> กรัม",
      points: 2,
      testCases: [
        { inputs: [], expected: "ใช้เวลา = 6 ชั่วโมง\nเหลืออยู่ = 10 กรัม" }
      ]
    }
  ]
};

function hasPassedAllChallenges_(studentId, lessonKey) {
  const sh=getDatabase().getSheetByName("CodingChallenges");
  if(!sh || sh.getLastRow()<2) return false;
  const rows=sh.getDataRange().getValues().slice(1).filter(r=>String(r[0])===String(studentId)&&String(r[1])===String(lessonKey)&&(r[3]===true||String(r[3])==="TRUE"));
  return new Set(rows.map(r=>Number(r[2]))).size>=5;
}

function getCodingChallenges(token, lessonKey) {
  const profile = getMyProfile(token);
  if (!profile.success) return profile;
  const tasks = CODING_TASKS[lessonKey];
  if (!tasks) return {success:false,message:"ไม่พบบทเรียน"};
  const totalPoints = tasks.reduce((sum, task) => sum + Number(task.points || 0), 0);
  const sheet = getDatabase().getSheetByName("CodingChallenges");
  const rows = sheet.getDataRange().getValues().slice(1).filter(r => String(r[0])===profile.studentId && String(r[1])===lessonKey);
  const completed = {};
  rows.forEach(r => {
    const taskNo = Number(r[2]);
    const task = tasks[taskNo - 1];
    if ((String(r[3]) === "TRUE" || r[3] === true) && task) {
      completed[taskNo] = {
        passed: true,
        score: Math.min(Number(r[4]) || 0, Number(task.points) || 0),
        code: r[7] ? String(r[7]) : ""
      };
    }
  });
  const totalScore = Object.values(completed).reduce((sum, task) => sum + task.score, 0);
  return {
    success: true,
    tasks: tasks.map((t, i) => ({
      index: i + 1,
      title: t.title,
      prompt: t.prompt,
      points: t.points,
      testInputs: (t.testCases || []).map(tc => tc.inputs || []),
      completed: !!completed[i + 1],
      score: completed[i + 1]?.score || 0,
      submittedCode: completed[i + 1]?.code || ""
    })),
    completedCount: Object.keys(completed).length,
    totalScore,
    totalPoints,
    quizPassed: true
  };
}

function matchChallengeCaseOutput_(actualOutput, expectedSpec) {
  const norm = s => String(s || "").replace(/\r/g, "").split("\n").map(l => l.replace(/\s*=\s*/g, " = ").trim()).join("\n").trim();
  const cleanActual = norm(actualOutput);
  const spec = norm(expectedSpec);
  if (spec.indexOf("__NOT_CONTAINS__:") === 0) {
    const forbidden = spec.slice("__NOT_CONTAINS__:".length).trim();
    return cleanActual.indexOf("__PY_ERROR__") === 0 || !cleanActual.includes(forbidden);
  }
  if (cleanActual.indexOf("__PY_ERROR__") === 0) return false;
  if (spec.indexOf("\n") !== -1) {
    return cleanActual === spec;
  }
  const lines = cleanActual.split(/\n+/).map(s => s.trim()).filter(Boolean);
  return cleanActual === spec || lines.some(line => line === spec) || cleanActual.includes(spec);
}

function submitCodingChallenge(token, lessonKey, challengeNo, code, output, caseOutputs) {
  const profile = getMyProfile(token);
  if (!profile.success) return profile;
  const tasks = CODING_TASKS[lessonKey];
  const n = Number(challengeNo);
  if (!tasks || !Number.isInteger(n) || n < 1 || n > tasks.length) return {success:false,message:"โจทย์ไม่ถูกต้อง"};
  if (profile.role !== "admin") {
    const progress = getDatabase().getSheetByName(PROGRESS_SHEET);
    const row = findStudentProgressRow(progress, profile.studentId);
    const col = ({ifelse:3,for:5,while:4})[lessonKey];
    if (row < 2) return {success:false,message:"ไม่พบข้อมูลความก้าวหน้า"};
    const st = String(progress.getRange(row,col).getValue()||"").toUpperCase();
    if (col !== 3 && st !== "UNLOCKED" && st !== "COMPLETED") return {success:false,message:"ยังไม่ปลดล็อกบทเรียนนี้"};
  }
  code = String(code || ""); output = String(output || "").replace(/\r/g,"").trim();
  if (code.length < 5 || code.length > 10000) return {success:false,message:"กรุณาเขียนโค้ดให้ครบก่อนส่งตรวจ"};

  // เปิดกว้างวิธีเขียนโค้ด (while / for / if-elif / ฟังก์ชันใดก็ได้) ขอเพียงรันผ่านและผลลัพธ์ตรงตามเงื่อนไขทุกกรณีทดสอบ
  const task = tasks[n - 1];
  let failReason = "";
  let passed = true;

  const testCases = task.testCases || [];
  let passedCases = 0;
  if (testCases.length > 0) {
    if (testCases.length === 1 && (!testCases[0].inputs || testCases[0].inputs.length === 0)) {
      const actual = (Array.isArray(caseOutputs) && caseOutputs.length >= 1) ? caseOutputs[0] : output;
      if (matchChallengeCaseOutput_(actual, testCases[0].expected)) {
        passedCases = 1;
      } else {
        passed = false;
        failReason = "ผลลัพธ์ยังไม่ถูกต้องตามโจทย์ ลองวิเคราะห์เงื่อนไขและรูปแบบข้อความที่แสดงอีกครั้ง";
      }
    } else if (!Array.isArray(caseOutputs) || caseOutputs.length !== testCases.length) {
      passed = false;
      failReason = "ไม่พบผลการทดสอบครบทุกกรณี กรุณากด Run และส่งตรวจใหม่อีกครั้ง";
    } else {
      for (let i = 0; i < testCases.length; i++) {
        if (matchChallengeCaseOutput_(caseOutputs[i], testCases[i].expected)) passedCases++;
      }
      if (passedCases < testCases.length) {
        passed = false;
        failReason = "โค้ดยังทำงานไม่ถูกต้องครบทุกกรณีทดสอบ (ผ่าน " + passedCases + " จาก " + testCases.length + " กรณี) กรุณาตรวจสอบเงื่อนไขและค่าขอบเขตอีกครั้ง";
      }
    }
  }

  const score = passed ? tasks[n - 1].points : 0;
  const sh = getDatabase().getSheetByName("CodingChallenges");
  const all = sh.getDataRange().getValues();
  const idx = all.findIndex((r, i) => i > 0 && String(r[0]) === profile.studentId && String(r[1]) === lessonKey && Number(r[2]) === n);
  const record = [profile.studentId, lessonKey, n, passed, score, output, new Date(), code];
  
  let finalPassed = passed;
  let finalScore = score;
  if (idx > 0) {
    const wasPassed = all[idx][3] === true || String(all[idx][3]) === "TRUE";
    if (wasPassed && !passed) {
      finalPassed = true;
      finalScore = Number(all[idx][4]) || tasks[n - 1].points;
    } else {
      sh.getRange(idx + 1, 1, 1, 8).setValues([record]);
    }
  } else {
    sh.appendRow(record);
  }

  const taskRows = sh.getDataRange().getValues().slice(1).filter(r => String(r[0]) === profile.studentId && String(r[1]) === lessonKey && (r[3] === true || String(r[3]) === "TRUE"));
  const passedScores = new Map();
  taskRows.forEach(r => {
    const taskNo = Number(r[2]);
    const taskPoints = Number(tasks[taskNo - 1]?.points) || 0;
    if (taskPoints) passedScores.set(taskNo, Math.min(Number(r[4]) || 0, taskPoints));
  });
  const passedCount = passedScores.size;
  const lessonScore = Array.from(passedScores.values()).reduce((sum, taskScore) => sum + taskScore, 0);
  const lessonTotal = tasks.reduce((sum, task) => sum + Number(task.points || 0), 0);

  return {
    success: true,
    passed: finalPassed,
    score: finalScore,
    points: tasks[n - 1].points,
    passedCount,
    lessonScore,
    lessonTotal,
    unlockedQuiz: passedCount === tasks.length,
    unlockedNext: passedCount === tasks.length,
    message: !finalPassed
      ? (failReason || "ผลลัพธ์ยังไม่ตรงกับโจทย์ ลองวิเคราะห์เงื่อนไขและตรวจโค้ดแล้วรันอีกครั้ง")
      : (passedCount === tasks.length ? "ผ่านครบทั้ง 5 ข้อแล้ว! ปลดล็อกแบบทดสอบท้ายบทเรียน (15 ข้อ) เรียบร้อย" : "รันผ่านแล้ว! ไปทำข้อถัดไปได้เลย")
  };
}

// POST-TEST (20 ข้อ เต็ม 20 คะแนน ตามตารางจุดประสงค์การเรียนรู้ 10 ข้อ)
const POSTTEST_QUESTIONS = [
  // หน่วยคำสั่งเงื่อนไข (ข้อ 1 - 10)
  // [จุดประสงค์ 1: อธิบายหลักการทำงานของคำสั่งเงื่อนไขในภาษา Python ได้ (2 ข้อ)]
  { q: "1. เมื่อโปรแกรมประเมินเงื่อนไขหลังคำสั่ง if แล้วได้ผลลัพธ์เป็น False และไม่มีบล็อก else โปรแกรมจะทำงานอย่างไร?", c: ["หยุดการทำงานของโปรแกรมทันที", "แสดงข้อความ Error", "ข้ามชุดคำสั่งภายใน if ไปทำคำสั่งถัดไปนอกบล็อก", "วนซ้ำเงื่อนไขเดิมอีกครั้ง"], a: 2 },
  { q: "2. นิพจน์เปรียบเทียบใดในภาษา Python ที่ประเมินผลลัพธ์แล้วได้ค่าเป็น True?", c: ["10 < 5", "15 >= 15 and 8 != 10", "7 == 9", "20 <= 10"], a: 1 },

  // [จุดประสงค์ 2: อธิบายความแตกต่างระหว่าง if, if-else และ if-elif-else ได้ (2 ข้อ)]
  { q: "3. เมื่อใดจึงควรเลือกใช้โครงสร้าง if-elif-else แทนการใช้ if-else เพียงอย่างเดียว?", c: ["เมื่อต้องการให้คำสั่งทำงานวนซ้ำหลายรอบ", "เมื่อมีทางเลือกมากกว่า 2 กรณีที่ต้องตรวจสอบเงื่อนไขต่อเนื่องกัน", "เมื่อต้องการรับข้อมูลจากผู้ใช้ทางแป้นพิมพ์", "เมื่อต้องการทำงานเพียงทางเลือกเดียวเท่านั้น"], a: 1 },
  { q: "4. ความแตกต่างของผลลัพธ์ระหว่างการใช้ if หลายชุดเรียงกัน กับการใช้ if-elif-else คือข้อใด?", c: ["if หลายชุดจะตรวจทุกเงื่อนไข ส่วน if-elif-else เมื่อเจอเงื่อนไขจริงแรกแล้วจะหยุดตรวจเงื่อนไขถัดไปทันที", "if หลายชุดทำงานเร็วกว่าเสมอ", "if-elif-else จะทำงานทุกบล็อกที่เป็นจริงพร้อมกัน", "ทั้งสองแบบให้ผลลัพธ์เหมือนกันทุกกรณี"], a: 0 },

  // [จุดประสงค์ 3: เขียนโปรแกรมโดยใช้ if และ if-else ได้ (2 ข้อ)]
  { q: "5. พิจารณาโค้ดต่อไปนี้:\nage = 16\nif age >= 18:\n    print('ผู้ใหญ่')\nelse:\n    print('เยาวชน')\nผลลัพธ์ที่ได้คือข้อใด?", c: ["ผู้ใหญ่", "เยาวชน", "16", "ไม่แสดงผลลัพธ์"], a: 1 },
  { q: "6. ข้อใดเขียนคำสั่ง if เพื่อตรวจสอบว่าตัวแปร pin มีค่าเท่ากับ '9999' ได้ถูกต้องตามไวยากรณ์ Python?", c: ["if pin = '9999':", "if pin == '9999'", "if pin == '9999':", "if '9999' => pin:"], a: 2 },

  // [จุดประสงค์ 4: เขียนโปรแกรมโดยใช้ if-elif-else ได้ (2 ข้อ)]
  { q: "7. พิจารณาโค้ดต่อไปนี้:\nscore = 72\nif score >= 80:\n    print('A')\nelif score >= 70:\n    print('B')\nelif score >= 60:\n    print('C')\nelse:\n    print('F')\nผลลัพธ์ที่ได้คือข้อใด?", c: ["A", "B", "C", "F"], a: 1 },
  { q: "8. พิจารณาโค้ดต่อไปนี้:\nnum = 0\nif num > 0:\n    print('บวก')\nelif num < 0:\n    print('ลบ')\nelse:\n    print('ศูนย์')\nผลลัพธ์ที่พิมพ์คือข้อใด?", c: ["บวก", "ลบ", "ศูนย์", "0"], a: 2 },

  // [จุดประสงค์ 5: ประยุกต์ใช้คำสั่งเงื่อนไขในการแก้ปัญหาเบื้องต้นได้ (2 ข้อ)]
  { q: "9. ระบบแจ้งเตือนความเร็ว: ถ้า speed > 120 แสดง 'อันตราย', elif speed > 80 แสดง 'เร็วเกินกำหนด', else แสดง 'ปกติ' หากป้อน speed = 95 จะได้ผลลัพธ์ใด?", c: ["อันตราย", "เร็วเกินกำหนด", "ปกติ", "อันตราย เร็วเกินกำหนด"], a: 1 },
  { q: "10. สวนสนุกให้เข้าฟรีเมื่ออายุน้อยกว่า 12 ปี (age < 12) หรืออายุตั้งแต่ 60 ปีขึ้นไป (age >= 60) ข้อใดเขียนเงื่อนไขได้ถูกต้อง?", c: ["if age < 12 or age >= 60:", "if age < 12 and age >= 60:", "if age <= 12 not age >= 60:", "if age > 12 or age < 60:"], a: 0 },

  // หน่วยคำสั่งวนซ้ำ (ข้อ 11 - 20)
  // [จุดประสงค์ 6: อธิบายหลักการทำงานของการวนซ้ำ (Loop) ได้ (2 ข้อ)]
  { q: "11. หากไม่ใช้คำสั่ง Loop ในการพิมพ์ตัวเลข 1 ถึง 100 จะส่งผลเสียอย่างไรต่อการเขียนโปรแกรม?", c: ["คอมพิวเตอร์ไม่สามารถแสดงตัวเลขเกิน 10 ได้", "ต้องเขียนคำสั่งซ้ำซ้อนจำนวนมาก แก้ไขและดูแลโค้ดได้ยาก", "ตัวแปรไม่สามารถเก็บค่าตัวเลขได้", "คำสั่ง print() จะใช้งานไม่ได้"], a: 1 },
  { q: "12. คำสั่ง continue ภายในลูปมีหน้าที่อย่างไร?", c: ["หยุดออกจากลูปทันที", "ลบตัวแปรนับรอบ", "ข้ามคำสั่งที่เหลือในรอบปัจจุบันแล้วไปเริ่มต้นตรวจสอบรอบถัดไปทันที", "ย้อนกลับไปบรรทัดแรกของโปรแกรม"], a: 2 },

  // [จุดประสงค์ 7: อธิบายความแตกต่างระหว่าง While Loop และ For Loop ได้ (2 ข้อ)]
  { q: "13. สถานการณ์ใดเหมาะสมกับการเลือกใช้ While Loop มากที่สุด?", c: ["วนรับรหัสผ่านจากผู้ใช้ซ้ำ ๆ จนกว่าจะป้อนถูกต้อง (ไม่ทราบจำนวนรอบล่วงหน้า)", "พิมพ์คำว่า Hello จำนวน 5 ครั้ง", "วนอ่านชื่อนักเรียนในลิสต์ 10 คน", "พิมพ์ตัวเลข 1 ถึง 20"], a: 0 },
  { q: "14. หากต้องการประมวลผลข้อมูลสินค้าทุกรายการที่มีอยู่แล้วใน items = ['A', 'B', 'C'] คำสั่งใดเหมาะสมและสั้นที่สุด?", c: ["while items > 0:", "if item in items:", "for item in items:", "elif item in items:"], a: 2 },

  // [จุดประสงค์ 8: เขียนโปรแกรมโดยใช้ While Loop ได้ (2 ข้อ)]
  { q: "15. พิจารณาโค้ดต่อไปนี้:\nc = 5\nwhile c > 0:\n    print(c, end=' ')\n    c -= 2\nผลลัพธ์ที่ได้คือข้อใด?", c: ["5 4 3 2 1", "5 3 1", "5 3", "3 1"], a: 1 },
  { q: "16. พิจารณาโค้ดต่อไปนี้:\ni = 0\nwhile i < 4:\n    i += 1\n    if i == 2:\n        continue\n    print(i, end=' ')\nผลลัพธ์ที่ได้คือข้อใด?", c: ["1 2 3 4", "0 1 3 4", "1 3 4", "1"], a: 2 },

  // [จุดประสงค์ 9: เขียนโปรแกรมโดยใช้ For Loop ได้ (2 ข้อ)]
  { q: "17. คำสั่ง range(2, 10, 2) ในลูป for i in range(2, 10, 2): จะสร้างตัวเลขใดบ้าง?", c: ["2, 4, 6, 8", "2, 4, 6, 8, 10", "0, 2, 4, 6, 8", "1, 3, 5, 7, 9"], a: 0 },
  { q: "18. พิจารณาโค้ดต่อไปนี้:\nfor i in range(1, 6):\n    if i == 4:\n        break\n    print(i, end=' ')\nผลลัพธ์ที่ได้คือข้อใด?", c: ["1 2 3 4 5", "1 2 3", "1 2 3 4", "4 5"], a: 1 },

  // [จุดประสงค์ 10: ประยุกต์ใช้ Loop ในการแก้ปัญหาเบื้องต้นได้ (2 ข้อ)]
  { q: "19. พิจารณาโค้ดต่อไปนี้:\nsum_even = 0\nfor x in range(1, 7):\n    if x % 2 == 0:\n        sum_even += x\nprint(sum_even)\nค่าที่พิมพ์ออกมาคือเท่าใด?", c: ["6", "10", "12", "21"], a: 2 },
  { q: "20. ออมเงินเดือนละ 200 บาท มีเริ่มต้น 100 บาท ต้องการให้ครบ 700 บาท:\nmoney = 100\nmonths = 0\nwhile money < 700:\n    money += 200\n    months += 1\nprint(months)\nค่า months ที่แสดงคือเท่าใด?", c: ["2", "3", "4", "5"], a: 1 }
];
function getPosttest(token) {
 const p=getMyProfile(token); if(!p.success)return p;
 if(p.role!=="admin") {
  const arena=getDatabase().getSheetByName("ArenaResults");
  const rows=arena.getDataRange().getValues().slice(1).filter(r=>String(r[0])===p.studentId&&(r[2]===true||String(r[2])==="TRUE"));
  const stages=new Set(rows.map(r=>Number(r[1])));
  if(stages.size<10) return {success:false,message:"ต้องผ่าน Cyber Code Arena ให้ครบทั้ง 10 ด่านก่อน"};
 }
 return {success:true,total:POSTTEST_QUESTIONS.length,maxScore:POSTTEST_QUESTIONS.length,questions:POSTTEST_QUESTIONS.map((x,i)=>({id:i,question:x.q,choices:x.c}))};
}
function submitPosttest(token, answers) {
 const p=getMyProfile(token); if(!p.success)return p;
 const total=POSTTEST_QUESTIONS.length;
 if(!Array.isArray(answers)||answers.length!==total||answers.some(x=>!Number.isInteger(x)||x<0||x>3))return {success:false,message:"กรุณาตอบให้ครบทั้ง "+total+" ข้อ"};
 if(p.role!=="admin") { const gate=getPosttest(token); if(!gate.success)return gate; }
 const correct=answers.reduce((n,a,i)=>n+(a===POSTTEST_QUESTIONS[i].a?1:0),0);
 const score=correct; // คะแนนเต็ม 20 (ข้อละ 1 คะแนน)
 const percent=Math.round((correct/total)*100);
 const minPass=Math.ceil(total*0.6);
 const passed=score>=minPass;
 getDatabase().getSheetByName(QUIZ_SHEET).appendRow([Utilities.getUuid(),p.studentId,"Post-test (เต็ม 20)",score,new Date()]);
 if(passed && p.role!=="admin") {
  const progress=getDatabase().getSheetByName(PROGRESS_SHEET), prow=findStudentProgressRow(progress,p.studentId); if(prow>1){progress.getRange(prow,11).setValue("COMPLETED");progress.getRange(prow,12).setValue("COMPLETED");}
  const sh=getDatabase().getSheetByName("PostTestStatus"), rows=sh.getDataRange().getValues();
  const idx=rows.findIndex((r,i)=>i>0&&String(r[0])===p.studentId);
  const oldBest=idx>0?Number(rows[idx][2])||0:0;
  const rec=[p.studentId,true,Math.max(oldBest,score),new Date()];
  if(idx>0) sh.getRange(idx+1,1,1,4).setValues([rec]); else sh.appendRow(rec);
 }
 return {success:true,correct,total,maxScore:total,score,percent,minPass,passed,message:passed?"ยินดีด้วย! คุณเรียนผ่านหลักสูตร PYTHON CODE ACADEMY แล้ว":"ยังไม่ผ่านเกณฑ์ 60% (ต้องได้อย่างน้อย "+minPass+"/"+total+" คะแนน) สามารถทบทวนและทำใหม่ได้"};
}

function adminGetStudentOverview(token) {
 requireAdmin_(token); ensureAcademySheets_();
 const ss=getDatabase(), students=ss.getSheetByName(STUDENTS_SHEET).getDataRange().getValues().slice(1);
 const progress=ss.getSheetByName(PROGRESS_SHEET).getDataRange().getValues().slice(1);
 const quizzes=ss.getSheetByName(QUIZ_SHEET).getDataRange().getValues().slice(1);
 const pretestSh=ss.getSheetByName("PreTestStatus");
 const pretestRows=pretestSh&&pretestSh.getLastRow()>=2?pretestSh.getDataRange().getValues().slice(1):[];
 const posttestSh=ss.getSheetByName("PostTestStatus");
 const posttestRows=posttestSh&&posttestSh.getLastRow()>=2?posttestSh.getDataRange().getValues().slice(1):[];
 return {success:true,students:students.filter(r=>r[0]).map(r=>{
  const sid=String(r[0]).trim().toLowerCase();
  const x=progress.find(v=>String(v[0]).trim().toLowerCase()===sid)||[];
  const preRow=pretestRows.find(v=>String(v[0]).trim().toLowerCase()===sid);
  const postRow=posttestRows.find(v=>String(v[0]).trim().toLowerCase()===sid);
  const preQuizRows=quizzes.filter(q=>String(q[1]).trim().toLowerCase()===sid&&String(q[2]||"").toLowerCase().includes("pre-test"));
  const pretestScore=preRow?Number(preRow[2]):(preQuizRows.length?Number(preQuizRows[preQuizRows.length-1][3]):null);
  const pretestMaxScore=preRow?(Number(preRow[3])||20):20;
  const posttestScore=postRow?Number(postRow[2]):null;
  return {id:r[0],name:r[1],status:r[3]||"ACTIVE",passwordSet:!!String(r[2]||"").trim(),pretestScore,pretestMaxScore,posttestScore,progress:{pretest:x[1]||"LOCKED",pretestScore,pretestMaxScore,posttestScore,ifElse:x[2]||"LOCKED",whileLoop:x[3]||"LOCKED",forLoop:x[4]||"LOCKED",ifElseActivity:x[6]||"LOCKED",whileActivity:x[7]||"LOCKED",forActivity:x[8]||"LOCKED",arena:x[9]||"LOCKED",posttest:x[10]||"LOCKED",course:x[11]||"LOCKED",finalQuiz:x[12]||"LOCKED"},results:quizzes.filter(q=>String(q[1]).trim().toLowerCase()===sid).map(q=>({name:q[2],score:q[3],date:q[4]}))};
 })};
}


// =========================================================
// COURSE FLOW: 3 coding challenges per lesson -> 10-stage Arena -> Post-test
// =========================================================
const LESSON_ACTIVITY_TASKS = {
  ifelse: {title:"กิจกรรมท้ายบท If-Else: คำนวณค่าจัดส่ง", prompt:"กำหนด distance = 12 และเขียน if-else: ถ้า distance มากกว่า 10 ให้พิมพ์ 'ค่าส่ง 50' มิฉะนั้นพิมพ์ 'ค่าส่ง 25'", expected:"ค่าส่ง 50"},
  while: {title:"กิจกรรมท้ายบท While Loop: สะสมแต้ม", prompt:"กำหนด points = 0 ใช้ while เพิ่ม points ทีละ 4 จน points มีค่าเท่ากับ 20 แล้วพิมพ์ points", expected:"20"},
  for: {title:"กิจกรรมท้ายบท For Loop: รวมราคาสินค้า", prompt:"กำหนด prices = [12, 18, 25, 5] ใช้ for รวมราคาและพิมพ์ผลรวม", expected:"60"}
};
const ARENA_TASKS = [
  // หมวดที่ 1: If-Else Logic (ด่าน 1 - 3)
  {
    category: "🔀 หมวด If-Else",
    title: "ตรวจสอบดัชนีคุณภาพอากาศ (AQI)",
    prompt: "กำหนดให้ค่าดัชนีคุณภาพอากาศ aqi เท่ากับ 125 จงเขียนโปรแกรมวิเคราะห์ระดับคุณภาพอากาศ:\n- ตั้งแต่ 0 ถึง 50 แสดง 'อากาศดี'\n- 51 ถึง 100 แสดง 'ปานกลาง'\n- 101 ถึง 150 แสดง 'เริ่มมีผลกระทบ'\n- มากกว่า 150 แสดง 'อันตราย'",
    expected: "เริ่มมีผลกระทบ",
    starterCode: ""
  },
  {
    category: "🔀 หมวด If-Else",
    title: "ตรวจสอบคุณสมบัติยื่นขอสินเชื่อ",
    prompt: "ผู้ยื่นขอสินเชื่อต้องมีรายได้ต่อเดือนตั้งแต่ 25,000 บาทขึ้นไป และมีอายุงานตั้งแต่ 2 ปีขึ้นไปพร้อมกัน หากกำหนดรายได้ = 28000 และอายุงาน = 3 จงเขียนโปรแกรมวิเคราะห์แสดงผลว่า 'อนุมัติสินเชื่อ' หรือ 'ไม่ผ่านเกณฑ์'",
    expected: "อนุมัติสินเชื่อ",
    starterCode: ""
  },
  {
    category: "🔀 หมวด If-Else",
    title: "ประเมินระดับการใช้พลังงานไฟฟ้า",
    prompt: "หากหน่วยการใช้ไฟฟ้าเท่ากับ 450 หน่วย จงวิเคราะห์ตามเกณฑ์: เกิน 400 หน่วยแสดง 'ค่าไฟระดับสูง', 151 ถึง 400 หน่วยแสดง 'ค่าไฟปานกลาง' และไม่เกิน 150 หน่วยแสดง 'ค่าไฟประหยัด'",
    expected: "ค่าไฟระดับสูง",
    starterCode: ""
  },

  // หมวดที่ 2: For Loop (ด่าน 4 - 6)
  {
    category: "🔁 หมวด For Loop",
    title: "แสดงพหุคูณของ 4",
    prompt: "จงเขียนโปรแกรมแสดงลำดับตัวเลขที่เป็นพหุคูณของ 4 ตั้งแต่ 4 จนถึง 20 ออกมาทีละบรรทัด",
    expected: "4\n8\n12\n16\n20",
    starterCode: ""
  },
  {
    category: "🔁 หมวด For Loop",
    title: "ผลรวมกำลังสอง 1 ถึง 4",
    prompt: "จงเขียนโปรแกรมคำนวณผลรวมของเลขยกกำลังสองตั้งแต่ 1 ถึง 4 แล้วแสดงผลลัพธ์ในรูปแบบ 'ผลรวมกำลังสอง = <ค่าที่ได้>'",
    expected: "ผลรวมกำลังสอง = 30",
    starterCode: ""
  },
  {
    category: "🔁 หมวด For Loop",
    title: "คัดกรองตัวเลขที่หาร 3 ลงตัว",
    prompt: "จงเขียนโปรแกรมวิเคราะห์และพิมพ์เฉพาะตัวเลขตั้งแต่ 1 ถึง 15 ที่หารด้วย 3 ลงตัว ออกมาทีละบรรทัด",
    expected: "3\n6\n9\n12\n15",
    starterCode: ""
  },

  // หมวดที่ 3: While Loop (ด่าน 7 - 9)
  {
    category: "⏳ หมวด While Loop",
    title: "ควบคุมการระบายน้ำในถัง",
    prompt: "ถังเก็บน้ำมีน้ำเริ่มต้น 70 ลิตร จะระบายน้ำออกครั้งละ 15 ลิตรเฉพาะเมื่อมีน้ำเหลือมากกว่า 25 ลิตร จงเขียนโปรแกรมคำนวณจำนวนครั้งที่ระบายน้ำออกได้ แล้วแสดงผลในรูปแบบ 'ระบายน้ำ = <จำนวนครั้ง> ครั้ง'",
    expected: "ระบายน้ำ = 3 ครั้ง",
    starterCode: ""
  },
  {
    category: "⏳ หมวด While Loop",
    title: "รวมเลขคู่จนถึงขอบเขตเงื่อนไข",
    prompt: "เริ่มบวกเลขคู่จาก 2, 4, 6, ... ทีละจำนวนตราบใดที่ผลรวมสะสมยังมีค่าน้อยกว่า 20 เมื่อหยุดทำงานให้แสดงผลลัพธ์ในรูปแบบ 'ผลรวม = <ค่าที่ได้>'",
    expected: "ผลรวม = 20",
    starterCode: ""
  },
  {
    category: "⏳ หมวด While Loop",
    title: "หยุดระบบเมื่อความดันเกินขีดจำกัด",
    prompt: "ระบบมีความดันเริ่มต้น 20 หน่วย และเพิ่มขึ้นรอบละ 18 หน่วย จงเขียนโปรแกรมจำลองการทำงานและสั่งหยุดทันทีเมื่อความดันมีค่าตั้งแต่ 70 หน่วยขึ้นไป จากนั้นแสดงผลในรูปแบบ 'หยุดฉุกเฉินที่ = <ค่าความดัน>'",
    expected: "หยุดฉุกเฉินที่ = 74",
    starterCode: ""
  },

  // หมวดที่ 4: Capstone Integration (ด่าน 10)
  {
    category: "🏆 ภารกิจบูรณาการ",
    title: "รวมคะแนนเฉพาะผู้ผ่านเกณฑ์",
    prompt: "กำหนดคะแนนสอบ 6 คน ได้แก่ 40, 70, 55, 80, 90 และ 50 จงเขียนโปรแกรมคัดกรองเฉพาะผู้ที่ได้คะแนนตั้งแต่ 60 ขึ้นไป แล้วคำนวณผลรวมแสดงผลในรูปแบบ 'คะแนนรวมผู้ผ่านเกณฑ์ = <ผลลัพธ์>'",
    expected: "คะแนนรวมผู้ผ่านเกณฑ์ = 240",
    starterCode: ""
  }
];
function hasPassedLessonActivity_(studentId, lessonKey) {
 const sh=getDatabase().getSheetByName("LessonActivities"); if(!sh||sh.getLastRow()<2)return false;
 return sh.getDataRange().getValues().slice(1).some(r=>String(r[0])===String(studentId)&&String(r[1])===lessonKey&&(r[2]===true||String(r[2])==="TRUE"));
}
function getLessonActivity(token, lessonKey) {
 const p=getMyProfile(token); if(!p.success)return p;
 const task=LESSON_ACTIVITY_TASKS[lessonKey]; if(!task)return {success:false,message:"ไม่พบบทเรียน"};
 if(p.role!=="admin") { const ps=getDatabase().getSheetByName(PROGRESS_SHEET),row=findStudentProgressRow(ps,p.studentId),col=({ifelse:3,while:4,for:5})[lessonKey]; if(row<2||!["UNLOCKED","COMPLETED"].includes(String(ps.getRange(row,col).getValue()).toUpperCase()))return {success:false,message:"ยังไม่ปลดล็อกบทเรียน"}; }
 return {success:true,task};
}
function submitLessonActivity(token, lessonKey, code, output) {
 const p=getMyProfile(token); if(!p.success)return p;
 const task=LESSON_ACTIVITY_TASKS[lessonKey]; if(!task)return {success:false,message:"ไม่พบบทเรียน"};
 if(p.role!=="admin"){const ps=getDatabase().getSheetByName(PROGRESS_SHEET),row=findStudentProgressRow(ps,p.studentId),col=({ifelse:3,while:4,for:5})[lessonKey];if(row<2||!["UNLOCKED","COMPLETED"].includes(String(ps.getRange(row,col).getValue()).toUpperCase()))return {success:false,message:"ยังไม่ปลดล็อกบทเรียน"};}
 code=String(code||""); output=String(output||"").replace(/\r/g,"").trim(); if(code.length<5||code.length>10000)return {success:false,message:"กรุณาเขียนโค้ดก่อนส่งตรวจ"};
 const passed=output===task.expected; const sh=getDatabase().getSheetByName("LessonActivities"); sh.appendRow([p.studentId,lessonKey,passed,passed?10:0,code,output,new Date()]);
 if(passed&&p.role!=="admin"){const ps=getDatabase().getSheetByName(PROGRESS_SHEET),row=findStudentProgressRow(ps,p.studentId),col=({ifelse:7,while:8,for:9})[lessonKey];if(row>1)ps.getRange(row,col).setValue("COMPLETED");}
 return {success:true,passed,score:passed?10:0,message:passed?"ผ่านกิจกรรมท้ายบทแล้ว! ปลดล็อก Coding Challenge":"ผลลัพธ์ยังไม่ตรง ลองแก้ไขแล้วรันอีกครั้ง"};
}
function getArenaChallenges(token) {
  const p = getMyProfile(token);
  if (!p.success) return p;
  if (p.role !== "admin") {
    const ps = getDatabase().getSheetByName(PROGRESS_SHEET), row = findStudentProgressRow(ps, p.studentId);
    if (row < 2 || !["COMPLETED", "UNLOCKED"].includes(String(ps.getRange(row, 5).getValue()).toUpperCase())) {
      return { success: false, message: "ต้องผ่านบทเรียนทั้ง 3 บทก่อน" };
    }
    const ch = getDatabase().getSheetByName("CodingChallenges").getDataRange().getValues().slice(1).filter(r => String(r[0]) === p.studentId && (r[3] === true || String(r[3]) === "TRUE"));
    const done = {};
    ch.forEach(r => done[r[1]] = (done[r[1]] || new Set()).add(Number(r[2])));
    if (!["ifelse", "while", "for"].every(k => done[k] && done[k].size >= 5)) {
      return { success: false, message: "ต้องผ่าน Coding Challenge ของทั้ง 3 บทก่อน" };
    }
  }
  const rows = getDatabase().getSheetByName("ArenaResults").getDataRange().getValues().slice(1).filter(r => String(r[0]) === p.studentId && (r[2] === true || String(r[2]) === "TRUE"));
  const passed = new Set(rows.map(r => Number(r[1])));
  return {
    success: true,
    tasks: ARENA_TASKS.map((t, i) => ({
      stage: i + 1,
      category: t.category,
      title: t.title,
      prompt: t.prompt,
      starterCode: t.starterCode || "",
      completed: passed.has(i + 1)
    })),
    completedCount: passed.size,
    score: passed.size
  };
}
function submitArenaChallenge(token, stageNo, code, output) {
  const p = getMyProfile(token);
  if (!p.success) return p;
  const n = Number(stageNo), task = ARENA_TASKS[n - 1];
  if (!task || !Number.isInteger(n)) return { success: false, message: "ด่านไม่ถูกต้อง" };
  if (p.role !== "admin") {
    const gate = getArenaChallenges(token);
    if (!gate.success) return gate;
  }
  code = String(code || "");
  output = String(output || "").replace(/\r/g, "").trim();
  if (code.length < 5 || code.length > 10000) return { success: false, message: "กรุณาเขียนโค้ดก่อนส่งตรวจ" };
  
  let passed = false;
  if (Array.isArray(task.expected)) {
    passed = task.expected.some(exp => {
      const e = String(exp).trim();
      return output === e || output.includes(e);
    });
  } else {
    const e = String(task.expected).trim();
    passed = (output === e) || output.includes(e);
  }

  const sh = getDatabase().getSheetByName("ArenaResults");
  const rows = sh.getDataRange().getValues();
  const old = rows.find((r, i) => i > 0 && String(r[0]) === p.studentId && Number(r[1]) === n && (r[2] === true || String(r[2]) === "TRUE"));
  
  if (!old || passed) {
    sh.appendRow([p.studentId, n, passed, passed ? 1 : 0, code, output, new Date()]);
  }
  
  const all = sh.getDataRange().getValues().slice(1).filter(r => String(r[0]) === p.studentId && (r[2] === true || String(r[2]) === "TRUE"));
  const count = new Set(all.map(r => Number(r[1]))).size;
  
  if (count === 10 && p.role !== "admin") {
    const ps = getDatabase().getSheetByName(PROGRESS_SHEET), row = findStudentProgressRow(ps, p.studentId);
    if (row > 1) ps.getRange(row, 10).setValue("COMPLETED");
  }
  
  return {
    success: true,
    passed: passed || !!old,
    score: passed ? 1 : (old ? 1 : 0),
    completedCount: count,
    total: 10,
    message: passed ? (count === 10 ? "🎉 ยอดเยี่ยมมาก! คุณผ่าน Cyber Code Arena ครบทั้ง 10 ด่านแล้ว ปลดล็อก Post-test เรียบร้อย" : "✓ ผ่านด่านแล้ว! เก่งมาก ลุยด่านถัดไปได้เลย") : "ผลลัพธ์ยังไม่ตรงกับโจทย์ ลองตรวจโค้ดตามตัวอย่างแล้วรันใหม่อีกครั้ง"
  };
}

function setupAcademySheets() {
  return ensureAcademySheets_();
}
function testDatabase_() {
  const ss = getDatabase();

  const requiredSheets = [
    "Students",
    "Progress",
    "QuizResults",
    "ActivityLogs"
  ];

  requiredSheets.forEach(function(name) {
    const sheet = ss.getSheetByName(name);

    if (!sheet) {
      throw new Error("ไม่พบชีต: " + name);
    }

    Logger.log("พบชีต: " + name);
  });

  Logger.log("เชื่อมต่อ Google Sheets สำเร็จ");
}
function testAdminOverview_() {
  try {
    const ss = getDatabase();

    const result = {
      database: ss.getName(),
      adminsSheet: !!ss.getSheetByName("Admins"),
      studentsSheet: !!ss.getSheetByName("Students"),
      progressSheet: !!ss.getSheetByName("Progress"),
      quizSheet: !!ss.getSheetByName("QuizResults"),
      adminRows: ss.getSheetByName("Admins").getLastRow(),
      studentRows: ss.getSheetByName("Students").getLastRow(),
      progressRows: ss.getSheetByName("Progress").getLastRow(),
      quizRows: ss.getSheetByName("QuizResults").getLastRow()
    };

    Logger.log(JSON.stringify(result, null, 2));

  } catch (e) {
    Logger.log("ERROR: " + e.message);
  }
}

// =========================================================
// NEW GATE: FINAL COURSE QUIZ (separate from Pre-test/Post-test)
// Unlock only after all 3 lessons AND all 10 Cyber Code Arena stages.
// =========================================================
const FINAL_COURSE_QUIZ = [
 {q:"1. หากต้องการตรวจสอบว่าคะแนน >= 50 และเวลาเข้าเรียน <= 8 ควรใช้ตัวดำเนินการใด",c:["or","and","not","in"],a:1},
 {q:"2. คำสั่งใดใช้เลือกทำงานเมื่อเงื่อนไข if เป็นเท็จ",c:["elif","for","else","break"],a:2},
 {q:"3. while loop จะทำงานซ้ำเมื่อใด",c:["เงื่อนไขเป็นจริง","เงื่อนไขเป็นเท็จ","ทำงานครบ 1 รอบเท่านั้น","เมื่อใช้ print()"],a:0},
 {q:"4. คำสั่งใดใช้หยุดการวนซ้ำทันที",c:["continue","break","pass","range"],a:1},
 {q:"5. range(1, 5) สร้างตัวเลขใดบ้าง",c:["1, 2, 3, 4","1, 2, 3, 4, 5","0, 1, 2, 3, 4","2, 3, 4, 5"],a:0},
 {q:"6. คำสั่ง for เหมาะกับงานลักษณะใด",c:["วนซ้ำตามสมาชิกในลำดับข้อมูล","ตรวจสอบรหัสผ่านอย่างเดียว","ประกาศคลาสเท่านั้น","หยุดโปรแกรม"],a:0},
 {q:"7. ผลลัพธ์ของ for i in range(3): print(i) คือข้อใด",c:["1 2 3","0 1 2","0 1 2 3","3 2 1"],a:1},
 {q:"8. หากต้องการข้ามคำสั่งที่เหลือในรอบปัจจุบันแล้วไปทำรอบถัดไป ใช้คำสั่งใด",c:["break","stop","continue","exit"],a:2},
 {q:"9. โครงสร้างใดเหมาะกับการตรวจสอบหลายเงื่อนไขต่อเนื่อง",c:["if-elif-else","while อย่างเดียว","print หลายบรรทัด","import"],a:0},
 {q:"10. หากต้องการประมวลผลสมาชิกทุกตัวใน list ควรเลือกใช้โครงสร้างใด",c:["if","for","break","else"],a:1}
];
function getFinalCourseQuiz(token) {
 const p=getMyProfile(token); if(!p.success)return p;
 if(p.role!=="admin") {
  const ps=getDatabase().getSheetByName(PROGRESS_SHEET), row=findStudentProgressRow(ps,p.studentId);
  if(row<2)return {success:false,message:"ไม่พบข้อมูลความก้าวหน้า"};
  const vals=ps.getRange(row,3,1,11).getValues()[0];
  const lessons=vals.slice(0,3).every(v=>String(v).toUpperCase()==="COMPLETED");
  const arena=String(vals[7]||"").toUpperCase()==="COMPLETED";
  if(!lessons||!arena)return {success:false,message:"ด่านนี้จะปลดล็อกเมื่อผ่าน If-Else, While Loop, For Loop และ Cyber Code Arena ครบทั้ง 10 ด่าน"};
 }
 return {success:true,questions:FINAL_COURSE_QUIZ.map((x,i)=>({id:i,question:x.q,choices:x.c}))};
}
function submitFinalCourseQuiz(token, answers) {
 const p=getMyProfile(token); if(!p.success)return p;
 if(!Array.isArray(answers)||answers.length!==FINAL_COURSE_QUIZ.length)return {success:false,message:"กรุณาตอบคำถามให้ครบทั้ง 10 ข้อ"};
 if(p.role!=="admin") {
  const ps=getDatabase().getSheetByName(PROGRESS_SHEET), row=findStudentProgressRow(ps,p.studentId);
  if(row<2)return {success:false,message:"ไม่พบข้อมูลความก้าวหน้า"};
  const vals=ps.getRange(row,3,1,8).getValues()[0];
  if(!vals.slice(0,3).every(v=>String(v).toUpperCase()==="COMPLETED")||String(vals[7]||"").toUpperCase()!=="COMPLETED")return {success:false,message:"ยังผ่านเงื่อนไขปลดล็อกด่านไม่ครบ"};
 }
 let score=0; FINAL_COURSE_QUIZ.forEach((q,i)=>{if(Number(answers[i])===q.a)score++;});
 const sh=getDatabase().getSheetByName(QUIZ_SHEET); sh.appendRow([Utilities.getUuid(),p.studentId,"Final Course Quiz",score*10,new Date()]);
 if(p.role!=="admin" && score>=7){const ps=getDatabase().getSheetByName(PROGRESS_SHEET),row=findStudentProgressRow(ps,p.studentId);if(row>1)ps.getRange(row,13).setValue("COMPLETED");}
 return {success:true,score:score*10,correct:score,total:10,passed:score>=7,message:score>=7?"ผ่านด่านแบบทดสอบท้ายหลักสูตรแล้ว!":"ยังไม่ผ่านเกณฑ์ 70% สามารถกลับไปทบทวนและทำใหม่ได้"};
}
