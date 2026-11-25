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

  return {
    success: true,
    student: {
      studentId: profile.studentId,
      studentName: profile.studentName
    },
    progress: {
      pretestStatus: progress[1],
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
// PRE-TEST
// ==========================================
const PRETEST_QUESTIONS = [
  { question: "1. Python เป็นภาษาโปรแกรมประเภทใด?", choices: ["ภาษาเครื่อง (Machine Language)", "ภาษาระดับสูง (High-level Language)", "ภาษาโครงสร้างเว็บ (HTML)", "ภาษาจัดการฐานข้อมูล (SQL)"], answer: 1 },
  { question: "2. คำสั่งใดใช้สำหรับแสดงผลข้อความหรือข้อมูลออกทางหน้าจอ?", choices: ["input()", "show()", "print()", "display()"], answer: 2 },
  { question: "3. ข้อใดเป็นการตั้งชื่อตัวแปรที่ถูกต้องตามหลักไวยากรณ์ของ Python?", choices: ["2score = 100", "student_name = 'Alex'", "class = 10", "my-score = 80"], answer: 1 },
  { question: "4. ชนิดข้อมูลใดในภาษา Python ใช้สำหรับจัดเก็บข้อความตัวอักษร?", choices: ["int", "float", "str", "bool"], answer: 2 },
  { question: "5. เครื่องหมายใดใช้สำหรับตรวจสอบว่าค่าสองค่ามีค่า 'เท่ากัน' หรือไม่?", choices: ["=", "==", "!=", ">="], answer: 1 },
  { question: "6. คำสั่ง if ในภาษา Python ใช้สำหรับวัตถุประสงค์ใด?", choices: ["วนซ้ำการทำงานเรื่อย ๆ", "ตรวจสอบเงื่อนไขเพื่อตัดสินใจทำงาน", "รับข้อมูลจากแป้นพิมพ์", "ลบตัวแปรออกจากหน่วยความจำ"], answer: 1 },
  { question: "7. ผลลัพธ์จากการประเมินค่านิพจน์ทางตรรกะ 10 >= 5 คือข้อใด?", choices: ["10", "5", "True", "False"], answer: 2 },
  { question: "8. คำสั่งใดในภาษา Python ใช้สำหรับรับข้อมูลที่ผู้ใช้พิมพ์เข้ามาทางแป้นพิมพ์?", choices: ["print()", "input()", "read()", "scan()"], answer: 1 },
  { question: "9. ผลลัพธ์ของนิพจน์ 7 % 3 (ตัวดำเนินการ Modulo) มีค่าเท่ากับข้อใด?", choices: ["2", "1", "2.33", "0"], answer: 1 },
  { question: "10. คำสั่ง while ในภาษา Python มีหลักการทำงานอย่างไร?", choices: ["วนซ้ำการทำงานตราบใดที่เงื่อนไขยังคงเป็นจริง", "วนซ้ำเพียงครั้งเดียวแล้วหยุดเสมอ", "สร้างฟังก์ชันใหม่ขึ้นมาใช้งาน", "หยุดการทำงานของโปรแกรมทันที"], answer: 0 },
  { question: "11. คำสั่งใดใช้แปลงข้อความตัวเลข '25' ให้กลายเป็นข้อมูลชนิดตัวเลขจำนวนเต็ม?", choices: ["int('25')", "float('25')", "str(25)", "bool('25')"], answer: 0 },
  { question: "12. คำสั่ง for ในภาษา Python เหมาะสำหรับการทำงานในลักษณะใดมากที่สุด?", choices: ["ทำงานเมื่อเงื่อนไขเป็นเท็จเท่านั้น", "วนซ้ำตามสมาชิกในลำดับข้อมูล เช่น range หรือ list", "รับค่ารหัสผ่านจากผู้ใช้", "หยุดการประมวลผลของคอมพิวเตอร์"], answer: 1 },
  { question: "13. ฟังก์ชัน range(4) ในลูป for i in range(4): จะสร้างลำดับตัวเลขใด?", choices: ["1, 2, 3, 4", "0, 1, 2, 3", "0, 1, 2, 3, 4", "1, 2, 3"], answer: 1 },
  { question: "14. สัญลักษณ์ใดในภาษา Python ใช้สำหรับเขียนข้อความหมายเหตุ (Comment) บรรทัดเดียว?", choices: ["//", "/* */", "#", "--"], answer: 2 },
  { question: "15. ข้อใดคือประโยชน์หลักของการใช้การวนซ้ำ (Loop) ในการเขียนโปรแกรม?", choices: ["ช่วยลดการเขียนโค้ดซ้ำซ้อนและสั่งให้ทำงานซ้ำได้อัตโนมัติ", "ช่วยแปลงโค้ด Python ให้เป็นภาษาเครื่องทันที", "ใช้ตกแต่งสีสันของหน้าต่างโปรแกรม", "ใช้ป้องกันไม่ให้เกิดข้อผิดพลาดในการเขียนโค้ด"], answer: 0 }
];

function getPretestQuestions(token) {
  const profile = getMyProfile(token);
  if (!profile || !profile.success) return { success: false, message: "Session หมดอายุ" };
  return {
    success: true,
    questions: PRETEST_QUESTIONS.map((item, index) => ({ id: index, question: item.question, choices: item.choices }))
  };
}

function submitPretest(token, answers) {
  const profile = getMyProfile(token);
  if (!profile || !profile.success) return { success: false, message: "Session หมดอายุ" };

  let correct = 0;
  PRETEST_QUESTIONS.forEach((item, index) => {
    if (answers[index] === item.answer) correct++;
  });
  const score = Math.round((correct / PRETEST_QUESTIONS.length) * 100);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getDatabase();
    const quizSheet = ss.getSheetByName(QUIZ_SHEET);
    const progressSheet = ss.getSheetByName(PROGRESS_SHEET);

    quizSheet.appendRow([Utilities.getUuid(), profile.studentId, "Pre-test", score, new Date()]);

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
      correct: correct,
      total: PRETEST_QUESTIONS.length,
      message: "บันทึกคะแนน Pre-test เรียบร้อยแล้ว"
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

  // 2. ตรวจสอบจาก QuizResults Sheet (Score >= 60%)
  try {
    const qs = getDatabase().getSheetByName(QUIZ_SHEET);
    if (qs && qs.getLastRow() >= 2) {
      const data = qs.getDataRange().getValues().slice(1);
      const title = config.title;
      const passed = data.some(r =>
        String(r[1]).trim().toLowerCase() === String(studentId).trim().toLowerCase() &&
        (String(r[2]).includes(title) || String(r[2]).toLowerCase().includes(String(lessonKey).toLowerCase())) &&
        Number(r[3]) >= 60
      );
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

  const quizPassed = profile.role === "admin" || hasPassedLessonQuiz_(profile.studentId, lessonKey);

  return {
    success: true,
    lessonKey: lessonKey,
    quizPassed: quizPassed,
    questions: config.questions.map((item, index) => ({ id: index, question: item.question, choices: item.choices })),
    totalQuestions: config.questions.length,
    minPassQuestions: Math.ceil(config.questions.length * 0.6),
    passScore: 60
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

  const questions = config.questions;
  if (!Array.isArray(answers) || answers.length !== questions.length || answers.some(a => !Number.isInteger(a) || a < 0 || a > 3)) {
    return { success: false, message: "รูปแบบคำตอบไม่ถูกต้อง กรุณาตอบให้ครบทุกข้อ (มีทั้งหมด " + questions.length + " ข้อ)" };
  }

  let correct = 0;
  answers.forEach((ans, idx) => {
    if (ans === questions[idx].answer) correct++;
  });
  const score = Math.round((correct / questions.length) * 100);
  const minPassQuestions = Math.ceil(questions.length * 0.6);
  const passed = score >= 60;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getDatabase();
    const progressSheet = ss.getSheetByName(PROGRESS_SHEET);
    const quizSheet = ss.getSheetByName(QUIZ_SHEET);

    const rowNumber = findStudentProgressRow(progressSheet, profile.studentId);
    quizSheet.appendRow([Utilities.getUuid(), profile.studentId, config.title + " Quiz", score, new Date()]);

    // เมื่อสอบผ่านเกณฑ์ 60% ขึ้นไป -> อัปเดต QuizColumn ให้เป็น COMPLETED เพื่อปลดล็อกโจทย์ปฏิบัติ 5 ข้อ
    if (passed && rowNumber > 1 && config.quizColumn && profile.role !== "admin") {
      progressSheet.getRange(rowNumber, config.quizColumn).setValue("COMPLETED");
    }
    SpreadsheetApp.flush();

    return {
      success: true,
      score: score,
      correct: correct,
      total: questions.length,
      minPassQuestions: minPassQuestions,
      passed: passed,
      unlockedPractical: passed,
      message: passed
        ? "🎉 ยินดีด้วย! คุณผ่านแบบทดสอบท้ายบทแล้ว (" + correct + "/" + questions.length + " ข้อ หรือ " + score + "%) ปลดล็อกโจทย์เขียนโค้ดปฏิบัติ 5 ข้อเรียบร้อยแล้ว!"
        : "⚠️ ยังไม่ผ่านเกณฑ์ 60% (ได้ " + correct + "/" + questions.length + " ข้อ หรือ " + score + "%) ต้องการอย่างน้อย " + minPassQuestions + " ข้อ กรุณาทบทวนบทเรียนและทำแบบทดสอบอีกครั้งเพื่อปลดล็อกภาคปฏิบัติ"
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
// CODING CHALLENGES: 3 tasks per lesson, score total 10 / lesson
// Client runs actual Python with Pyodide; server checks exact expected output.
// =========================================================
const CODING_TASKS = {
  ifelse: [
    {
      title: "ตรวจสอบสิทธิ์เลือกตั้ง",
      prompt: "จงเขียนโปรแกรมรับค่าอายุ (จำนวนเต็ม) จากผู้ใช้ 1 ค่า\n• หากอายุตั้งแต่ 18 ปีขึ้นไป ให้แสดงข้อความ: มีสิทธิ์เลือกตั้ง\n• หากอายุน้อยกว่า 18 ปี ให้แสดงข้อความ: ยังไม่มีสิทธิ์เลือกตั้ง\n(ข้อมูลทดสอบ: ป้อน 20)",
      expected: "มีสิทธิ์เลือกตั้ง",
      points: 2,
      defaultInputs: ["20"]
    },
    {
      title: "ตรวจสอบจำนวนบวก ลบ หรือศูนย์",
      prompt: "จงเขียนโปรแกรมรับค่าตัวเลขจำนวนเต็ม 1 ค่า แล้วจำแนกประเภทของตัวเลขดังนี้:\n• มากกว่า 0 ให้แสดง: จำนวนบวก\n• น้อยกว่า 0 ให้แสดง: จำนวนลบ\n• เท่ากับ 0 ให้แสดง: ศูนย์\n(ข้อมูลทดสอบ: ป้อน 15)",
      expected: "จำนวนบวก",
      points: 2,
      defaultInputs: ["15"]
    },
    {
      title: "ระบบตัดเกรดวัดผล",
      prompt: "จงเขียนโปรแกรมรับคะแนนสอบ (รองรับทศนิยม) จากผู้ใช้ แล้วประเมินเกรดตามเกณฑ์:\n• 80 คะแนนขึ้นไป: เกรด A\n• 70 - 79.99 คะแนน: เกรด B\n• 60 - 69.99 คะแนน: เกรด C\n• 50 - 59.99 คะแนน: เกรด D\n• ต่ำกว่า 50 คะแนน: เกรด F\n(ข้อมูลทดสอบ: ป้อน 85)",
      expected: "เกรด A",
      points: 2,
      defaultInputs: ["85"]
    },
    {
      title: "ตรวจสอบเงื่อนไขซื้อสินค้า",
      prompt: "จงเขียนโปรแกรมรับข้อมูล 2 ค่าตามลำดับ คือ อายุ และ จำนวนเงินที่มี (จำนวนเต็ม)\nเงื่อนไขการซื้อสินค้า: ต้องมีอายุตั้งแต่ 18 ปีขึ้นไป และมีเงินตั้งแต่ 100 บาทขึ้นไปพร้อมกัน\n• หากผ่านเกณฑ์ทั้งคู่ ให้แสดง: สามารถซื้อสินค้าได้\n• หากไม่ผ่านเกณฑ์ ให้แสดง: ไม่สามารถซื้อสินค้าได้\n(ข้อมูลทดสอบ: ป้อน 20 และ 150)",
      expected: "สามารถซื้อสินค้าได้",
      points: 2,
      defaultInputs: ["20", "150"]
    },
    {
      title: "ระบบตรวจสอบรหัสผ่าน",
      prompt: "จงเขียนโปรแกรมรับรหัสผ่านจากผู้ใช้\n• หากรหัสผ่านตรงกับ \"1234\" ให้แสดง: เข้าสู่ระบบสำเร็จ\n• หากไม่ตรง ให้แสดง: รหัสผ่านไม่ถูกต้อง\n(ข้อมูลทดสอบ: ป้อน 1234)",
      expected: "เข้าสู่ระบบสำเร็จ",
      points: 2,
      defaultInputs: ["1234"]
    }
  ],
  for: [
    {
      title: "แสดงตัวเลข 1 ถึง N",
      prompt: "จงเขียนโปรแกรมรับจำนวนเต็ม N จากผู้ใช้ แล้วแสดงลำดับตัวเลขตั้งแต่ 1 จนถึง N ออกมาทีละบรรทัด\n(ข้อมูลทดสอบ: ป้อน 5)",
      expected: "1\n2\n3\n4\n5",
      points: 2,
      defaultInputs: ["5"]
    },
    {
      title: "แสดงเฉพาะเลขคู่ 2 ถึง 10",
      prompt: "จงเขียนโปรแกรมแสดงเฉพาะเลขคู่ตั้งแต่ 2 ถึง 10 (ได้แก่ 2, 4, 6, 8, 10) ออกมาทางหน้าจอทีละบรรทัด",
      expected: "2\n4\n6\n8\n10",
      points: 2,
      defaultInputs: []
    },
    {
      title: "คำนวณผลรวมสะสม 1 ถึง 5",
      prompt: "จงเขียนโปรแกรมคำนวณผลรวมของตัวเลขตั้งแต่ 1 ถึง 5 แล้วแสดงผลลัพธ์ในรูปแบบ:\nผลรวม = 15",
      expected: "ผลรวม = 15",
      points: 2,
      defaultInputs: []
    },
    {
      title: "โปรแกรมสูตรคูณแม่ N",
      prompt: "จงเขียนโปรแกรมรับแม่สูตรคูณ N จากผู้ใช้ แล้วแสดงตารางสูตรคูณ 5 แถวแรก (คูณ 1 ถึง 5)\nตัวอย่างผลลัพธ์เมื่อป้อน 3:\n3 x 1 = 3\n3 x 2 = 6\n... จนถึง 3 x 5 = 15",
      expected: "3 x 1 = 3\n3 x 2 = 6\n3 x 3 = 9\n3 x 4 = 12\n3 x 5 = 15",
      points: 2,
      defaultInputs: ["3"]
    },
    {
      title: "คำนวณผลรวมคะแนนในลิสต์",
      prompt: "กำหนดให้มีชุดข้อมูลคะแนน 4 ค่า คือ 10, 20, 30, 40\nจงเขียนโปรแกรมคำนวณผลรวมของคะแนนทั้งหมด แล้วแสดงผลลัพธ์ในรูปแบบ:\nคะแนนรวม = 100",
      expected: "คะแนนรวม = 100",
      points: 2,
      defaultInputs: []
    }
  ],
  while: [
    {
      title: "นับถอยหลังปล่อยยานอวกาศ",
      prompt: "จงเขียนโปรแกรมวนซ้ำแบบตรวจสอบเงื่อนไขเพื่อนับถอยหลังตั้งแต่ 3, 2, 1 ทีละบรรทัด และเมื่อนับจบแล้วให้แสดงข้อความ: เริ่มการทำงาน!",
      expected: ["3\n2\n1\nเริ่มการทำงาน! 🚀", "3\n2\n1\nเริ่มการทำงาน!", "3\n2\n1\nเริ่ม!"],
      points: 2,
      defaultInputs: []
    },
    {
      title: "รับตัวเลขจนกว่าจะป้อนเลข 0",
      prompt: "จงเขียนโปรแกรมรับค่าตัวเลขจำนวนเต็มจากผู้ใช้ซ้ำ ๆ ตราบใดที่ยังไม่ได้ป้อนเลข 0\n• หากไม่ใช่ 0 ให้แสดง: คุณกรอกเลข <ค่าที่กรอก> แล้วรับค่าใหม่\n• เมื่อผู้ใช้ป้อน 0 ให้หยุดการทำงานและแสดง: หยุดการทำงานแล้ว\n(ข้อมูลทดสอบ: ป้อน 5 แล้วตามด้วย 0)",
      expected: ["คุณกรอกเลข 5\nหยุดการทำงานแล้ว", "คุณกรอกเลข: 5\nหยุดการทำงานแล้ว"],
      points: 2,
      defaultInputs: ["5", "0"]
    },
    {
      title: "คำนวณผลรวม 1 ถึง 5 ด้วย While Loop",
      prompt: "จงเขียนโปรแกรมคำนวณหาผลรวมของตัวเลขตั้งแต่ 1 ถึง 5 โดยใช้การวนซ้ำแบบตรวจสอบเงื่อนไข แล้วแสดงผลลัพธ์ในรูปแบบ:\nผลรวม = 15",
      expected: "ผลรวม = 15",
      points: 2,
      defaultInputs: []
    },
    {
      title: "แยกเลขคู่และเลขคี่ 1 ถึง 5",
      prompt: "จงเขียนโปรแกรมตรวจสอบตัวเลขตั้งแต่ 1 ถึง 5 ทีละตัวว่าเป็นเลขคู่หรือเลขคี่ แล้วแสดงผลทีละบรรทัดในรูปแบบ:\n1 เป็นเลขคี่\n2 เป็นเลขคู่\n... จนถึง 5",
      expected: "1 เป็นเลขคี่\n2 เป็นเลขคู่\n3 เป็นเลขคี่\n4 เป็นเลขคู่\n5 เป็นเลขคี่",
      points: 2,
      defaultInputs: []
    },
    {
      title: "ระบบรับรหัสผ่านจนกว่าจะถูกต้อง",
      prompt: "จงเขียนโปรแกรมตรวจสอบรหัสผ่าน โดยกำหนดรหัสที่ถูกต้องคือ \"1234\"\n• หากผู้ใช้กรอกรหัสผิด ให้แสดงว่า \"รหัสผ่านไม่ถูกต้อง\" และวนรับรหัสใหม่\n• หากกรอกถูกต้อง ให้หยุดวนซ้ำและแสดง: เข้าสู่ระบบสำเร็จ\n(ข้อมูลทดสอบ: ป้อน 1234)",
      expected: "เข้าสู่ระบบสำเร็จ",
      points: 2,
      defaultInputs: ["1234"]
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
  const quizPassed = profile.role === "admin" || hasPassedLessonQuiz_(profile.studentId, lessonKey);
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
      defaultInputs: t.defaultInputs || [],
      completed: !!completed[i + 1],
      score: completed[i + 1]?.score || 0,
      submittedCode: completed[i + 1]?.code || ""
    })),
    completedCount: Object.keys(completed).length,
    totalScore,
    totalPoints,
    quizPassed
  };
}

function submitCodingChallenge(token, lessonKey, challengeNo, code, output) {
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
    if (!hasPassedLessonQuiz_(profile.studentId, lessonKey)) {
      return { success: false, message: "กรุณาทำแบบทดสอบท้ายบทให้ผ่านเกณฑ์ 60% ก่อนส่งโจทย์ปฏิบัติ" };
    }
  }
  code = String(code || ""); output = String(output || "").replace(/\r/g,"").trim();
  if (code.length < 5 || code.length > 10000) return {success:false,message:"กรุณาเขียนโค้ดให้ครบก่อนส่งตรวจ"};
  
  const expected = tasks[n - 1].expected;
  let passed = false;
  if (Array.isArray(expected)) {
    passed = expected.some(exp => {
      const e = String(exp).trim();
      return output === e || output.includes(e);
    });
  } else {
    const e = String(expected).trim();
    passed = (output === e) || output.includes(e);
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

  if (passedCount === tasks.length && profile.role !== "admin") {
    const ps = getDatabase().getSheetByName(PROGRESS_SHEET), row = findStudentProgressRow(ps, profile.studentId);
    const col = ({ifelse:3, for:5, while:4})[lessonKey];
    const next = ({ifelse:5, for:4, while:6})[lessonKey];
    if (row > 1) {
      ps.getRange(row, col).setValue("COMPLETED");
      if (lessonKey !== "while") {
        ps.getRange(row, next).setValue("UNLOCKED");
      } else {
        ps.getRange(row, 6).setValue("UNLOCKED");
        ps.getRange(row, 10).setValue("UNLOCKED");
      }
    }
  }

  return {
    success: true,
    passed: finalPassed,
    score: finalScore,
    points: tasks[n - 1].points,
    passedCount,
    lessonScore,
    lessonTotal,
    unlockedNext: passedCount === tasks.length,
    message: finalPassed ? (passedCount === tasks.length ? "ผ่านครบทั้ง 5 ข้อแล้ว! ปลดล็อกขั้นตอนถัดไปเรียบร้อย" : "รันผ่านแล้ว! ไปทำข้อถัดไปได้เลย") : "ผลลัพธ์ยังไม่ตรงกับโจทย์ ลองตรวจโค้ดหรือป้อนค่าตามที่โจทย์แนะนำแล้วรันอีกครั้ง"
  };
}

const POSTTEST_QUESTIONS = [
 {q:"คำสั่งใดใช้ตรวจสอบเงื่อนไขใน Python?",c:["if","for","print","input"],a:0},
 {q:"ถ้า score = 45 และเงื่อนไข score >= 50 จะได้ค่าใด?",c:["True","False","45","None"],a:1},
 {q:"คำสั่ง elif ใช้เพื่ออะไร?",c:["ทำซ้ำ","ตรวจเงื่อนไขเพิ่มเติม","หยุดโปรแกรม","รับข้อมูล"],a:1},
 {q:"while loop ทำงานต่อเมื่อใด?",c:["เงื่อนไขเป็นจริง","เงื่อนไขเป็นเท็จ","ครบ 10 รอบเสมอ","มีคำสั่ง print เท่านั้น"],a:0},
 {q:"ถ้าไม่เปลี่ยนค่าตัวแปรควบคุมใน while ที่เงื่อนไขยังจริง จะเกิดอะไรได้?",c:["ลูปไม่ทำงาน","ลูปไม่สิ้นสุด","โปรแกรมแปลงชนิดข้อมูล","ทำงานครั้งเดียว"],a:1},
 {q:"range(4) สร้างค่าใด?",c:["1 ถึง 4","0 ถึง 3","0 ถึง 4","4 เท่านั้น"],a:1},
 {q:"for x in [2,4,6] ทำงานกี่รอบ?",c:["2","3","4","6"],a:1},
 {q:"ผลรวม 1 ถึง 4 เท่ากับเท่าใด?",c:["8","9","10","11"],a:2},
 {q:"ต้องการตรวจว่า x อยู่ระหว่าง 10 ถึง 20 รวมขอบเขต ควรใช้เงื่อนไขใด?",c:["x > 10 and x < 20","x >= 10 and x <= 20","x == 10 or 20","x < 10 and x > 20"],a:1},
 {q:"หากต้องการทำงานกับสมาชิกทุกตัวใน list ควรเลือกใช้คำสั่งใด?",c:["if","else","for","break"],a:2}
];
function getPosttest(token) {
 const p=getMyProfile(token); if(!p.success)return p;
 if(p.role!=="admin") {
  const arena=getDatabase().getSheetByName("ArenaResults");
  const rows=arena.getDataRange().getValues().slice(1).filter(r=>String(r[0])===p.studentId&&(r[2]===true||String(r[2])==="TRUE"));
  const stages=new Set(rows.map(r=>Number(r[1])));
  if(stages.size<10) return {success:false,message:"ต้องผ่าน Cyber Code Arena ให้ครบทั้ง 10 ด่านก่อน"};
 }
 return {success:true,questions:POSTTEST_QUESTIONS.map((x,i)=>({id:i,question:x.q,choices:x.c}))};
}
function submitPosttest(token, answers) {
 const p=getMyProfile(token); if(!p.success)return p;
 if(!Array.isArray(answers)||answers.length!==10||answers.some(x=>!Number.isInteger(x)||x<0||x>3))return {success:false,message:"กรุณาตอบให้ครบทั้ง 10 ข้อ"};
 if(p.role!=="admin") { const gate=getPosttest(token); if(!gate.success)return gate; }
 const correct=answers.reduce((n,a,i)=>n+(a===POSTTEST_QUESTIONS[i].a?1:0),0), score=correct*10, passed=correct>=6;
 getDatabase().getSheetByName(QUIZ_SHEET).appendRow([Utilities.getUuid(),p.studentId,"Post-test Final",score,new Date()]);
 if(passed && p.role!=="admin") {
  const progress=getDatabase().getSheetByName(PROGRESS_SHEET), prow=findStudentProgressRow(progress,p.studentId); if(prow>1){progress.getRange(prow,11).setValue("COMPLETED");progress.getRange(prow,12).setValue("COMPLETED");}
  const sh=getDatabase().getSheetByName("PostTestStatus"), rows=sh.getDataRange().getValues();
  const idx=rows.findIndex((r,i)=>i>0&&String(r[0])===p.studentId);
  const oldBest=idx>0?Number(rows[idx][2])||0:0;
  const rec=[p.studentId,true,Math.max(oldBest,score),new Date()];
  if(idx>0) sh.getRange(idx+1,1,1,4).setValues([rec]); else sh.appendRow(rec);
 }
 return {success:true,correct,total:10,score,passed,message:passed?"ยินดีด้วย! คุณเรียนผ่านหลักสูตร PYTHON CODE ACADEMY แล้ว":"ยังไม่ผ่านเกณฑ์ 60 คะแนน สามารถทบทวนและทำใหม่ได้"};
}

function adminGetStudentOverview(token) {
 requireAdmin_(token); ensureAcademySheets_();
 const ss=getDatabase(), students=ss.getSheetByName(STUDENTS_SHEET).getDataRange().getValues().slice(1);
 const progress=ss.getSheetByName(PROGRESS_SHEET).getDataRange().getValues().slice(1);
 const quizzes=ss.getSheetByName(QUIZ_SHEET).getDataRange().getValues().slice(1);
 return {success:true,students:students.filter(r=>r[0]).map(r=>{const x=progress.find(v=>String(v[0]).trim().toLowerCase()===String(r[0]).trim().toLowerCase())||[];return {id:r[0],name:r[1],status:r[3]||"ACTIVE",passwordSet:!!String(r[2]||"").trim(),progress:{pretest:x[1]||"LOCKED",ifElse:x[2]||"LOCKED",whileLoop:x[3]||"LOCKED",forLoop:x[4]||"LOCKED",ifElseActivity:x[6]||"LOCKED",whileActivity:x[7]||"LOCKED",forActivity:x[8]||"LOCKED",arena:x[9]||"LOCKED",posttest:x[10]||"LOCKED",course:x[11]||"LOCKED",finalQuiz:x[12]||"LOCKED"},results:quizzes.filter(q=>String(q[1]).trim().toLowerCase()===String(r[0]).trim().toLowerCase()).map(q=>({name:q[2],score:q[3],date:q[4]}))};})};
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
    title: "ตรวจสอบอุณหภูมิห้องแล็บ",
    prompt: "กำหนด temperature = 32 ให้เขียน if-elif-else ตรวจสอบสภาพอากาศ:\n- ถ้า temperature >= 35 แสดง 'อากาศร้อนมาก'\n- ถ้า temperature >= 30 แสดง 'อากาศร้อน'\n- ถ้า temperature >= 25 แสดง 'อากาศอบอุ่น'\n- มิฉะนั้น (else) แสดง 'อากาศเย็น'",
    expected: "อากาศร้อน",
    starterCode: "temperature = 32\n\n# เขียน if-elif-else ตรวจสอบอุณหภูมิ\n"
  },
  {
    category: "🔀 หมวด If-Else",
    title: "ตรวจสอบคุณสมบัติผู้เข้าแข่งขัน",
    prompt: "กำหนด age = 19 และ level = 2 ให้เขียน if-else ร่วมกับ and ตรวจสอบว่า:\nถ้า age >= 18 and level >= 2 ให้แสดง 'ผ่านคุณสมบัติ' มิฉะนั้นแสดง 'ไม่ผ่านคุณสมบัติ'",
    expected: "ผ่านคุณสมบัติ",
    starterCode: "age = 19\nlevel = 2\n\n# เขียน if-else ตรวจสอบเงื่อนไขคู่ด้วย and\n"
  },
  {
    category: "🔀 หมวด If-Else",
    title: "ระบบคำนวณตัดเกรดจำแนกระดับ",
    prompt: "กำหนด score = 75 ให้เขียน if-elif-else ตัดเกรด:\n- score >= 80 แสดง 'เกรด A'\n- score >= 70 แสดง 'เกรด B'\n- score >= 60 แสดง 'เกรด C'\n- score >= 50 แสดง 'เกรด D'\n- else แสดง 'เกรด F'",
    expected: "เกรด B",
    starterCode: "score = 75\n\n# เขียน if-elif-else ตัดเกรดจากตัวแปร score\n"
  },

  // หมวดที่ 2: For Loop (ด่าน 4 - 6)
  {
    category: "🔁 หมวด For Loop",
    title: "สร้างลำดับตัวเลขก้าวละ 2 ด้วย range()",
    prompt: "ใช้คำสั่ง for ร่วมกับ range(1, 10, 2) เพื่อแสดงตัวเลข 1, 3, 5, 7, 9 ออกมาทีละบรรทัดด้วย print(i)",
    expected: "1\n3\n5\n7\n9",
    starterCode: "# ใช้ for ร่วมกับ range(1, 10, 2) พิมพ์ตัวเลขทีละบรรทัด\n"
  },
  {
    category: "🔁 หมวด For Loop",
    title: "คำนวณผลรวมสะสม 1 ถึง 10",
    prompt: "กำหนด total = 0 ใช้ for i in range(1, 11): บวกสะสมค่า total = total + i เมื่อจบลูปสั่ง print(\"ผลรวม =\", total) (ต้องได้ผลลัพธ์: ผลรวม = 55)",
    expected: "ผลรวม = 55",
    starterCode: "total = 0\n\n# ใช้ for วนซ้ำ 1 ถึง 10 บวกสะสมเข้า total แล้วพิมพ์ 'ผลรวม = 55'\n"
  },
  {
    category: "🔁 หมวด For Loop",
    title: "คัดกรองพิมพ์เฉพาะเลขคู่ด้วย For + If",
    prompt: "ใช้ for i in range(1, 11): ร่วมกับเงื่อนไข if i % 2 == 0: เพื่อพิมพ์เฉพาะเลขคู่ 2, 4, 6, 8, 10 ออกมาทีละบรรทัด",
    expected: "2\n4\n6\n8\n10",
    starterCode: "# ใช้ for 1 ถึง 10 และ if ตรวจสอบหาร 2 ลงตัว เพื่อพิมพ์เฉพาะเลขคู่\n"
  },

  // หมวดที่ 3: While Loop (ด่าน 7 - 9)
  {
    category: "⏳ หมวด While Loop",
    title: "ระบบนับถอยหลังสู่อวกาศ",
    prompt: "กำหนด countdown = 5 ใช้ while countdown > 0: พิมพ์ค่า countdown และลดค่า countdown -= 1 ทีละรอบ เมื่อจบลูปสั่ง print(\"เริ่มภารกิจสำเร็จ! 🚀\")",
    expected: ["5\n4\n3\n2\n1\nเริ่มภารกิจสำเร็จ! 🚀", "5\n4\n3\n2\n1\nเริ่มภารกิจสำเร็จ!"],
    starterCode: "countdown = 5\n\n# เขียน while loop นับถอยหลัง 5 ถึง 1 และพิมพ์ 'เริ่มภารกิจสำเร็จ! 🚀'\n"
  },
  {
    category: "⏳ หมวด While Loop",
    title: "หาผลรวมสะสมด้วย While Loop",
    prompt: "กำหนด i = 1 และ total = 0 ใช้ while i <= 6: นำค่าบวกสะสม total = total + i และเพิ่ม i += 1 ในแต่ละรอบ เมื่อจบลูปสั่ง print(\"ผลรวม =\", total) (ต้องได้ผลลัพธ์: ผลรวม = 21)",
    expected: "ผลรวม = 21",
    starterCode: "i = 1\ntotal = 0\n\n# ใช้ while loop หาผลรวม 1 ถึง 6 แล้วพิมพ์ 'ผลรวม = 21'\n"
  },
  {
    category: "⏳ หมวด While Loop",
    title: "ควบคุมการออกจากลูปฉุกเฉินด้วย break",
    prompt: "กำหนด power = 10 ในลูป while True: ให้เพิ่ม power += 15 และเขียน if power >= 50: สั่ง break เมื่อจบลูปให้พิมพ์ 'พลังงานปลอดภัย = ' ตามด้วยค่า power (ต้องได้: พลังงานปลอดภัย = 55)",
    expected: "พลังงานปลอดภัย = 55",
    starterCode: "power = 10\n\nwhile True:\n    power += 15\n    # เขียน if ตรวจสอบ power >= 50 แล้วสั่ง break\n    \n\nprint(\"พลังงานปลอดภัย =\", power)\n"
  },

  // หมวดที่ 4: Capstone Integration (ด่าน 10)
  {
    category: "🏆 ภารกิจบูรณาการ",
    title: "นับจำนวนผู้ผ่านเกณฑ์จากลิสต์คะแนน",
    prompt: "กำหนด scores = [45, 80, 65, 30, 90, 72] และ passed_count = 0\nใช้ for s in scores: ตรวจสอบ if s >= 60: ให้เพิ่ม passed_count += 1 เมื่อจบลูปพิมพ์ print(\"จำนวนผู้ผ่านเกณฑ์ =\", passed_count) (ต้องได้: จำนวนผู้ผ่านเกณฑ์ = 4)",
    expected: "จำนวนผู้ผ่านเกณฑ์ = 4",
    starterCode: "scores = [45, 80, 65, 30, 90, 72]\npassed_count = 0\n\n# ใช้ for ท่อง scores ถ้าคะแนน >= 60 ให้เพิ่ม passed_count += 1\n# เมื่อจบลูป พิมพ์ 'จำนวนผู้ผ่านเกณฑ์ = 4'\n"
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
