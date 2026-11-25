PYTHON CODE ACADEMY – FINAL GATE + ADMIN REPAIR

สิ่งที่เพิ่ม/แก้
1) เพิ่มด่านใหม่ชื่อ “แบบทดสอบท้ายหลักสูตร” แยกจาก Pre-test และ Post-test เดิม จำนวน 10 ข้อ เกณฑ์ผ่าน 70 คะแนน
2) นักเรียนปลดล็อกได้เมื่อ Progress มีสถานะ COMPLETED ทั้ง If-Else, While Loop, For Loop และ ArenaStatus (คอลัมน์เดิมที่ 10) เป็น COMPLETED (ครบ 10 ด่าน Cyber Code Arena)
3) ผลคะแนนบันทึกใน QuizResults ชื่อ Final Course Quiz และสถานะผ่านบันทึกในคอลัมน์ใหม่ FinalQuizStatus ที่ต่อท้าย Progress (คอลัมน์ที่ 13) ไม่เขียนทับข้อมูลเดิม
4) แก้หน้าแอดมิน: ไม่ซ่อน dashboard wrapper ระหว่าง session admin เพราะ lessonArea/gameArea อยู่ภายใน wrapper นี้ จึงเปิดบทเรียนจากปุ่มแอดมินแล้วไม่เป็นหน้าว่าง
5) แก้ข้อมูลภาพรวมแอดมินให้ส่ง progress เป็น object ที่ตรงกับ renderer ของหน้าเว็บ

ติดตั้ง
- สำรอง Apps Script และ Google Sheets ก่อนทุกครั้ง
- ใน Apps Script ให้แทน Code.gs ด้วยไฟล์นี้
- HTML file ต้องชื่อ index (ไม่ใช่ index(1)) แล้วแทนเนื้อหาด้วย index.html
- Save แล้ว Deploy > Manage deployments > Edit > New version > Deploy
- ทดสอบด้วยบัญชีแอดมินเดิม: เปิดบทเรียนทั้ง 3, Arena, Post-test และด่านใหม่
- ทดสอบบัญชีนักเรียน: ด่านใหม่ต้องล็อกจนกว่าจะผ่านครบ 3 บทและ Arena ทั้ง 10 ด่าน

หมายเหตุความปลอดภัย
- ไม่มีการลบชีต ล้างข้อมูล รีเซ็ตรหัสผ่าน หรือสร้างบัญชีใหม่
- ensureAcademySheets_ จะเพิ่มเฉพาะหัวคอลัมน์ FinalQuizStatus ต่อท้าย Progress หากยังไม่มี
- ยังต้องทดสอบบน deployment จริง เพราะการตรวจในเครื่องไม่สามารถยืนยันสิทธิ์ Apps Script, permissions และข้อมูลจริงในบัญชีได้

เริ่มใช้งานจริงและ Deploy เว็บ
1) สร้างหรือเปิด Apps Script project ที่ผูกกับ Google Spreadsheet ที่ต้องการใช้ ตรวจค่า SPREADSHEET_ID ใน Code.gs ให้ตรงกับไฟล์นั้น และให้บัญชีเจ้าของ Script มีสิทธิ์แก้ไข Spreadsheet
2) คัดลอก Code.gs ลงไฟล์สคริปต์ และสร้างไฟล์ HTML ชื่อ index จาก index.html จากนั้นบันทึก
3) ใน Project Settings > Script properties ตั้งค่า PCA_INITIAL_ADMIN_ID, PCA_INITIAL_ADMIN_PASSWORD (อย่างน้อย 10 ตัวอักษร) และ PCA_INITIAL_ADMIN_NAME ตามต้องการ
4) จาก Apps Script editor เลือก setupInitialAdmin() ใน function dropdown แล้วกด Run และอนุญาตสิทธิ์ ฟังก์ชันจะสร้างชีตที่จำเป็นและสร้างบัญชี Admin จากนั้นลบ Script properties ที่ใช้ bootstrap ออก (ฟังก์ชันชื่อที่ลงท้าย _ เป็น private จึงไม่แสดงใน dropdown)
5) การเข้าใช้งาน:
   - สำหรับนักเรียน: กรอกรหัสนักเรียนและชื่อ-นามสกุล หากยังไม่เคยทำ Pre-test ระบบจะพาไปทำ Pre-test ทันที แต่ถ้าเคยเข้าหรือทำแล้ว ระบบจะพาเข้าสู่หน้า Dashboard เลย
   - สำหรับผู้ดูแลระบบ: เพียงพิมพ์คำว่า 'admin' ในช่องรหัสนักเรียน แล้วกดเข้าสู่ห้องเรียน ระบบจะสลับเข้าสู่หน้า Admin ทันที (ไม่มีปุ่มแอดมินแสดงที่หน้าเว็บเพื่อความเรียบร้อย)
   - ระบบเซสชัน: บันทึกสถานะอัตโนมัติ เมื่อกดรีเฟรช (Refresh) หน้าเว็บ ระบบจะไม่หลุดออกจากระบบและยังคงอยู่ในหน้า Dashboard / Admin ตามเดิม
6) Deploy > New deployment > Web app เลือก Execute as: Me เพื่อให้เว็บใช้สิทธิ์เข้าถึง Spreadsheet ของเจ้าของ Script ส่วน Who has access ให้เลือกตามนโยบายของโรงเรียน (เช่น Anyone หรือ Anyone with Google account) จากนั้น Deploy และใช้ URL ที่ได้
7) เมื่อแก้โค้ดภายหลัง ให้ Deploy > Manage deployments > Edit > New version > Deploy เพื่อเผยแพร่เวอร์ชันใหม่

ข้อควรระวัง
- อย่าใส่รหัสผ่าน Admin ลงใน Code.gs หรือแชร์ Script properties ให้ผู้อื่น
- ห้ามใช้ไฟล์ที่มี admin password ตัวอย่างแบบ hardcode; bootstrap ปัจจุบันรับค่าจาก Script properties และลบค่าเหล่านั้นหลังสร้าง Admin สำเร็จ
- ระบบนักเรียนออกแบบให้เข้าถึงง่าย ไม่ต้องจำรหัสผ่าน เพียงใส่รหัสและชื่อก็เริ่มเรียนรู้และบันทึกความก้าวหน้าได้ทันที
