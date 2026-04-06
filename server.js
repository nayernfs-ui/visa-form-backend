const express = require("express");
const cors = require("cors");
const nodemailer = require("nodemailer");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageOrientation, Header, ImageRun
} = require("docx");
const fs = require("fs");
const path = require("path");

const app = express();
app.use(cors());
app.use(express.json({ limit: "10mb" }));

// ── CONFIG ─────────────────────────────────────────────────────────────────
// Fill these in your .env or directly here before running
const EMAIL_USER = process.env.EMAIL_USER || "your_email@gmail.com";
const EMAIL_PASS = process.env.EMAIL_PASS || "your_app_password";   // Gmail App Password
const EMAIL_TO   = process.env.EMAIL_TO   || "your_email@gmail.com"; // Where to receive forms
// ───────────────────────────────────────────────────────────────────────────

const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: { user: EMAIL_USER, pass: EMAIL_PASS },
});

// ── HELPERS ────────────────────────────────────────────────────────────────
const RED    = "C8102E";
const NAVY   = "1B2A4A";
const LGRAY  = "F2F4F7";
const BORDER_COLOR = "D0D5DD";
const W = 9360; // content width in DXA (US Letter 1" margins)

function border(color = BORDER_COLOR) {
  const b = { style: BorderStyle.SINGLE, size: 1, color };
  return { top: b, bottom: b, left: b, right: b };
}

function sectionTitle(text) {
  return new Paragraph({
    spacing: { before: 320, after: 120 },
    children: [
      new TextRun({
        text,
        bold: true,
        size: 26,
        color: "FFFFFF",
        font: "Arial",
      }),
    ],
    shading: { fill: NAVY, type: ShadingType.CLEAR },
    alignment: AlignmentType.RIGHT,
    indent: { right: 160, left: 160 },
    border: { top: { style: BorderStyle.SINGLE, size: 3, color: RED } },
  });
}

function fieldRow(label, value) {
  const val = value || "—";
  return new TableRow({
    children: [
      new TableCell({
        width: { size: 6240, type: WidthType.DXA },
        borders: border(),
        shading: { fill: "FFFFFF", type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: String(val), size: 20, font: "Arial", color: "1A1A2E" })],
        })],
      }),
      new TableCell({
        width: { size: 3120, type: WidthType.DXA },
        borders: border(),
        shading: { fill: LGRAY, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: label, bold: true, size: 20, font: "Arial", color: NAVY })],
        })],
      }),
    ],
  });
}

function spacer() {
  return new Paragraph({ spacing: { before: 80, after: 80 }, children: [] });
}

function subHeading(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    alignment: AlignmentType.RIGHT,
    children: [new TextRun({ text, bold: true, size: 22, color: RED, font: "Arial" })],
  });
}

function dynamicTable(rows) {
  if (!rows || rows.length === 0) return spacer();
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [6240, 3120],
    rows,
  });
}

// ── WORD GENERATOR ─────────────────────────────────────────────────────────
async function generateWordDoc(data) {

  const yesNo = (v) => v === true || v === "true" || v === "yes" || v === "نعم" ? "نعم" : "لا";

  const sections_children = [
    // ── TITLE ──
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      children: [
        new TextRun({ text: "🍁  ", size: 36 }),
        new TextRun({ text: "استمارة طلب تأشيرة السياحة الكندية", bold: true, size: 36, color: RED, font: "Arial" }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 240 },
      children: [new TextRun({ text: `تاريخ التقديم: ${new Date().toLocaleDateString("ar-EG")}`, size: 20, color: "666666", font: "Arial" })],
    }),

    // ── 1. PERSONAL INFO ──
    sectionTitle("١. المعلومات الشخصية"),
    spacer(),
    dynamicTable([
      fieldRow("الاسم بالكامل", data.fullName),
      fieldRow("رقم الهاتف", data.phone),
      fieldRow("البريد الإلكتروني", data.email),
      fieldRow("عنوان الإقامة الحالي", data.address),
      fieldRow("هوية وطنية سارية؟", yesNo(data.hasNationalId)),
      fieldRow("مواطن لأكثر من بلد؟", yesNo(data.multiNational)),
    ]),

    // multi-national details
    ...(data.multiNational === "نعم" || data.multiNational === true ? [
      spacer(),
      dynamicTable([
        fieldRow("البلد الإضافي", data.additionalCountry),
        fieldRow("مواطن منذ الولادة؟", yesNo(data.citizenSinceBirth)),
        fieldRow("تاريخ الجنسية الإضافية", data.citizenSinceDate),
      ]),
    ] : []),

    spacer(),
    dynamicTable([
      fieldRow("مقيم دائم في الولايات المتحدة؟", yesNo(data.usResident)),
      fieldRow("تأشيرة غير مهاجر للولايات المتحدة؟", yesNo(data.usVisa)),
      fieldRow("تأشيرة كندية سابقة (10 سنوات)؟", yesNo(data.prevCanadianVisa)),
      fieldRow("بيانات حيوية لكندا سابقاً؟", yesNo(data.biometrics)),
      fieldRow("العام الذي قدمت فيه البيانات الحيوية", data.biometricsYear),
      fieldRow("جواز سفر مختلف عن تأشيرة الولايات المتحدة؟", yesNo(data.differentPassport)),
    ]),

    spacer(),

    // ── 2. PREVIOUS RESIDENCY ──
    sectionTitle("٢. الإقامة في دول أخرى (آخر 5 سنوات)"),
    spacer(),
    ...(data.otherResidency && data.otherResidency.length > 0
      ? data.otherResidency.flatMap((r, i) => [
          subHeading(`الإقامة ${i + 1}`),
          dynamicTable([
            fieldRow("اسم البلد", r.country),
            fieldRow("الحالة", r.status),
            fieldRow("من", r.from),
          ]),
        ])
      : [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "لا يوجد", size: 20, font: "Arial", color: "999999" })] })]),

    spacer(),

    // ── 3. EDUCATION ──
    sectionTitle("٣. التعليم"),
    spacer(),
    ...(data.education && data.education.length > 0
      ? data.education.flatMap((e, i) => [
          subHeading(`المؤهل ${i + 1}`),
          dynamicTable([
            fieldRow("المؤهل التعليمي", e.degree),
            fieldRow("المؤسسة التعليمية", e.institution),
            fieldRow("تاريخ التخرج", e.graduationDate),
            fieldRow("الشعبة / التخصص", e.major),
            fieldRow("عنوان المؤسسة", e.institutionAddress),
          ]),
        ])
      : [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "لا يوجد", size: 20, font: "Arial", color: "999999" })] })]),

    spacer(),

    // ── 4. MILITARY ──
    sectionTitle("٤. الخدمة العسكرية"),
    spacer(),
    dynamicTable([
      fieldRow("هل أديت الخدمة العسكرية؟", yesNo(data.military)),
      ...(data.military === "نعم" || data.military === true ? [
        fieldRow("المحافظة", data.militaryGov),
        fieldRow("السلاح", data.militaryBranch),
        fieldRow("الرتبة", data.militaryRank),
        fieldRow("من - إلى", data.militaryDates),
      ] : []),
    ]),

    spacer(),

    // ── 5. EMPLOYMENT ──
    sectionTitle("٥. التاريخ الوظيفي (آخر 10 سنوات)"),
    spacer(),
    ...(data.employment && data.employment.length > 0
      ? data.employment.flatMap((j, i) => [
          subHeading(`الوظيفة ${i + 1}`),
          dynamicTable([
            fieldRow("المسمى الوظيفي", j.title),
            fieldRow("اسم الشركة", j.company),
            fieldRow("عنوان الشركة", j.companyAddress),
            fieldRow("تاريخ البداية", j.startDate),
          ]),
        ])
      : [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "لا يوجد", size: 20, font: "Arial", color: "999999" })] })]),

    spacer(),

    // ── 6. TRAVEL ──
    sectionTitle("٦. سجل السفر (آخر 5 سنوات)"),
    spacer(),
    ...(data.travel && data.travel.length > 0
      ? data.travel.flatMap((t, i) => [
          subHeading(`الرحلة ${i + 1}`),
          dynamicTable([
            fieldRow("البلد", t.country),
            fieldRow("المدينة", t.city),
            fieldRow("غرض الزيارة", t.purpose),
            fieldRow("دخول", t.entry),
            fieldRow("خروج", t.exit),
          ]),
        ])
      : [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "لا يوجد", size: 20, font: "Arial", color: "999999" })] })]),

    spacer(),

    // ── 7. PARENTS ──
    sectionTitle("٧. معلومات الوالدين"),
    spacer(),
    subHeading("معلومات الأب"),
    dynamicTable([
      fieldRow("الاسم", data.fatherName),
      fieldRow("تاريخ الميلاد", data.fatherDOB),
      fieldRow("بلد الميلاد", data.fatherCountry),
      fieldRow("تاريخ الوفاة", data.fatherDeathDate),
      fieldRow("العمل الحالي", data.fatherJob),
      fieldRow("العنوان الحالي", data.fatherAddress),
    ]),
    spacer(),
    subHeading("معلومات الأم"),
    dynamicTable([
      fieldRow("الاسم", data.motherName),
      fieldRow("تاريخ الميلاد", data.motherDOB),
      fieldRow("بلد الميلاد", data.motherCountry),
      fieldRow("تاريخ الوفاة", data.motherDeathDate),
      fieldRow("العمل الحالي", data.motherJob),
      fieldRow("العنوان الحالي", data.motherAddress),
    ]),

    spacer(),

    // ── 8. MARITAL STATUS ──
    sectionTitle("٨. الحالة الاجتماعية"),
    spacer(),
    dynamicTable([
      fieldRow("الحالة الاجتماعية", data.maritalStatus),
    ]),

    // Spouse
    ...(data.maritalStatus === "متزوج" ? [
      spacer(),
      subHeading("معلومات الزوج / الزوجة"),
      dynamicTable([
        fieldRow("الاسم الكامل", data.spouseName),
        fieldRow("تاريخ الميلاد", data.spouseDOB),
        fieldRow("تاريخ الزواج", data.marriageDate),
        fieldRow("محافظة الميلاد", data.spouseBirthGov),
        fieldRow("العمل الحالي", data.spouseJob),
      ]),
    ] : []),

    // Divorced
    ...(data.maritalStatus === "مطلق" ? [
      spacer(),
      subHeading("معلومات الزوج / الزوجة السابق"),
      dynamicTable([
        fieldRow("الاسم الكامل", data.exSpouseName),
        fieldRow("تاريخ الميلاد", data.exSpouseDOB),
        fieldRow("تاريخ الزواج", data.exMarriageDate),
        fieldRow("تاريخ الطلاق", data.divorceDate),
        fieldRow("محافظة الميلاد", data.exSpouseBirthGov),
        fieldRow("العمل الحالي", data.exSpouseJob),
      ]),
    ] : []),

    // Widowed
    ...(data.maritalStatus === "أرمل" || data.maritalStatus === "ارمل" ? [
      spacer(),
      subHeading("معلومات الزوج / الزوجة المتوفى"),
      dynamicTable([
        fieldRow("الاسم الكامل", data.deceasedSpouseName),
        fieldRow("تاريخ الميلاد", data.deceasedSpouseDOB),
        fieldRow("تاريخ الوفاة", data.deceasedSpouseDeathDate),
        fieldRow("محافظة الميلاد", data.deceasedSpouseBirthGov),
      ]),
    ] : []),

    spacer(),

    // ── 9. CHILDREN ──
    sectionTitle("٩. معلومات الأبناء"),
    spacer(),
    ...(data.children && data.children.length > 0
      ? data.children.flatMap((c, i) => [
          subHeading(`الابن / الابنة ${i + 1}`),
          dynamicTable([
            fieldRow("الاسم", c.name),
            fieldRow("تاريخ الميلاد", c.dob),
            fieldRow("محافظة الميلاد", c.birthGov),
          ]),
        ])
      : [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "لا يوجد أبناء", size: 20, font: "Arial", color: "999999" })] })]),

    spacer(),

    // ── 10. VISA HISTORY ──
    sectionTitle("١٠. سجل التأشيرات"),
    spacer(),
    dynamicTable([
      fieldRow("هل تم الرفض في أي سفارة سابقاً؟", data.previousRejection),
      fieldRow("هل تقدمت على تأشيرة كندا سابقاً؟", yesNo(data.appliedBefore)),
      ...(data.appliedBefore === "نعم" || data.appliedBefore === true ? [
        fieldRow("تاريخ التقديم السابق", data.previousApplicationDate),
        fieldRow("نتيجة الطلب السابق", data.previousApplicationResult),
      ] : []),
    ]),

    spacer(),

    // ── FOOTER ──
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 320, after: 0 },
      border: { top: { style: BorderStyle.SINGLE, size: 2, color: RED } },
      children: [
        new TextRun({ text: "تم إنشاء هذا المستند تلقائياً عبر نموذج التقديم الإلكتروني", size: 18, color: "999999", font: "Arial" }),
      ],
    }),
  ];

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
        },
      },
      children: sections_children,
    }],
  });

  return await Packer.toBuffer(doc);
}

// ── ROUTE ──────────────────────────────────────────────────────────────────
app.post("/submit", async (req, res) => {
  try {
    const data = req.body;
    const buffer = await generateWordDoc(data);
    const filename = `visa_${(data.fullName || "applicant").replace(/\s+/g, "_")}_${Date.now()}.docx`;
    const filepath = path.join(__dirname, "tmp_" + filename);
    fs.writeFileSync(filepath, buffer);

    await transporter.sendMail({
      from: `"نموذج التأشيرة الكندية" <${EMAIL_USER}>`,
      to: EMAIL_TO,
      subject: `📋 طلب تأشيرة جديد – ${data.fullName || "مجهول"}`,
      html: `
        <div dir="rtl" style="font-family:Arial;padding:24px;background:#f8f9fa;border-radius:8px;">
          <h2 style="color:#C8102E;">🍁 طلب تأشيرة كندية جديد</h2>
          <p><strong>الاسم:</strong> ${data.fullName || "—"}</p>
          <p><strong>الهاتف:</strong> ${data.phone || "—"}</p>
          <p><strong>البريد الإلكتروني:</strong> ${data.email || "—"}</p>
          <p><strong>تاريخ التقديم:</strong> ${new Date().toLocaleDateString("ar-EG")}</p>
          <hr/>
          <p style="color:#555;">تجد في المرفق ملف Word يحتوي على جميع بيانات الطلب.</p>
        </div>
      `,
      attachments: [{ filename, content: buffer }],
    });

    fs.unlinkSync(filepath);
    res.json({ success: true, message: "تم إرسال الطلب بنجاح!" });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, message: "حدث خطأ أثناء الإرسال: " + err.message });
  }
});

app.get("/health", (_, res) => res.json({ status: "ok" }));

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`✅  Server running on http://localhost:${PORT}`));
