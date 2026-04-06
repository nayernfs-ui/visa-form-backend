const nodemailer = require("nodemailer");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require("docx");

const EMAIL_USER = process.env.EMAIL_USER;
const EMAIL_PASS = process.env.EMAIL_PASS;
const EMAIL_TO   = process.env.EMAIL_TO;

const RED  = "C8102E";
const NAVY = "1B2A4A";
const LGRAY = "F2F4F7";
const BORDER_COLOR = "D0D5DD";
const W = 9360;

function border(color = BORDER_COLOR) {
  const b = { style: BorderStyle.SINGLE, size: 1, color };
  return { top: b, bottom: b, left: b, right: b };
}

function sectionTitle(text) {
  return new Paragraph({
    spacing: { before: 320, after: 120 },
    children: [new TextRun({ text, bold: true, size: 26, color: "FFFFFF", font: "Arial" })],
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
        children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: String(val), size: 20, font: "Arial", color: "1A1A2E" })] })],
      }),
      new TableCell({
        width: { size: 3120, type: WidthType.DXA },
        borders: border(),
        shading: { fill: LGRAY, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: label, bold: true, size: 20, font: "Arial", color: NAVY })] })],
      }),
    ],
  });
}

function spacer() { return new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }); }
function subHeading(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    alignment: AlignmentType.RIGHT,
    children: [new TextRun({ text, bold: true, size: 22, color: RED, font: "Arial" })],
  });
}
function dynamicTable(rows) {
  if (!rows || rows.length === 0) return spacer();
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [6240, 3120], rows });
}

async function generateWordDoc(data) {
  const yesNo = (v) => v === true || v === "true" || v === "yes" || v === "نعم" ? "نعم" : "لا";

  const children = [
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
    sectionTitle("١. المعلومات الشخصية"), spacer(),
    dynamicTable([
      fieldRow("الاسم بالكامل", data.fullName),
      fieldRow("رقم الهاتف", data.phone),
      fieldRow("البريد الإلكتروني", data.email),
      fieldRow("عنوان الإقامة الحالي", data.address),
      fieldRow("هوية وطنية سارية؟", yesNo(data.hasNationalId)),
      fieldRow("مواطن لأكثر من بلد؟", yesNo(data.multiNational)),
    ]),
    spacer(),
    dynamicTable([
      fieldRow("مقيم دائم في الولايات المتحدة؟", yesNo(data.usResident)),
      fieldRow("تأشيرة غير مهاجر للولايات المتحدة؟", yesNo(data.usVisa)),
      fieldRow("تأشيرة كندية سابقة (10 سنوات)؟", yesNo(data.prevCanadianVisa)),
      fieldRow("بيانات حيوية لكندا سابقاً؟", yesNo(data.biometrics)),
      fieldRow("العام الذي قدمت فيه البيانات الحيوية", data.biometricsYear),
    ]),
    spacer(),
    sectionTitle("٢. التعليم"), spacer(),
    ...(data.education && data.education.length > 0
      ? data.education.flatMap((e, i) => [
          subHeading(`المؤهل ${i + 1}`),
          dynamicTable([
            fieldRow("المؤهل التعليمي", e.degree),
            fieldRow("المؤسسة التعليمية", e.institution),
            fieldRow("تاريخ التخرج", e.graduationDate),
            fieldRow("الشعبة / التخصص", e.major),
          ]),
        ])
      : [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "لا يوجد", size: 20, font: "Arial", color: "999999" })] })]),
    spacer(),
    sectionTitle("٣. التاريخ الوظيفي"), spacer(),
    ...(data.employment && data.employment.length > 0
      ? data.employment.flatMap((j, i) => [
          subHeading(`الوظيفة ${i + 1}`),
          dynamicTable([
            fieldRow("المسمى الوظيفي", j.title),
            fieldRow("اسم الشركة", j.company),
            fieldRow("تاريخ البداية", j.startDate),
          ]),
        ])
      : [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "لا يوجد", size: 20, font: "Arial", color: "999999" })] })]),
    spacer(),
    sectionTitle("٤. الحالة الاجتماعية"), spacer(),
    dynamicTable([fieldRow("الحالة الاجتماعية", data.maritalStatus)]),
    spacer(),
    sectionTitle("٥. سجل التأشيرات"), spacer(),
    dynamicTable([
      fieldRow("هل تم الرفض في أي سفارة سابقاً؟", data.previousRejection),
      fieldRow("هل تقدمت على تأشيرة كندا سابقاً؟", yesNo(data.appliedBefore)),
    ]),
    spacer(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 320, after: 0 },
      border: { top: { style: BorderStyle.SINGLE, size: 2, color: RED } },
      children: [new TextRun({ text: "تم إنشاء هذا المستند تلقائياً عبر نموذج التقديم الإلكتروني", size: 18, color: "999999", font: "Arial" })],
    }),
  ];

  const doc = new Document({
    sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } }, children }],
  });
  return await Packer.toBuffer(doc);
}

module.exports = async (req, res) => {
  // Handle CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const data = req.body;
    const buffer = await generateWordDoc(data);
    const filename = `visa_${(data.fullName || "applicant").replace(/\s+/g, "_")}_${Date.now()}.docx`;

    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: { user: EMAIL_USER, pass: EMAIL_PASS },
    });

    await transporter.sendMail({
      from: `"نموذج التأشيرة الكندية" <${EMAIL_USER}>`,
      to: EMAIL_TO,
      subject: `📋 طلب تأشيرة جديد – ${data.fullName || "مجهول"}`,
      html: `<div dir="rtl" style="font-family:Arial;padding:24px;background:#f8f9fa;border-radius:8px;">
        <h2 style="color:#C8102E;">🍁 طلب تأشيرة كندية جديد</h2>
        <p><strong>الاسم:</strong> ${data.fullName || "—"}</p>
        <p><strong>الهاتف:</strong> ${data.phone || "—"}</p>
        <p><strong>البريد الإلكتروني:</strong> ${data.email || "—"}</p>
        <p><strong>تاريخ التقديم:</strong> ${new Date().toLocaleDateString("ar-EG")}</p>
        <hr/>
        <p style="color:#555;">تجد في المرفق ملف Word يحتوي على جميع بيانات الطلب.</p>
      </div>`,
      attachments: [{ filename, content: buffer }],
    });

    res.json({ success: true, message: "تم إرسال الطلب بنجاح!" });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, message: "حدث خطأ: " + err.message });
  }
};
