import express from "express";
import cors from "cors";
import fs from "fs";
import ExcelJS from "exceljs";

const app = express();
app.use(cors());
app.use(express.json());

const FILE_PATH = "./RSVP.xlsx";

// 📝 Ghi RSVP vào file Excel
app.post("/api/rsvp", async (req, res) => {
  try {
    const { name, relation, phone, attendance, message } = req.body;
    const now = new Date().toLocaleString("vi-VN");

    let workbook;
    if (fs.existsSync(FILE_PATH)) {
      workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(FILE_PATH);
    } else {
      workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("RSVP");
      sheet.addRow(["Name", "Relation", "Phone", "Attendance", "Message", "Time"]);
    }

    const sheet = workbook.getWorksheet("RSVP") || workbook.worksheets[0];
    sheet.addRow([name, relation, phone, attendance, message, now]);
    await workbook.xlsx.writeFile(FILE_PATH);

    res.json({ success: true, message: "RSVP saved to Excel!" });
  } catch (error) {
    console.error("RSVP Error:", error);
    res.status(500).json({ error: "Server error" });
  }
});

// 📋 Xem toàn bộ RSVP
app.get("/api/rsvp", async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(FILE_PATH);
    const sheet = workbook.getWorksheet("RSVP");
    const rows = sheet.getSheetValues().slice(2);

    const data = rows.map((r) => ({
      name: r[1],
      relation: r[2],
      phone: r[3],
      attendance: r[4],
      message: r[5],
      time: r[6],
    }));

    res.json(data.reverse());
  } catch {
    res.status(500).json({ error: "Không thể đọc dữ liệu" });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`✅ Server chạy tại cổng ${PORT}`));
