import express from "express";
import cors from "cors";
import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";
import bodyParser from "body-parser";
import { fileURLToPath } from "url";

const app = express();
const PORT = process.env.PORT || 5000;

// Đường dẫn file Excel
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const excelPath = path.join(__dirname, "rsvp.xlsx");

// ✅ Cho phép frontend truy cập
app.use(
  cors({
    origin: [
      "https://ngocthang-huyentrang.vercel.app",
      "http://localhost:5173",
    ],
  })
);

app.use(bodyParser.json());

// 📩 API nhận RSVP
app.post("/api/rsvp", async (req, res) => {
  try {
    const { name, attendance, message } = req.body;
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    if (fs.existsSync(excelPath)) {
      await workbook.xlsx.readFile(excelPath);
      worksheet = workbook.getWorksheet(1);
    } else {
      worksheet = workbook.addWorksheet("RSVP");
      worksheet.addRow(["Tên", "Tham dự", "Lời chúc"]);
    }

    worksheet.addRow([name, attendance, message]);
    await workbook.xlsx.writeFile(excelPath);

    res.json({ success: true, message: "Gửi thành công!" });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: "Lỗi server!" });
  }
});

// 📤 API lấy danh sách lời chúc
app.get("/api/rsvp", async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);
    const worksheet = workbook.getWorksheet(1);

    const data = worksheet
      .getRows(2, worksheet.rowCount - 1)
      .map((row) => ({
        name: row.getCell(1).value,
        attendance: row.getCell(2).value,
        message: row.getCell(3).value,
      }));

    res.json(data);
  } catch (error) {
    res.status(500).json({ message: "Không thể đọc file Excel" });
  }
});

// 📥 API tải file Excel
app.get("/api/download", (req, res) => {
  if (fs.existsSync(excelPath)) {
    res.download(excelPath, "rsvp.xlsx");
  } else {
    res.status(404).json({ message: "Chưa có file RSVP nào." });
  }
});

app.listen(PORT, () =>
  console.log(`✅ Server đang chạy tại cổng ${PORT}`)
);
