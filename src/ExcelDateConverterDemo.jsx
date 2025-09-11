import React, { useState } from "react";
import * as XLSX from "xlsx";

export default function ExcelDateConverterDemo() {
  const [workbook, setWorkbook] = useState(null);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      // Đọc giữ value gốc (cố gắng lấy date nếu có)
      const wb = XLSX.read(data, { type: "array", cellDates: true, cellNF: false });
      setWorkbook(wb);
    };
    reader.readAsArrayBuffer(file);
  };

  const processAndExport = () => {
    if (!workbook) return;
    const wbCopy = XLSX.utils.book_new();

    workbook.SheetNames.forEach((sheetName) => {
      const ws = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

      const header = json.slice(0, 6); // giữ 6 dòng đầu
      const dataRows = json.slice(6);

      const seen = new Set();
      const filtered = [];

      dataRows.forEach((row) => {
        const g = row[6] || "";
        const h = row[7] || "";
        const i = row[8] || "";
        const key = `${g || "null"}__${h || "null"}__${i || "null"}`;
        if (!seen.has(key)) {
          seen.add(key);

          // parse cột B (index 1) và I (index 8) -> Date object nếu có thể
          if (row[1]) row[1] = parseToDate(row[1]);
          if (row[8]) row[8] = parseToDate(row[8]);

          filtered.push(row);
        }
      });

      const finalData = [...header, ...filtered];

      // --- Clean & chuẩn hóa dữ liệu (không dùng aoa_to_sheet) ---
      const cleaned = finalData.map((row) =>
        (row || []).map((cell) => {
          if (cell instanceof Date) return cell;
          if (typeof cell === "number") return Number(cell);
          if (typeof cell === "string") return cell.trim();
          return cell == null ? "" : String(cell);
        })
      );

      // --- Build sheet manually để KHÔNG vướng các custom SSF format ---
      const newWs = {};
      let maxC = 0;
      for (let r = 0; r < cleaned.length; r++) {
        const row = cleaned[r] || [];
        for (let c = 0; c < row.length; c++) {
          const v = row[c];
          if (v === undefined || v === null || v === "") {
            // optional: skip empty cells (keeps sheet small)
            continue;
          }
          if (c > maxC) maxC = c;
          const addr = XLSX.utils.encode_cell({ r, c });
          // tạo cell object rõ ràng, tránh dùng aoa_to_sheet
          if (v instanceof Date) {
            newWs[addr] = { v: v, t: "d", z: "dd/mm/yyyy" }; // Date object, type d
          } else if (typeof v === "number") {
            newWs[addr] = { v: v, t: "n", z: "0" }; // number, base format
          } else {
            // chuỗi — đảm bảo escape
            newWs[addr] = { v: String(v), t: "s" };
          }
        }
      }

      // set reference range (từ A1 tới last cell)
      const rows = cleaned.length;
      const range = { s: { r: 0, c: 0 }, e: { r: Math.max(0, rows - 1), c: Math.max(0, maxC) } };
      newWs["!ref"] = XLSX.utils.encode_range(range);

      // append sheet hoàn toàn mới (không mang theo format cũ)
      XLSX.utils.book_append_sheet(wbCopy, newWs, sheetName);
    });

    try {
      const timestamp = getTimestamp();
      XLSX.writeFile(wbCopy, `ketqua_${timestamp}.xlsx`);
    } catch (err) {
      console.error("Write error:", err);
      alert("Có lỗi khi ghi file: " + (err && err.message ? err.message : err));
    }
  };

  // Parse to Date (cố gắng nhận nhiều dạng)
  const parseToDate = (value) => {
    if (!value && value !== 0) return value;
    if (value instanceof Date) return value;

    // nếu là số có thể là excel serial
    if (typeof value === "number") {
      try {
        const obj = XLSX.SSF.parse_date_code(value);
        if (obj && obj.y) {
          const hr = obj.H || 0;
          const mn = obj.M || 0;
          const sc = Math.floor(obj.S || 0);
          return new Date(obj.y, obj.m - 1, obj.d, hr, mn, sc);
        }
      } catch (e) {
        // fallback bên dưới
      }
      // fallback: treat as excel serial to JS date (accounting 1900 bug)
      const epoch = Date.UTC(1899, 11, 30);
      const offset = value > 60 ? value - 1 : value; // skip fake 1900-02-29
      const ms = Math.round(offset * 24 * 60 * 60 * 1000);
      return new Date(epoch + ms);
    }

    if (typeof value === "string") {
      const s = value.trim();
      // dd/mm/yyyy hoặc d/m/yyyy
      const dm = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+.*)?$/);
      if (dm) {
        const dd = parseInt(dm[1], 10);
        const mm = parseInt(dm[2], 10);
        let yyyy = parseInt(dm[3], 10);
        if (yyyy < 100) yyyy += 2000;
        return new Date(yyyy, mm - 1, dd);
      }
      // yyyy-mm-dd
      const ym = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
      if (ym) {
        return new Date(parseInt(ym[1], 10), parseInt(ym[2], 10) - 1, parseInt(ym[3], 10));
      }
      // cuối cùng thử Date.parse
      const parsed = new Date(s);
      if (!isNaN(parsed)) return parsed;
    }

    return value;
  };

  const getTimestamp = () => {
    const now = new Date();
    const pad = (n) => (n < 10 ? "0" + n : n);
    return (
      now.getFullYear().toString() +
      pad(now.getMonth() + 1) +
      pad(now.getDate()) +
      pad(now.getHours()) +
      pad(now.getMinutes()) +
      pad(now.getSeconds())
    );
  };

  return (
    <div style={{ padding: 20 }}>
      <h2>Tools hỗ trợ danh sách quà tặng</h2>
      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      <button onClick={processAndExport} disabled={!workbook}>
        Process & Export
      </button>
    </div>
  );
}
