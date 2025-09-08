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
      const wb = XLSX.read(data, { type: "array" });
      setWorkbook(wb);
    };
    reader.readAsArrayBuffer(file);
  };

  const processAndExport = () => {
    if (!workbook) return;

    const wbCopy = XLSX.utils.book_new();

    workbook.SheetNames.forEach((sheetName) => {
      const ws = workbook.Sheets[sheetName];

      // Chuyển sheet thành JSON (mảng object)
      const json = XLSX.utils.sheet_to_json(ws, { header: 1 }); // dạng mảng 2D

      const header = json.slice(0, 6); // giữ nguyên 6 dòng đầu
      const dataRows = json.slice(6); // dữ liệu từ dòng 7 trở đi

      const seen = new Set();
      const filtered = [];

      dataRows.forEach((row) => {
        const g = row[6] || "";
        const h = row[7] || "";
        const i = row[8] || "";

        const key = `${g}__${h}__${i}`;
        if (!seen.has(key)) {
          seen.add(key);

          // Convert cột B (index 1) và I (index 8) sang Date dd/MM/yyyy
          if (row[1]) row[1] = formatDate(row[1]);
          if (row[8]) row[8] = formatDate(row[8]);

          filtered.push(row);
        }
      });

      const finalData = [...header, ...filtered];
      const newWs = XLSX.utils.aoa_to_sheet(finalData);

      XLSX.utils.book_append_sheet(wbCopy, newWs, sheetName);
    });

    const timestamp = getTimestamp();
    XLSX.writeFile(wbCopy, `ketqua_${timestamp}.xlsx`);
  };

  // Format chuỗi/Excel date về dd/MM/yyyy
  const formatDate = (value) => {
    let d;
    if (typeof value === "number") {
      const obj = XLSX.SSF.parse_date_code(value);
      if (obj) d = new Date(obj.y, obj.m - 1, obj.d);
    } else if (typeof value === "string") {
      const parts = value.split(" ")[0].split("/");
      if (parts.length === 3) {
        const [dd, mm, yyyy] = parts.map((p) => parseInt(p, 10));
        d = new Date(yyyy, mm - 1, dd);
      }
    }
    if (!d) return value;
    return `${String(d.getDate()).padStart(2, "0")}/${String(
      d.getMonth() + 1
    ).padStart(2, "0")}/${d.getFullYear()}`;
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
      <h2>Excel Date Converter + Deduplicate</h2>
      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      <button onClick={processAndExport} disabled={!workbook}>
        Process & Export
      </button>
    </div>
  );
}
