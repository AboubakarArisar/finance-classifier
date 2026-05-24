import fs from "fs";
import * as XLSX from "xlsx";

const rows = [
  { keyword: "דלק", category: "רכב ונסיעות", status: "מוכר" },
  { keyword: "fuel", category: "רכב ונסיעות", status: "מוכר" },
  { keyword: "office", category: "ציוד משרדי", status: "מוכר" },
  { keyword: "software", category: "מערכות ותוכנה", status: "מוכר" },
  { keyword: 'מע"מ', category: "מסים ותשלומים", status: "מאומת" },
  { keyword: "vat", category: "מסים ותשלומים", status: "מאומת" },
  { keyword: "משכורת", category: "הכנסות", status: "מאומת" },
  { keyword: "salary", category: "הכנסות", status: "מאומת" },
];

fs.mkdirSync("data", { recursive: true });

const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.json_to_sheet(rows);
worksheet["!rtl"] = true;
worksheet["!cols"] = [{ wch: 20 }, { wch: 24 }, { wch: 14 }];
workbook.Workbook = { Views: [{ RTL: true }] };
XLSX.utils.book_append_sheet(workbook, worksheet, "mapping");
XLSX.writeFile(workbook, "data/category-mapping.xlsx");
