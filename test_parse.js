const XLSX = require('xlsx');
const fs = require('fs');
// Create a dummy workbook to test sheet_to_json with objects
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet([
  [null, "Schedule", null, null, null, "Program", null, "Website", null, null],
  ["Date", "Start Time", "End Time", "Duration", "Schedule Type", "Rerun", "Program", "EP", "Peak Views", "Views", "Peak Views"],
  ["2026-01-04", "10:30", "11:00", 1800, "schedule", "N", "มหาอำนาจบ้านนา", "กาแฟขี้ชะมด", 579, 486, 575]
]);
// add merges
ws['!merges'] = [
  { s: { r: 0, c: 1 }, e: { r: 0, c: 4 } }, // Schedule (B1:E1)
  { s: { r: 0, c: 5 }, e: { r: 0, c: 6 } }, // Program (F1:G1)
  { s: { r: 0, c: 7 }, e: { r: 0, c: 10 } } // Website (H1:K1)
];
XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
XLSX.writeFile(wb, "test.xlsx");

const readWb = XLSX.readFile("test.xlsx");
const readWs = readWb.Sheets[readWb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(readWs, { header: 1, defval: "", blankrows: true });
console.log(data.slice(0, 3));
console.log("Merges:", readWs['!merges']);
