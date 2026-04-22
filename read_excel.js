const xlsx = require('xlsx');
const fs = require('fs');

const filePath = 'C:\\Users\\W-11\\Desktop\\مدرسة التربية بالقرآن الكريم\\Excel\\النتيجة الكلية 2025.xlsx';
try {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  console.log(`\n=== Sheet: ${sheetName} ===`);
  const json = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  console.log('Rows 5 to 20:');
  console.log(JSON.stringify(json.slice(5, 20), null, 2));
} catch (error) {
  console.error('Error reading excel file:', error);
}
