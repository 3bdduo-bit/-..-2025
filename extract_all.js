const xlsx = require('xlsx');
const fs = require('fs');

const filePath = 'C:\\Users\\W-11\\Desktop\\مدرسة التربية بالقرآن الكريم\\Excel\\النتيجة الكلية 2025.xlsx';
try {
  const workbook = xlsx.readFile(filePath);
  const allData = {};
  
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    // console.log(`\n=== Sheet: ${sheetName} ===`);
    const json = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    
    const students = [];
    // Data seems to start at row index 5 (6th row) based on previous log for 'الدرجات'
    // but other sheets had headers around row 5-8.
    // Let's find the row that has 'اسم الطالب' (Student Name) and data starts below it.
    let dataStartIdx = -1;
    for (let i = 0; i < json.length; i++) {
      const row = json[i];
      if (row && row.includes('اسم الطالب')) {
         // data might start a few rows after
         // find the first row that has a number in the first column or string in second
         for(let j = i+1; j < json.length; j++) {
            if (json[j] && typeof json[j][1] === 'string' && json[j][1].trim().length > 0 && json[j][1] !== 'اسم الطالب') {
                dataStartIdx = j;
                break;
            }
         }
         break;
      }
    }
    
    if (dataStartIdx !== -1) {
       for (let i = dataStartIdx; i < json.length; i++) {
           const row = json[i];
           if (!row || !row[1]) continue; // Skip empty rows or rows without name
           if (typeof row[1] !== 'string') continue;
           if (row[1].includes('مدير المدرسة') || row[1].includes('القيمة العظمى')) continue;
           
           // In some sheets, the columns might differ. Let's capture the whole row.
           students.push(row);
       }
    } else {
       // fallback: just grab rows that look like student data (index 1 is string, index 2 is number)
       for (let i = 0; i < json.length; i++) {
           const row = json[i];
           if (row && row.length > 2 && typeof row[1] === 'string' && typeof row[2] === 'number') {
               students.push(row);
           }
       }
    }
    
    if (students.length > 0) {
      allData[sheetName] = students;
    }
  }
  
  fs.writeFileSync('C:\\Users\\W-11\\Desktop\\مدرسة التربية بالقرآن الكريم\\Excel\\dashboard_data.json', JSON.stringify(allData, null, 2));
  console.log('Data extracted successfully to dashboard_data.json');
} catch (error) {
  console.error('Error reading excel file:', error);
}
