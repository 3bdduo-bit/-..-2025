const xlsx = require('xlsx');
const fs = require('fs');

const filePath = 'C:\\Users\\W-11\\Desktop\\مدرسة التربية بالقرآن الكريم\\Excel\\النتيجة الكلية 2025.xlsx';
try {
  const workbook = xlsx.readFile(filePath);
  const allStudents = [];
  
  // The 'الدرجات' sheet is the main summary organized by teacher sections
  const mainSheet = workbook.Sheets['الدرجات'];
  if (mainSheet) {
    const json = xlsx.utils.sheet_to_json(mainSheet, { header: 1 });
    let currentTeacher = 'غير معروف';

    for (let i = 0; i < json.length; i++) {
        const row = json[i];
        if (!row || row.length === 0) continue;

        const rowStr = row.join(' ');
        if (rowStr.includes('مجموعة الأستاذ')) {
            const match = rowStr.match(/مجموعة الأستاذ\/ة\s*:-\s*(.+)/);
            if (match) {
                currentTeacher = match[1].trim();
            } else {
                const parts = rowStr.split(':-');
                currentTeacher = parts[1] ? parts[1].trim() : rowStr;
            }
            continue;
        }

        const id = row[0];
        const name = row[1];
        if (typeof id === 'number' && typeof name === 'string' && name.trim().length > 0 && name !== 'اسم الطالب') {
            let total = row[6] !== undefined ? row[6] : 0;
            let percentage = row[7] !== undefined ? row[7] : 0;
            
            allStudents.push({
                sheet: currentTeacher,
                id: id,
                name: name.trim(),
                test1: row[2] || 0,
                test2: row[3] || 0,
                test3: row[4] || 0,
                test4: row[5] || 0,
                total: total,
                percentage: percentage
            });
        }
    }
  }

  // If we found no students in 'الدرجات', fallback to processing all sheets (old logic)
  if (allStudents.length === 0) {
    for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const json = xlsx.utils.sheet_to_json(sheet, { header: 1 });
        
        let dataStartIdx = -1;
        for (let i = 0; i < json.length; i++) {
          const row = json[i];
          if (row && row.includes('اسم الطالب')) {
             for(let j = i+1; j < json.length; j++) {
                if (json[j] && typeof json[j][1] === 'string' && json[j][1].trim().length > 0 && json[j][1] !== 'اسم الطالب') {
                    dataStartIdx = j;
                    break;
                }
             }
             break;
          }
        }
        
        const extractRow = (row) => {
            if (!row || !row[1] || typeof row[1] !== 'string') return;
            if (row[1].includes('مدير المدرسة') || row[1].includes('القيمة العظمى') || row[1].includes('الدرجة العظمي')) return;
            
            let total = row[6] !== undefined ? row[6] : 0;
            let percentage = row[7] !== undefined ? row[7] : 0;
            if (total > 0 && (!percentage || percentage === 0)) percentage = total / 250;

            allStudents.push({
                sheet: sheetName,
                id: row[0] || '-',
                name: row[1].trim(),
                total: total,
                percentage: percentage
            });
        }

        if (dataStartIdx !== -1) {
           for (let i = dataStartIdx; i < json.length; i++) extractRow(json[i]);
        }
    }
  }
  
  const htmlContent = `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>لوحة بيانات مدرسة التربية بالقرآن الكريم</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Cairo', sans-serif; background-color: #f8fafc; }
        .glass { background: #ffffff; border-radius: 1rem; border: 1px solid #e2e8f0; }
        .card-hover:hover { transform: translateY(-3px); transition: all 0.2s ease; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1); }
    </style>
</head>
<body class="text-slate-800 antialiased p-4 md:p-8">
    <div class="max-w-7xl mx-auto space-y-8">
        
        <header class="glass p-6 shadow-sm flex flex-col lg:flex-row justify-between items-center gap-6">
            <div>
                <h1 class="text-3xl font-bold text-slate-900 mb-2">النتيجة الكلية 2025</h1>
                <p class="text-slate-500">لوحة بيانات مدرسة التربية بالقرآن الكريم - تحليل جميع الفصول والمعلمين</p>
            </div>
            <div class="flex flex-wrap gap-4 text-center justify-center">
                <div class="bg-slate-50 px-6 py-4 rounded-xl border border-slate-200 min-w-[140px]">
                    <p class="text-sm text-slate-500 font-semibold mb-1">إجمالي الطلاب</p>
                    <p class="text-3xl font-bold text-slate-800" id="totalStudents">-</p>
                </div>
                <div class="bg-blue-50 px-6 py-4 rounded-xl border border-blue-200 min-w-[140px]">
                    <p class="text-sm text-blue-600 font-semibold mb-1">متوسط النسبة العام</p>
                    <p class="text-3xl font-bold text-blue-700" id="avgScore">-</p>
                </div>
                <div class="bg-emerald-50 px-6 py-4 rounded-xl border border-emerald-200 min-w-[140px]">
                    <p class="text-sm text-emerald-600 font-semibold mb-1">نسبة النجاح الكلية</p>
                    <p class="text-3xl font-bold text-emerald-700" id="passRate">-</p>
                </div>
            </div>
        </header>

        <div id="teachersDetailedContainer" class="space-y-8 mt-8">
            <!-- Teacher sections will be injected here -->
        </div>

        <div id="allStudentsRankingContainer" class="glass p-6 shadow-sm mt-12 mb-12">
            <h2 class="text-2xl font-bold text-slate-800 mb-6 border-b pb-4 text-center">ترتيب جميع الطلاب حسب النسبة</h2>
            <div class="overflow-x-auto overflow-y-auto rounded-xl border border-slate-200 bg-white" style="max-height: 600px;">

                <table class="w-full text-right">
                    <thead class="bg-slate-100 text-slate-700 border-b border-slate-200 sticky top-0 z-10 shadow-sm">
                        <tr>
                            <th class="p-4 font-bold">الترتيب</th>
                            <th class="p-4 font-bold">اسم الطالب</th>
                            <th class="p-4 font-bold">المعلم / الفصل</th>
                            <th class="p-4 font-bold">المجموع</th>
                            <th class="p-4 font-bold">النسبة</th>
                            <th class="p-4 font-bold">التقدير</th>
                        </tr>
                    </thead>
                    <tbody id="rankingTableBody" class="divide-y divide-slate-100">
                        <!-- Ranking will be injected here -->
                    </tbody>
                </table>
            </div>
        </div>


    </div>

    <script>
        // Disable right click
        document.addEventListener('contextmenu', event => event.preventDefault());

        const studentData = ${JSON.stringify(allStudents)};
        
        // Use all data without deduplication to ensure all classes show their students
        const validStudents = studentData.filter(s => s.total > 0 && s.name);
        
        // Top stats
        document.getElementById('totalStudents').textContent = validStudents.length;
        
        let totalPercentageSum = 0;
        let passCount = 0;
        validStudents.forEach(s => {
            const p = s.percentage * 100;
            totalPercentageSum += p;
            if(p >= 50) passCount++;
        });

        document.getElementById('avgScore').textContent = validStudents.length ? (totalPercentageSum / validStudents.length).toFixed(1) + '%' : '0%';
        document.getElementById('passRate').textContent = validStudents.length ? ((passCount / validStudents.length) * 100).toFixed(1) + '%' : '0%';

        const getGradeText = (p) => {
            if(p >= 90) return '<span class="text-emerald-600 font-bold">امتياز</span>';
            if(p >= 80) return '<span class="text-blue-600 font-bold">جيد جدا</span>';
            if(p >= 70) return '<span class="text-yellow-600 font-bold">جيد</span>';
            if(p >= 50) return '<span class="text-orange-500 font-bold">مقبول</span>';
            return '<span class="text-red-600 font-bold">راسب</span>';
        };

        // Detailed Teacher Analysis
        const teachersDetailedContainer = document.getElementById('teachersDetailedContainer');
        const teacherData = {};

        validStudents.forEach(s => {
            if(!teacherData[s.sheet]) {
                teacherData[s.sheet] = {
                    name: s.sheet,
                    students: [],
                    grades: {
                        'امتياز (90-100%)': 0,
                        'جيد جدا (80-89%)': 0,
                        'جيد (70-79%)': 0,
                        'مقبول (50-69%)': 0,
                        'راسب (أقل من 50%)': 0
                    },
                    total: 0
                };
            }
            teacherData[s.sheet].students.push(s);
            
            const p = s.percentage * 100;
            if(p >= 90) teacherData[s.sheet].grades['امتياز (90-100%)']++;
            else if(p >= 80) teacherData[s.sheet].grades['جيد جدا (80-89%)']++;
            else if(p >= 70) teacherData[s.sheet].grades['جيد (70-79%)']++;
            else if(p >= 50) teacherData[s.sheet].grades['مقبول (50-69%)']++;
            else teacherData[s.sheet].grades['راسب (أقل من 50%)']++;
            
            teacherData[s.sheet].total++;
        });

        Object.keys(teacherData).forEach((teacher, idx) => {
            const data = teacherData[teacher];
            if (data.total === 0) return;

            const section = document.createElement('div');
            section.className = "glass p-6 shadow-sm mt-8";

            const header = document.createElement('h2');
            header.className = "text-2xl font-bold text-slate-800 mb-6 border-b pb-4 text-center";
            header.textContent = 'المعلم / الفصل: ' + teacher;
            section.appendChild(header);

            const grid = document.createElement('div');
            grid.className = "grid grid-cols-1 lg:grid-cols-3 gap-8";

            // Pie chart
            const chartDiv = document.createElement('div');
            chartDiv.className = "lg:col-span-1 bg-slate-50 border border-slate-200 rounded-xl p-4 flex flex-col items-center shadow-sm";
            const chartTitle = document.createElement('h3');
            chartTitle.className = "text-lg font-bold text-slate-800 mb-3 text-center";
            chartTitle.textContent = "تحليل التقديرات";
            chartDiv.appendChild(chartTitle);

            const canvasContainer = document.createElement('div');
            canvasContainer.className = "h-[250px] w-full relative";
            const canvas = document.createElement('canvas');
            canvas.id = 'detailedChart_' + idx;
            canvasContainer.appendChild(canvas);
            chartDiv.appendChild(canvasContainer);
            grid.appendChild(chartDiv);

            // Students table
            const tableDiv = document.createElement('div');
            tableDiv.className = "lg:col-span-2 overflow-y-auto overflow-x-auto rounded-xl border border-slate-200 bg-white";
            tableDiv.style.maxHeight = "350px";

            let tableHTML = \`<table class="w-full text-right relative">
                <thead class="bg-slate-100 text-slate-700 border-b border-slate-200 sticky top-0 shadow-sm z-10">
                    <tr>
                        <th class="p-4 font-bold">م</th>
                        <th class="p-4 font-bold">اسم الطالب</th>
                        <th class="p-4 font-bold">المجموع</th>
                        <th class="p-4 font-bold">النسبة</th>
                        <th class="p-4 font-bold">التقدير</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-slate-100">\`;

            data.students.sort((a, b) => b.total - a.total).forEach((s, sIdx) => {
                const p = s.percentage * 100;
                tableHTML += \`<tr class="hover:bg-slate-50 transition-colors">
                    <td class="p-4 text-slate-500 font-medium">\${sIdx + 1}</td>
                    <td class="p-4 font-bold text-slate-800">\${s.name}</td>
                    <td class="p-4 font-semibold text-slate-800">\${s.total}</td>
                    <td class="p-4"><span class="bg-slate-100 text-slate-700 py-1 px-3 rounded-md text-sm font-bold border border-slate-200">\${p.toFixed(1)}%</span></td>
                    <td class="p-4">\${getGradeText(p)}</td>
                </tr>\`;
            });

            tableHTML += \`</tbody></table>\`;
            tableDiv.innerHTML = tableHTML;
            grid.appendChild(tableDiv);

            section.appendChild(grid);
            teachersDetailedContainer.appendChild(section);

            // Generate pie chart
            new Chart(canvas, {
                type: 'pie',
                data: {
                    labels: Object.keys(data.grades),
                    datasets: [{
                        data: Object.values(data.grades),
                        backgroundColor: ['#10b981', '#3b82f6', '#eab308', '#f97316', '#ef4444'],
                        borderWidth: 1,
                        borderColor: '#ffffff'
                    }]
                },
                options: { 
                    responsive: true, 
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { position: 'bottom', labels: { boxWidth: 12, padding: 10, font: { size: 10 } } },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    let label = context.label || '';
                                    if (label) label += ': ';
                                    let val = context.parsed;
                                    let perc = ((val / data.total) * 100).toFixed(1);
                                    return label + val + ' (' + perc + '%)';
                                }
                            }
                        }
                    }
                }
            });
        });

        // Global Ranking Table
        const rankingTableBody = document.getElementById('rankingTableBody');
        validStudents
            .sort((a, b) => b.percentage - a.percentage)
            .forEach((s, idx) => {
                const p = s.percentage * 100;
                const row = document.createElement('tr');
                row.className = "hover:bg-slate-50 transition-colors";
                row.innerHTML = \`
                    <td class="p-4 text-slate-500 font-medium">\${idx + 1}</td>
                    <td class="p-4 font-bold text-slate-800">\${s.name}</td>
                    <td class="p-4 text-slate-600">\${s.sheet}</td>
                    <td class="p-4 font-semibold text-slate-800">\${s.total}</td>
                    <td class="p-4"><span class="bg-slate-100 text-slate-700 py-1 px-3 rounded-md text-sm font-bold border border-slate-200">\${p.toFixed(1)}%</span></td>
                    <td class="p-4">\${getGradeText(p)}</td>
                \`;
                rankingTableBody.appendChild(row);
            });

    </script>
</body>
</html>`;

  fs.writeFileSync('index.html', htmlContent);
  console.log('Dashboard generated successfully at index.html');
} catch (error) {
  console.error('Error reading excel file:', error);
}
