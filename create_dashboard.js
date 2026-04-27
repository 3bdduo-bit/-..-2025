const xlsx = require('xlsx');
const fs = require('fs');

const filePath = 'C:\\\\Users\\\\W-11\\\\Desktop\\\\مدرسة التربية بالقرآن الكريم\\\\Excel\\\\النتيجة الكلية 2025.xlsx';
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
            const match = rowStr.match(/مجموعة الأستاذ[\/\\]?ة?\s*[:-]\s*(.+)/);
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
  
  const validStudents = allStudents.filter(s => s.total > 0 && s.name);
  let totalPercentageSum = 0;
  let passCount = 0;
  validStudents.forEach(s => {
      const p = s.percentage * 100;
      totalPercentageSum += p;
      if(p >= 50) passCount++;
  });

  const avgScoreStr = validStudents.length ? (totalPercentageSum / validStudents.length).toFixed(1) + '%' : '0%';
  const passRateStr = validStudents.length ? ((passCount / validStudents.length) * 100).toFixed(1) + '%' : '0%';

  const getGradeText = (p) => {
      if(p >= 90) return '<span class="text-emerald-600 font-bold">امتياز</span>';
      if(p >= 80) return '<span class="text-blue-600 font-bold">جيد جدا</span>';
      if(p >= 70) return '<span class="text-yellow-600 font-bold">جيد</span>';
      if(p >= 50) return '<span class="text-orange-500 font-bold">مقبول</span>';
      return '<span class="text-red-600 font-bold">راسب</span>';
  };

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

  let teachersDetailedHTML = '';
  const chartConfigs = [];
  
  Object.keys(teacherData).forEach((teacher, idx) => {
      const data = teacherData[teacher];
      if (data.total === 0) return;

      chartConfigs.push({
          id: 'detailedChart_' + idx,
          labels: Object.keys(data.grades),
          data: Object.values(data.grades),
          total: data.total
      });

      teachersDetailedHTML += `
      <section class="glass bg-white p-4 md:p-6 shadow-sm rounded-2xl teacher-section hidden" data-teacher-name="${teacher}">
          <h2 class="text-xl md:text-2xl font-bold text-slate-800 mb-4 md:mb-6 border-b pb-4 text-center">المعلم / الفصل: ${teacher}</h2>
          <div class="grid grid-cols-1 lg:grid-cols-3 gap-6 md:gap-8">
              <div class="lg:col-span-1 bg-slate-50 border border-slate-200 rounded-xl p-4 flex flex-col items-center shadow-sm">
                  <h3 class="text-lg font-bold text-slate-800 mb-3 text-center">تحليل التقديرات</h3>
                  <div class="h-[250px] md:h-[300px] w-full relative">
                      <canvas id="detailedChart_${idx}" data-chart-idx="${idx}"></canvas>
                  </div>
              </div>
              <div class="lg:col-span-2 overflow-y-auto overflow-x-auto rounded-xl border border-slate-200 bg-white" style="max-height: 400px;">
                  <table class="w-full text-right relative mobile-card-table">
                      <thead class="bg-slate-100 text-slate-700 border-b border-slate-200 sticky top-0 shadow-sm z-10">
                          <tr>
                              <th class="p-3 md:p-4 font-bold">م</th>
                              <th class="p-3 md:p-4 font-bold">اسم الطالب</th>
                              <th class="p-3 md:p-4 font-bold">المجموع</th>
                              <th class="p-3 md:p-4 font-bold">النسبة</th>
                              <th class="p-3 md:p-4 font-bold">التقدير</th>
                          </tr>
                      </thead>
                      <tbody class="divide-y divide-slate-100">`;

      data.students.sort((a, b) => b.total - a.total).forEach((s, sIdx) => {
          const p = s.percentage * 100;
          teachersDetailedHTML += `
          <tr class="hover:bg-slate-50 transition-colors">
              <td class="p-3 md:p-4 text-slate-500 font-medium" data-label="م">${sIdx + 1}</td>
              <td class="p-3 md:p-4 font-bold text-slate-800" data-label="اسم الطالب">${s.name}</td>
              <td class="p-3 md:p-4 font-semibold text-slate-800" data-label="المجموع">${s.total}</td>
              <td class="p-3 md:p-4" data-label="النسبة"><span class="bg-slate-100 text-slate-700 py-1 px-3 rounded-md text-sm font-bold border border-slate-200">${p.toFixed(1)}%</span></td>
              <td class="p-3 md:p-4" data-label="التقدير">${getGradeText(p)}</td>
          </tr>`;
      });
      teachersDetailedHTML += `</tbody></table></div></div></section>`;
  });

  const teacherStats = Object.keys(teacherData).map(teacher => {
      const data = teacherData[teacher];
      const excellentCount = data.grades['امتياز (90-100%)'];
      const excellentPercentage = data.total > 0 ? (excellentCount / data.total) * 100 : 0;
      return {
          name: teacher,
          total: data.total,
          excellentCount: excellentCount,
          excellentPercentage: excellentPercentage
      };
  }).filter(t => t.total > 0);

  teacherStats.sort((a, b) => {
      if (b.excellentPercentage !== a.excellentPercentage) {
          return b.excellentPercentage - a.excellentPercentage;
      }
      return b.excellentCount - a.excellentCount;
  });

  let teacherRankingHTML = '';
  teacherStats.forEach((t, idx) => {
      teacherRankingHTML += `
          <tr class="hover:bg-slate-50 transition-colors">
              <td class="p-3 md:p-4 text-slate-500 font-medium" data-label="الترتيب">${idx + 1}</td>
              <td class="p-3 md:p-4 font-bold text-slate-800" data-label="المعلم / الفصل">${t.name}</td>
              <td class="p-3 md:p-4 text-slate-600" data-label="إجمالي الطلاب">${t.total}</td>
              <td class="p-3 md:p-4 font-semibold text-slate-800" data-label="عدد الامتياز">${t.excellentCount}</td>
              <td class="p-3 md:p-4" data-label="نسبة الامتياز"><span class="bg-slate-100 text-emerald-700 py-1 px-3 rounded-md text-sm font-bold border border-slate-200">${t.excellentPercentage.toFixed(1)}%</span></td>
          </tr>`;
  });

  let rankingHTML = '';
  validStudents.sort((a, b) => b.percentage - a.percentage).forEach((s, idx) => {
      const p = s.percentage * 100;
      rankingHTML += `
          <tr class="hover:bg-slate-50 transition-colors">
              <td class="p-3 md:p-4 text-slate-500 font-medium" data-label="الترتيب">${idx + 1}</td>
              <td class="p-3 md:p-4 font-bold text-slate-800" data-label="اسم الطالب">${s.name}</td>
              <td class="p-3 md:p-4 text-slate-600" data-label="المعلم / الفصل">${s.sheet}</td>
              <td class="p-3 md:p-4 font-semibold text-slate-800" data-label="المجموع">${s.total}</td>
              <td class="p-3 md:p-4" data-label="النسبة"><span class="bg-slate-100 text-slate-700 py-1 px-3 rounded-md text-sm font-bold border border-slate-200">${p.toFixed(1)}%</span></td>
              <td class="p-3 md:p-4" data-label="التقدير">${getGradeText(p)}</td>
          </tr>`;
  });

  const htmlContent = `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="لوحة بيانات مدرسة التربية بالقرآن الكريم تعرض النتيجة الكلية لعام 2025 وتحليل أداء الطلاب والمعلمين.">
    <meta name="keywords" content="لوحة بيانات, نتائج الطلاب, مدرسة التربية بالقرآن الكريم, تقييم المعلمين">
    <title>لوحة بيانات مدرسة التربية بالقرآن الكريم</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js" defer></script>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Cairo', sans-serif; background-color: #f8fafc; }
        .glass { background: #ffffff; border-radius: 1rem; border: 1px solid #e2e8f0; }
        .animate-fade-in { animation: fadeIn 0.5s ease-out forwards; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        
        @media (max-width: 768px) {
            .mobile-card-table thead { display: none; }
            .mobile-card-table tr {
                display: block;
                margin-bottom: 1.5rem;
                border: 1px solid #e2e8f0;
                border-radius: 1rem;
                padding: 1rem;
                background: #fff;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
            }
            .mobile-card-table td {
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 0.75rem 0;
                border-bottom: 1px solid #f1f5f9;
                text-align: left !important;
            }
            .mobile-card-table td:last-child { border-bottom: none; }
            .mobile-card-table td::before {
                content: attr(data-label);
                font-weight: 700;
                color: #64748b;
                margin-left: 1rem;
                text-align: right;
            }
        }
    </style>
</head>
<body class="text-slate-800 antialiased min-h-screen flex flex-col">

    <!-- Role Selection Screen -->
    <div id="roleSelectionScreen" class="flex-1 flex flex-col items-center justify-center p-4 md:p-8 animate-fade-in bg-slate-50 min-h-screen">
        <div class="text-center mb-10">
            <h1 class="text-3xl md:text-5xl font-bold text-slate-900 mb-4 tracking-tight">النتيجة الكلية 2025</h1>
            <p class="text-lg text-slate-500">مدرسة التربية بالقرآن الكريم - يرجى اختيار طريقة الدخول</p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-6 max-w-5xl w-full">
            <button onclick="setRole('manager')" class="glass bg-white p-8 flex flex-col items-center text-center hover:-translate-y-2 hover:shadow-xl transition-all duration-300 group cursor-pointer border-2 border-transparent hover:border-blue-500 rounded-2xl shadow-sm">
                <div class="w-20 h-20 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center mb-6 group-hover:scale-110 transition-transform">
                    <svg class="w-10 h-10" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 6V4m0 2a2 2 0 100 4m0-4a2 2 0 110 4m-6 8a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4m6 6v10m6-2a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4"></path></svg>
                </div>
                <h2 class="text-2xl font-bold text-slate-800 mb-2">دخول كمدير</h2>
                <p class="text-slate-500 text-sm">عرض لوحة التحكم الشاملة والبحث المتقدم في جميع الفصول</p>
            </button>
            
            <button onclick="setRole('teacher')" class="glass bg-white p-8 flex flex-col items-center text-center hover:-translate-y-2 hover:shadow-xl transition-all duration-300 group cursor-pointer border-2 border-transparent hover:border-emerald-500 rounded-2xl shadow-sm">
                <div class="w-20 h-20 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mb-6 group-hover:scale-110 transition-transform">
                    <svg class="w-10 h-10" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 20H5a2 2 0 01-2-2V6a2 2 0 012-2h10a2 2 0 012 2v1m2 13a2 2 0 01-2-2V7m2 13a2 2 0 002-2V9.5a2 2 0 00-2-2h-2m-4-3H9M7 16h6M7 8h6v4H7V8z"></path></svg>
                </div>
                <h2 class="text-2xl font-bold text-slate-800 mb-2">دخول كمعلم</h2>
                <p class="text-slate-500 text-sm">عرض نتائج وتقييمات طلاب فصلك فقط وتحديد نسب النجاح</p>
            </button>

            <button onclick="setRole('student')" class="glass bg-white p-8 flex flex-col items-center text-center hover:-translate-y-2 hover:shadow-xl transition-all duration-300 group cursor-pointer border-2 border-transparent hover:border-purple-500 rounded-2xl shadow-sm">
                <div class="w-20 h-20 bg-purple-100 text-purple-600 rounded-full flex items-center justify-center mb-6 group-hover:scale-110 transition-transform">
                    <svg class="w-10 h-10" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 14l9-5-9-5-9 5 9 5z"></path><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 14l6.16-3.422a12.083 12.083 0 01.665 6.479A11.952 11.952 0 0012 20.055a11.952 11.952 0 00-6.824-2.998 12.078 12.078 0 01.665-6.479L12 14z"></path><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 14l9-5-9-5-9 5 9 5zm0 0l6.16-3.422a12.083 12.083 0 01.665 6.479A11.952 11.952 0 0012 20.055a11.952 11.952 0 00-6.824-2.998 12.078 12.078 0 01.665-6.479L12 14zm-4 6v-7.5l4-2.222"></path></svg>
                </div>
                <h2 class="text-2xl font-bold text-slate-800 mb-2">دخول كطالب</h2>
                <p class="text-slate-500 text-sm">استعلام عن نتيجتك الشخصية ومجموعك الكلي (متاح 3 مرات كحد أقصى)</p>
            </button>
        </div>
    </div>

    <!-- Main App Container -->
    <div id="mainApp" class="hidden p-2 md:p-8 w-full max-w-7xl mx-auto space-y-6 md:space-y-8 animate-fade-in bg-slate-50">
        <!-- Top Nav / Back Button -->
        <div class="flex justify-end">
            <button onclick="logout()" class="text-sm bg-white border border-slate-200 px-4 py-2 rounded-xl shadow-sm hover:bg-slate-50 font-bold text-slate-700 transition-colors flex items-center gap-2">
                <span>عودة للرئيسية</span>
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 8l4 4m0 0l-4 4m4-4H3"></path></svg>
            </button>
        </div>

        <!-- Student Search View -->
        <div id="studentSearchView" class="hidden">
            <div class="glass bg-white p-6 md:p-10 shadow-sm max-w-2xl mx-auto text-center rounded-2xl">
                <div class="w-16 h-16 bg-purple-100 text-purple-600 rounded-full flex items-center justify-center mx-auto mb-4">
                    <svg class="w-8 h-8" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
                </div>
                <h2 class="text-2xl font-bold text-slate-800 mb-2">ابحث عن نتيجتك</h2>
                <p class="text-slate-500 mb-6 text-sm">تنبيه: لأسباب أمنية وللحفاظ على الخصوصية، يمكنك البحث وعرض نتيجة <span class="text-red-500 font-bold">3 طلاب كحد أقصى</span> من هذا الجهاز.</p>
                <div class="flex flex-col md:flex-row gap-3">
                    <input type="text" id="studentNameInput" class="flex-1 px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-purple-500 focus:border-purple-500 outline-none transition-all text-slate-800 font-semibold" placeholder="اكتب اسمك بالكامل هنا...">
                    <button id="studentSearchExecuteBtn" onclick="executeStudentSearch()" class="bg-purple-600 text-white px-8 py-3 rounded-xl font-bold hover:bg-purple-700 transition-colors shadow-md">بحث</button>
                </div>
                <div id="studentSearchError" class="text-red-500 text-sm mt-3 hidden font-semibold">لم يتم العثور على طالب بهذا الاسم. تأكد من كتابة الاسم ثلاثي أو رباعي بشكل صحيح.</div>
                <div id="studentResultCard" class="hidden mt-8 text-right bg-slate-50 border border-slate-200 rounded-xl p-6 shadow-inner animate-fade-in">
                    <!-- Result goes here -->
                </div>
            </div>
        </div>

        <!-- Teacher Selection View -->
        <div id="teacherSelectView" class="hidden">
            <div class="glass bg-white p-6 shadow-sm mb-6 rounded-2xl flex flex-col md:flex-row items-center gap-4 justify-between">
                <div class="text-center md:text-right">
                    <h2 class="text-xl font-bold text-slate-800">مرحباً بك أستاذي الفاضل</h2>
                    <p class="text-sm text-slate-500">اختر اسمك من القائمة لعرض بيانات طلاب فصلك حصرياً</p>
                </div>
                <select id="teacherDropdown" onchange="onTeacherSelected()" class="px-4 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none w-full md:w-72 bg-slate-50 font-bold text-slate-700 cursor-pointer">
                    <option value="">-- اضغط لاختيار اسمك --</option>
                    ${Object.keys(teacherData).map(t => `<option value="${t}">${t}</option>`).join('')}
                </select>
            </div>
        </div>

        <!-- Dashboard Content -->
        <div id="dashboardContent" class="hidden space-y-6 md:space-y-8">
            
            <header id="managerHeader" class="glass bg-white p-4 md:p-6 shadow-sm flex flex-col lg:flex-row justify-between items-center gap-6 rounded-2xl hidden">
                <div class="text-center lg:text-right w-full lg:w-auto">
                    <h1 class="text-2xl md:text-3xl font-bold text-slate-900 mb-2">لوحة تحكم الإدارة</h1>
                    <p class="text-sm md:text-base text-slate-500">نظرة عامة على أداء جميع الفصول والمعلمين</p>
                </div>
                <div class="grid grid-cols-2 md:grid-cols-3 gap-3 md:gap-4 w-full lg:w-auto">
                    <div class="bg-slate-50 p-3 md:px-6 md:py-4 rounded-xl border border-slate-200 text-center">
                        <p class="text-xs md:text-sm text-slate-500 font-semibold mb-1">إجمالي الطلاب</p>
                        <p class="text-xl md:text-3xl font-bold text-slate-800">${validStudents.length}</p>
                    </div>
                    <div class="bg-blue-50 p-3 md:px-6 md:py-4 rounded-xl border border-blue-200 text-center">
                        <p class="text-xs md:text-sm text-blue-600 font-semibold mb-1">متوسط النسبة</p>
                        <p class="text-xl md:text-3xl font-bold text-blue-700">${avgScoreStr}</p>
                    </div>
                    <div class="bg-emerald-50 p-3 md:px-6 md:py-4 rounded-xl border border-emerald-200 text-center col-span-2 md:col-span-1">
                        <p class="text-xs md:text-sm text-emerald-600 font-semibold mb-1">نسبة النجاح</p>
                        <p class="text-xl md:text-3xl font-bold text-emerald-700">${passRateStr}</p>
                    </div>
                </div>
            </header>

            <div id="managerGlobalSearch" class="glass bg-white p-6 shadow-sm rounded-2xl hidden mb-8">
                 <h2 class="text-lg font-bold text-slate-800 mb-4">البحث الشامل للطلاب</h2>
                 <div class="relative max-w-3xl">
                    <input type="text" id="managerSearchInput" onkeyup="filterAllStudents()" class="w-full px-4 py-3 pl-10 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all text-slate-800 font-semibold" placeholder="اكتب اسم الطالب للبحث السريع في جميع الفصول...">
                    <svg class="w-6 h-6 text-slate-400 absolute left-4 top-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
                 </div>
            </div>

            <div id="teachersDetailedContainer" class="space-y-6 md:space-y-8">
                ${teachersDetailedHTML}
            </div>

            <section id="teacherRankingContainer" class="glass bg-white p-4 md:p-6 shadow-sm rounded-2xl mt-8 mb-8 hidden">
                <h2 class="text-xl md:text-2xl font-bold text-slate-800 mb-4 md:mb-6 border-b pb-4 text-center">ترتيب المعلمين حسب نسبة الامتياز</h2>
                <div class="overflow-x-auto overflow-y-auto rounded-xl border border-slate-200 bg-white" style="max-height: 600px;">
                    <table class="w-full text-right mobile-card-table">
                        <thead class="bg-slate-100 text-slate-700 border-b border-slate-200 sticky top-0 z-10 shadow-sm">
                            <tr>
                                <th class="p-3 md:p-4 font-bold">الترتيب</th>
                                <th class="p-3 md:p-4 font-bold">المعلم / الفصل</th>
                                <th class="p-3 md:p-4 font-bold">إجمالي الطلاب</th>
                                <th class="p-3 md:p-4 font-bold">عدد الامتياز</th>
                                <th class="p-3 md:p-4 font-bold">نسبة الامتياز</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-100">${teacherRankingHTML}</tbody>
                    </table>
                </div>
            </section>

            <section id="allStudentsRankingContainer" class="glass bg-white p-4 md:p-6 shadow-sm rounded-2xl mt-8 mb-8 hidden">
                <h2 class="text-xl md:text-2xl font-bold text-slate-800 mb-4 md:mb-6 border-b pb-4 text-center">نتائج جميع الطلاب</h2>
                <div class="overflow-x-auto overflow-y-auto rounded-xl border border-slate-200 bg-white" style="max-height: 600px;">
                    <table class="w-full text-right mobile-card-table" id="allStudentsTable">
                        <thead class="bg-slate-100 text-slate-700 border-b border-slate-200 sticky top-0 z-10 shadow-sm">
                            <tr>
                                <th class="p-3 md:p-4 font-bold">الترتيب</th>
                                <th class="p-3 md:p-4 font-bold">اسم الطالب</th>
                                <th class="p-3 md:p-4 font-bold">المعلم / الفصل</th>
                                <th class="p-3 md:p-4 font-bold">المجموع</th>
                                <th class="p-3 md:p-4 font-bold">النسبة</th>
                                <th class="p-3 md:p-4 font-bold">التقدير</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-100">${rankingHTML}</tbody>
                    </table>
                    <div id="noStudentsFound" class="hidden p-8 text-center text-slate-500 font-bold">لا يوجد طلاب مطابقين للبحث.</div>
                </div>
            </section>
        </div>
    </div>

    <button id="backToTop" class="fixed bottom-6 right-6 bg-slate-800 text-white p-3 rounded-full shadow-lg opacity-0 transition-opacity duration-300 z-50">
        <svg xmlns="http://www.w3.org/2000/svg" class="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 10l7-7m0 0l7 7m-7-7v18" />
        </svg>
    </button>

    <script>
        document.addEventListener('contextmenu', event => event.preventDefault());

        const backToTopBtn = document.getElementById('backToTop');
        window.addEventListener('scroll', () => {
            if (window.scrollY > 500) {
                backToTopBtn.classList.add('opacity-100');
                backToTopBtn.classList.remove('opacity-0');
            } else {
                backToTopBtn.classList.add('opacity-0');
                backToTopBtn.classList.remove('opacity-100');
            }
        });
        backToTopBtn.addEventListener('click', () => {
            window.scrollTo({ top: 0, behavior: 'smooth' });
        });

        const chartConfigs = ${JSON.stringify(chartConfigs)};
        const allStudentsData = ${JSON.stringify(validStudents.map(s => ({
            name: s.name, sheet: s.sheet, total: s.total, percentage: s.percentage
        })))};
        
        let chartsInitialized = {};

        function initializeCharts() {
            document.querySelectorAll('canvas[data-chart-idx]').forEach(canvas => {
                const idx = canvas.getAttribute('data-chart-idx');
                if (chartsInitialized[idx]) return;
                
                // Only initialize if visible
                if (canvas.offsetParent !== null) {
                    const config = chartConfigs[idx];
                    if (config && typeof Chart !== 'undefined') {
                        new Chart(canvas, {
                            type: 'pie',
                            data: {
                                labels: config.labels,
                                datasets: [{
                                    data: config.data,
                                    backgroundColor: ['#030504ff', '#3b82f6', '#eab308', '#f97316', '#ef4444'],
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
                                                let perc = ((val / config.total) * 100).toFixed(1);
                                                return label + val + ' (' + perc + '%)';
                                            }
                                        }
                                    }
                                }
                            }
                        });
                        chartsInitialized[idx] = true;
                    }
                }
            });
        }

        // Logic for role switching
        function setRole(role) {
            if (role === 'manager') {
                const managerId = prompt("الرجاء إدخال الرقم القومي للمدير:");
                if (managerId !== "28608052200699") {
                    alert("رقم قومي خاطئ. لا تملك صلاحية الدخول كمدير.");
                    return;
                }
            }

            // Hide landing screen
            document.getElementById('roleSelectionScreen').classList.add('hidden');
            // Show main application container
            document.getElementById('mainApp').classList.remove('hidden');

            // Reset all view sections
            document.getElementById('studentSearchView').classList.add('hidden');
            document.getElementById('teacherSelectView').classList.add('hidden');
            document.getElementById('dashboardContent').classList.add('hidden');
            document.getElementById('managerHeader').classList.add('hidden');
            document.getElementById('managerGlobalSearch').classList.add('hidden');
            document.getElementById('teacherRankingContainer').classList.add('hidden');
            document.getElementById('allStudentsRankingContainer').classList.add('hidden');

            if (role === 'manager') {
                document.getElementById('dashboardContent').classList.remove('hidden');
                document.getElementById('managerHeader').classList.remove('hidden');
                document.getElementById('managerGlobalSearch').classList.remove('hidden');
                document.getElementById('teacherRankingContainer').classList.remove('hidden');
                document.getElementById('allStudentsRankingContainer').classList.remove('hidden');
                // Show all teacher sections for manager
                document.querySelectorAll('.teacher-section').forEach(el => el.classList.remove('hidden'));
                setTimeout(initializeCharts, 100);
            } else if (role === 'teacher') {
                document.getElementById('teacherSelectView').classList.remove('hidden');
                document.getElementById('dashboardContent').classList.remove('hidden');
                // Hide all teacher sections until one is selected
                document.querySelectorAll('.teacher-section').forEach(el => el.classList.add('hidden'));
                
                const prevTeacher = localStorage.getItem('selectedTeacher_2025');
                if (prevTeacher) {
                    document.getElementById('teacherDropdown').value = prevTeacher;
                    onTeacherSelected();
                } else {
                    document.getElementById('teacherDropdown').value = ""; 
                }
            } else if (role === 'student') {
                document.getElementById('studentSearchView').classList.remove('hidden');
            }
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        function logout() {
            document.getElementById('mainApp').classList.add('hidden');
            document.getElementById('roleSelectionScreen').classList.remove('hidden');
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        function onTeacherSelected() {
            const selected = document.getElementById('teacherDropdown').value;
            if (!selected) return;

            const previouslySelectedTeacher = localStorage.getItem('selectedTeacher_2025');
            if (previouslySelectedTeacher && previouslySelectedTeacher !== selected) {
                alert("عذراً، لقد قمت بتسجيل الدخول كمعلم مسبقاً باسم مختلف. لا يمكنك تغيير الاسم.");
                document.getElementById('teacherDropdown').value = previouslySelectedTeacher;
                onTeacherSelected(); // Recursively call with the correct value
                return;
            } else {
                localStorage.setItem('selectedTeacher_2025', selected);
            }

            document.querySelectorAll('.teacher-section').forEach(el => {
                if (el.getAttribute('data-teacher-name') === selected) {
                    el.classList.remove('hidden');
                } else {
                    el.classList.add('hidden');
                }
            });
            if(selected) setTimeout(initializeCharts, 100);
        }

        function filterAllStudents() {
            const query = document.getElementById('managerSearchInput').value.toLowerCase();
            const rows = document.querySelectorAll('#allStudentsTable tbody tr');
            let found = false;
            rows.forEach(row => {
                // Name is in the second column (index 1) for mobile layout as well
                const nameCell = row.children[1];
                if (nameCell) {
                    const name = nameCell.textContent.toLowerCase();
                    if(name.includes(query)) {
                        row.style.display = '';
                        found = true;
                    } else {
                        row.style.display = 'none';
                    }
                }
            });
            document.getElementById('noStudentsFound').style.display = found ? 'none' : 'block';
        }

        function getGradeTextJS(p) {
            if(p >= 90) return '<span class="text-emerald-600 font-bold">امتياز</span>';
            if(p >= 80) return '<span class="text-blue-600 font-bold">جيد جدا</span>';
            if(p >= 70) return '<span class="text-yellow-600 font-bold">جيد</span>';
            if(p >= 50) return '<span class="text-orange-500 font-bold">مقبول</span>';
            return '<span class="text-red-600 font-bold">راسب</span>';
        }

        function executeStudentSearch() {
            const query = document.getElementById('studentNameInput').value.trim();
            if(query.length < 3) {
                alert("يرجى كتابة الاسم بشكل صحيح (ثلاثي أو رباعي).");
                return;
            }
            
            const student = allStudentsData.find(s => s.name.includes(query));
            if(!student) {
                document.getElementById('studentSearchError').classList.remove('hidden');
                return;
            }

            let searchedNames = JSON.parse(localStorage.getItem('searchedStudents_2025') || '[]');
            
            if (!searchedNames.includes(student.name)) {
                if (searchedNames.length >= 3) {
                    alert("عذراً، لقد استنفدت الحد الأقصى (3 مرات) للبحث عن طلاب مختلفين من هذا الجهاز.");
                    return;
                }
                searchedNames.push(student.name);
                localStorage.setItem('searchedStudents_2025', JSON.stringify(searchedNames));
            }
            
            document.getElementById('studentSearchError').classList.add('hidden');
            
            const p = (student.percentage * 100).toFixed(1);
            const gradeHtml = getGradeTextJS(student.percentage * 100);
            
            document.getElementById('studentResultCard').innerHTML = \`
                <h3 class="text-xl font-bold text-slate-800 border-b pb-3 mb-6 text-center">نتيجة الطالب</h3>
                <div class="grid grid-cols-1 md:grid-cols-2 gap-6 text-right">
                    <div class="bg-white p-4 rounded-xl border border-slate-100 shadow-sm"><span class="text-slate-500 text-sm block mb-1">الاسم</span><span class="font-bold text-lg text-slate-800">\${student.name}</span></div>
                    <div class="bg-white p-4 rounded-xl border border-slate-100 shadow-sm"><span class="text-slate-500 text-sm block mb-1">المعلم / الفصل</span><span class="font-bold text-lg text-slate-800">\${student.sheet}</span></div>
                    <div class="bg-white p-4 rounded-xl border border-slate-100 shadow-sm"><span class="text-slate-500 text-sm block mb-1">المجموع</span><span class="font-bold text-xl text-blue-700">\${student.total}</span></div>
                    <div class="bg-white p-4 rounded-xl border border-slate-100 shadow-sm"><span class="text-slate-500 text-sm block mb-1">النسبة</span><span class="bg-slate-200 text-slate-800 px-3 py-1 rounded text-lg font-bold inline-block">\${p}%</span></div>
                    <div class="col-span-1 md:col-span-2 bg-white p-4 rounded-xl border border-slate-100 shadow-sm text-center mt-2"><span class="text-slate-500 text-sm block mb-2">التقدير العام</span><div class="text-2xl">\${gradeHtml}</div></div>
                </div>
            \`;
            document.getElementById('studentResultCard').classList.remove('hidden');
        }
    </script>
</body>
</html>`;

  fs.writeFileSync('index.html', htmlContent);
  console.log('Dashboard generated successfully at index.html');
} catch (error) {
  console.error('Error reading excel file:', error);
}
