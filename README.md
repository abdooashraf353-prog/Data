<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>نظام متابعة العملاء المتقدم | Raya CRM</title>
    
    <!-- Firebase SDK v8 -->
    <script src="https://www.gstatic.com/firebasejs/8.10.1/firebase-app.js"></script>
    <script src="https://www.gstatic.com/firebasejs/8.10.1/firebase-database.js"></script>
    
    <!-- SheetJS (XLSX) for Excel parsing and writing -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    
    <!-- Google Fonts: Cairo -->
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet">
    
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

    <style>
        body {
            font-family: 'Cairo', sans-serif;
            background-color: #0b0f19;
            color: #f1f5f9;
            min-height: 100vh;
        }
        .glass-panel {
            background: rgba(17, 24, 39, 0.85);
            backdrop-filter: blur(16px);
            border: 1px solid rgba(255, 255, 255, 0.07);
            box-shadow: 0 15px 35px -10px rgba(0, 0, 0, 0.5);
        }
        .glass-input, .glass-select {
            background: rgba(15, 23, 42, 0.7);
            border: 1px solid rgba(255, 255, 255, 0.1);
            color: #fff;
            transition: all 0.3s ease;
        }
        .glass-input:focus, .glass-select:focus {
            outline: none;
            border-color: #0ea5e9;
            box-shadow: 0 0 0 2px rgba(14, 165, 233, 0.2);
        }
        .glass-select option {
            background-color: #0f172a;
            color: #fff;
        }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: rgba(15, 23, 42, 0.6); }
        ::-webkit-scrollbar-thumb { background: #0ea5e9; border-radius: 4px; }
    </style>
</head>
<body class="p-4 md:p-6">

    <div class="max-w-[98%] mx-auto space-y-6">
        
        <!-- Header -->
        <div class="glass-panel p-6 rounded-3xl flex flex-col md:flex-row justify-between items-center gap-4">
            <div>
                <h1 class="text-2xl md:text-3xl font-black text-transparent bg-clip-text bg-gradient-to-r from-sky-400 to-blue-600">
                    نظام متابعة وعروض العملاء (Raya CRM)
                </h1>
                <p class="text-slate-400 text-xs md:text-sm mt-1">إدارة شيتات العملاء، تسجيل المتابعات الحية، والإرسال الفوري عبر واتساب</p>
            </div>
            
            <!-- Employee ID Box & Actions -->
            <div class="flex items-center gap-3 flex-wrap">
                <!-- Employee ID Input Box -->
                <div class="glass-input px-4 py-2 rounded-2xl flex items-center gap-2 border border-sky-500/30">
                    <i class="fa-solid fa-id-badge text-sky-400"></i>
                    <div class="flex flex-col">
                        <span class="text-[10px] text-slate-400 font-semibold">ID الموظف الحالي:</span>
                        <input type="text" id="employeeIdInput" oninput="saveEmployeeId()" placeholder="أدخل ID الموظف..." 
                               class="bg-transparent text-white text-xs font-bold focus:outline-none w-28">
                    </div>
                </div>

                <!-- Export Excel Button -->
                <button onclick="exportToExcel()" class="bg-gradient-to-r from-sky-600 to-blue-600 hover:from-sky-500 hover:to-blue-500 text-white font-bold py-3 px-4 rounded-2xl shadow-lg transition transform hover:scale-105 flex items-center gap-2 text-xs">
                    <i class="fa-solid fa-file-excel text-base"></i>
                    <span>تصدير إكسيل</span>
                </button>

                <!-- Import Excel Button -->
                <label class="cursor-pointer bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-bold py-3 px-4 rounded-2xl shadow-lg transition transform hover:scale-105 flex items-center gap-2 text-xs">
                    <i class="fa-solid fa-file-arrow-up text-base"></i>
                    <span>استيراد إكسيل</span>
                    <input type="file" id="excelFile" accept=".xlsx, .xls, .csv" class="hidden" onchange="handleExcelUpload(event)">
                </label>

                <button onclick="clearAllCustomers()" class="bg-red-500/15 hover:bg-red-500/25 text-red-400 border border-red-500/20 p-3 rounded-2xl transition" title="حذف كل العملاء">
                    <i class="fa-solid fa-trash-can"></i>
                </button>
            </div>
        </div>

        <!-- Main Layout: Grid -->
        <div class="grid grid-cols-1 xl:grid-cols-4 gap-6">
            
            <!-- Left 3 Columns: Table & Filters -->
            <div class="xl:col-span-3 space-y-6">
                
                <!-- Search & Filters Bar -->
                <div class="glass-panel p-4 rounded-2xl flex flex-col lg:flex-row gap-4 justify-between items-center">
                    <!-- Search Input -->
                    <div class="relative w-full lg:w-72">
                        <span class="absolute inset-y-0 right-0 flex items-center pr-4 text-slate-400">
                            <i class="fa-solid fa-magnifying-glass"></i>
                        </span>
                        <input type="text" id="searchInput" oninput="filterTable()" placeholder="ابحث بالاسم، الموبايل، الرقم القومي..." 
                               class="glass-input w-full rounded-xl py-2.5 pr-11 pl-4 text-xs">
                    </div>

                    <!-- Dropdowns Filters -->
                    <div class="flex flex-wrap gap-3 w-full lg:w-auto items-center">
                        <div class="flex items-center gap-2">
                            <span class="text-xs text-slate-300 font-semibold">الحالة:</span>
                            <select id="statusFilter" onchange="filterTable()" class="glass-select rounded-xl py-2 px-3 text-xs">
                                <option value="all">كل الحالات</option>
                            </select>
                        </div>

                        <div class="flex items-center gap-2">
                            <span class="text-xs text-slate-300 font-semibold">الشركة:</span>
                            <select id="companyFilter" onchange="filterTable()" class="glass-select rounded-xl py-2 px-3 text-xs">
                                <option value="all">كل الشركات</option>
                            </select>
                        </div>

                        <div class="flex gap-1.5 overflow-x-auto pb-1 lg:pb-0">
                            <button onclick="setFollowupFilter('all')" id="ff-all" class="bg-sky-600 text-white px-3 py-1.5 rounded-xl text-xs font-bold transition">الكل</button>
                            <button onclick="setFollowupFilter('مهتم')" id="ff-مهتم" class="bg-slate-800 text-emerald-400 hover:bg-slate-700 px-3 py-1.5 rounded-xl text-xs font-bold transition">🟢 مهتم</button>
                            <button onclick="setFollowupFilter('بيفكر')" id="ff-بيفكر" class="bg-slate-800 text-amber-400 hover:bg-slate-700 px-3 py-1.5 rounded-xl text-xs font-bold transition">🟡 بيفكر</button>
                            <button onclick="setFollowupFilter('مش مهتم')" id="ff-مش مهتم" class="bg-slate-800 text-rose-400 hover:bg-slate-700 px-3 py-1.5 rounded-xl text-xs font-bold transition">🔴 مش مهتم</button>
                        </div>
                    </div>
                </div>

                <!-- Customers Table -->
                <div class="glass-panel rounded-2xl overflow-hidden shadow-2xl">
                    <div class="overflow-x-auto">
                        <table class="w-full text-right border-collapse text-xs md:text-sm whitespace-nowrap">
                            <thead>
                                <tr class="bg-slate-900/90 text-sky-400 border-b border-slate-800">
                                    <th class="p-3">#</th>
                                    <th class="p-3">الرقم القومي</th>
                                    <th class="p-3">الإسم</th>
                                    <th class="p-3">الموبايل</th>
                                    <th class="p-3">الحالة الأساسية</th>
                                    <th class="p-3">الشركة</th>
                                    <th class="p-3">الملاحظات</th>
                                    <th class="p-3 bg-slate-900/50 text-center">حالة المتابعة</th>
                                    <th class="p-3 bg-slate-900/50 text-center">التعليق والموظف</th>
                                    <th class="p-3 text-center">واتساب</th>
                                </tr>
                            </thead>
                            <tbody id="customersTableBody" class="text-slate-200">
                                <tr>
                                    <td colspan="10" class="text-center p-10 text-slate-400">
                                        <i class="fa-solid fa-folder-open text-4xl mb-3 block text-sky-500"></i>
                                        لا توجد بيانات حالياً. قم باستيراد شيت إكسيل للبدء.
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- Right Column: Offers Message Box & Broadcast -->
            <div class="xl:col-span-1 space-y-4">
                <div class="glass-panel p-6 rounded-3xl sticky top-6 border-t-4 border-t-sky-400 space-y-4">
                    <div class="flex items-center justify-between">
                        <h3 class="font-bold text-base text-white flex items-center gap-2">
                            <i class="fa-solid fa-bullhorn text-sky-400"></i>
                            <span>مكان رسالة الواتساب</span>
                        </h3>
                        <span id="saveStatus" class="text-xs text-emerald-400 hidden font-semibold">تم الحفظ ✓</span>
                    </div>
                    <p class="text-xs text-slate-400 leading-relaxed">
                        اكتب رسالتك هنا تماماً كما تريدها. سيتم إرسالها للعميل مضافاً إليها اسمه في البداية فقط دون أي تعديل على النص الخاص بك.
                    </p>
                    
                    <textarea id="offersBox" rows="9" oninput="saveOffers()" 
                              class="glass-input w-full rounded-2xl p-4 text-sm leading-relaxed resize-none" 
                              placeholder="اكتب رسالتك هنا..."></textarea>
                    
                    <!-- Broadcast Button (Send to All) -->
                    <button onclick="sendWhatsAppToAll()" class="w-full bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-bold py-3 px-4 rounded-2xl shadow-lg transition flex items-center justify-center gap-2 text-xs">
                        <i class="fa-brands fa-whatsapp text-lg"></i>
                        <span>إرسال الرسالة لكل العملاء (Broadcast)</span>
                    </button>

                    <div class="bg-sky-500/10 border border-sky-500/20 rounded-xl p-3 text-xs text-sky-300">
                        <i class="fa-solid fa-circle-info ml-1"></i> يتم الحفظ الفوري في قاعدة البيانات لكل تعديلاتك على النص.
                    </div>
                </div>
            </div>

        </div>
    </div>

    <!-- Toast Notification -->
    <div id="toast" class="fixed bottom-[-100px] left-1/2 -translate-x-1/2 glass-panel px-6 py-3 rounded-full flex items-center gap-3 shadow-2xl border-l-4 border-l-emerald-500 transition-all duration-300 z-50">
        <i class="fa-solid fa-circle-check text-emerald-400 text-xl" id="toastIcon"></i>
        <span id="toastMsg" class="font-bold text-white text-sm"></span>
    </div>

    <!-- Script Application Logic -->
    <script>
        // Firebase Configuration
        const firebaseConfig = {
            apiKey: "AIzaSyAUfiVn9d6ssAXyJebJZE9lHzEXIzf38dc",
            authDomain: "raya-a4bb5.firebaseapp.com",
            databaseURL: "https://raya-a4bb5-default-rtdb.firebaseio.com",
            projectId: "raya-a4bb5",
            storageBucket: "raya-a4bb5.firebasestorage.app",
            messagingSenderId: "698811790975",
            appId: "1:698811790975:web:c1187a6b2a927a24e6714a",
            measurementId: "G-GHHDTZSC5G"
        };
        
        firebase.initializeApp(firebaseConfig);
        const db = firebase.database();

        let allCustomers = [];
        let currentFilteredCustomers = [];
        let currentFollowupFilter = 'all';

        window.onload = function() {
            const savedEmpId = localStorage.getItem('raya_employee_id');
            if(savedEmpId) {
                document.getElementById('employeeIdInput').value = savedEmpId;
            }

            db.ref('settings/currentOffers').on('value', snapshot => {
                if(snapshot.exists()) {
                    document.getElementById('offersBox').value = snapshot.val();
                }
            });

            db.ref('customers').on('value', snapshot => {
                allCustomers = [];
                if(snapshot.exists()) {
                    snapshot.forEach(child => {
                        allCustomers.push({ firebaseKey: child.key, ...child.val() });
                    });
                }
                updateDropdowns();
                filterTable();
            });
        };

        function saveEmployeeId() {
            const empId = document.getElementById('employeeIdInput').value.trim();
            localStorage.setItem('raya_employee_id', empId);
        }

        let saveTimeout;
        function saveOffers() {
            const text = document.getElementById('offersBox').value;
            clearTimeout(saveTimeout);
            saveTimeout = setTimeout(() => {
                db.ref('settings/currentOffers').set(text).then(() => {
                    const status = document.getElementById('saveStatus');
                    status.classList.remove('hidden');
                    setTimeout(() => status.classList.add('hidden'), 2000);
                });
            }, 500);
        }

        function handleExcelUpload(e) {
            const file = e.target.files[0];
            if(!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    if(json.length === 0) {
                        showToast('الملف فارغ', 'error');
                        return;
                    }

                    let headerRowIndex = 0;
                    for(let i = 0; i < Math.min(json.length, 5); i++) {
                        let rowStr = JSON.stringify(json[i]);
                        if(rowStr.includes('الإسم') || rowStr.includes('الاسم') || rowStr.includes('الرقم القومي')) {
                            headerRowIndex = i;
                            break;
                        }
                    }

                    const dataRows = json.slice(headerRowIndex + 1);
                    const formattedCustomers = [];

                    dataRows.forEach((row, index) => {
                        if(!row || row.length === 0) return;
                        
                        let cId = row[0] || (index + 1);
                        let cNid = row[1] || '';
                        let cName = row[2] || '';
                        let cPhone = row[3] || '';
                        let cCreated = row[4] || '-';
                        let cUpdated = row[5] || '-';
                        let cStatus = row[6] || 'Active';
                        let cCompany = row[7] || 'بدون';
                        let cNotes = row[8] || '-';

                        if (!cName || cName === 'الإسم' || cName === 'الاسم') return;

                        let companyVal = String(cCompany).trim();
                        if (!companyVal || companyVal === 'undefined' || companyVal === 'NaN' || companyVal === '') {
                            companyVal = 'بدون';
                        }

                        formattedCustomers.push({
                            id: cId,
                            nationalId: String(cNid).trim(),
                            name: String(cName).trim(),
                            phone: String(cPhone).trim(),
                            createdAt: String(cCreated),
                            updatedAt: String(cUpdated),
                            status: String(cStatus).trim(),
                            company: companyVal,
                            notes: String(cNotes) !== 'undefined' && String(cNotes) !== 'NaN' ? String(cNotes) : '-',
                            followupComment: '',
                            followupStatus: 'غير محدد',
                            followupDate: '-',
                            employeeId: '-'
                        });
                    });

                    if(formattedCustomers.length === 0) {
                        showToast('لم يتم العثور على بيانات عملاء في الملف', 'error');
                        return;
                    }

                    db.ref('customers').set(formattedCustomers).then(() => {
                        showToast(`تم استيراد ${formattedCustomers.length} عميل بنجاح!`);
                    });

                } catch (err) {
                    console.error(err);
                    showToast('حدث خطأ أثناء قراءة ملف الإكسيل', 'error');
                }
            };
            reader.readAsArrayBuffer(file);
        }

        function exportToExcel() {
            if (!currentFilteredCustomers || currentFilteredCustomers.length === 0) {
                showToast('لا توجد بيانات لتصديرها', 'error');
                return;
            }

            const exportData = currentFilteredCustomers.map((c, index) => ({
                '#': index + 1,
                'الرقم القومي': c.nationalId || '-',
                'الإسم': c.name || '-',
                'الموبايل': c.phone || '-',
                'الحالة الأساسية': c.status || '-',
                'الشركة': c.company || 'بدون',
                'الملاحظات': c.notes || '-',
                'حالة المتابعة': c.followupStatus || 'غير محدد',
                'التعليق': c.followupComment || '',
                'ID الموظف': c.employeeId || '-',
                'وقت المتابعة': c.followupDate || '-'
            }));

            const worksheet = XLSX.utils.json_to_sheet(exportData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "تقرير المتابعة");

            const dateStr = new Date().toISOString().slice(0, 10);
            XLSX.writeFile(workbook, `Raya_CRM_Report_${dateStr}.xlsx`);
            showToast('تم تصدير شيت الإكسيل بنجاح!');
        }

        function updateDropdowns() {
            const statusSet = new Set();
            const companySet = new Set();

            allCustomers.forEach(c => {
                if(c.status && c.status !== '-') statusSet.add(c.status);
                if(c.company && c.company !== '-') companySet.add(c.company);
                else companySet.add('بدون');
            });

            const statusSelect = document.getElementById('statusFilter');
            const companySelect = document.getElementById('companyFilter');

            const currentStatusVal = statusSelect.value;
            const currentCompanyVal = companySelect.value;

            statusSelect.innerHTML = '<option value="all">كل الحالات</option>';
            statusSet.forEach(st => {
                statusSelect.innerHTML += `<option value="${st}">${st}</option>`;
            });
            statusSelect.value = currentStatusVal;

            companySelect.innerHTML = '<option value="all">كل الشركات</option><option value="بدون">بدون</option>';
            companySet.forEach(comp => {
                if(comp !== 'بدون') {
                    companySelect.innerHTML += `<option value="${comp}">${comp}</option>`;
                }
            });
            companySelect.value = currentCompanyVal;
        }

        function updateCustomerFollowup(firebaseKey, field, value) {
            const empId = document.getElementById('employeeIdInput').value.trim() || 'غير محدد';
            const now = new Date();
            const dateStr = now.toLocaleDateString('ar-EG') + ' ' + now.toLocaleTimeString('ar-EG', {hour: '2-digit', minute:'2-digit'});
            
            const updates = {};
            updates[field] = value;
            updates['followupDate'] = dateStr;
            updates['employeeId'] = empId;

            db.ref('customers/' + firebaseKey).update(updates).then(() => {
                const cust = allCustomers.find(c => c.firebaseKey === firebaseKey);
                if(cust) {
                    cust[field] = value;
                    cust['followupDate'] = dateStr;
                    cust['employeeId'] = empId;
                }
                showToast('تم حفظ التعديل بواسطة (ID: ' + empId + ')');
            });
        }

        function renderTable(customers) {
            const tbody = document.getElementById('customersTableBody');
            tbody.innerHTML = '';

            if(customers.length === 0) {
                tbody.innerHTML = `
                    <tr>
                        <td colspan="10" class="text-center p-10 text-slate-400">
                            <i class="fa-solid fa-folder-open text-4xl mb-3 block text-sky-500"></i>
                            لا توجد بيانات مطابقة للبحث أو الفلترة الحالية.
                        </td>
                    </tr>`;
                return;
            }

            customers.forEach((c, index) => {
                let phoneDisplay = c.phone && c.phone !== '-' && c.phone !== 'undefined' ? c.phone : '<span class="text-slate-500">لا يوجد</span>';
                let nidDisplay = c.nationalId && c.nationalId !== '-' && c.nationalId !== 'undefined' ? c.nationalId : '<span class="text-slate-500">لا يوجد</span>';
                
                let statusBg = 'bg-slate-800 text-slate-300 border border-slate-700';
                let lowerStatus = (c.status || '').toLowerCase();
                if(lowerStatus.includes('active') || lowerStatus.includes('approved') || lowerStatus.includes('نشط')) {
                    statusBg = 'bg-emerald-500/15 text-emerald-400 border border-emerald-500/30';
                } else if(lowerStatus.includes('reject') || lowerStatus.includes('مرفوض')) {
                    statusBg = 'bg-rose-500/15 text-rose-400 border border-rose-500/30';
                } else if(lowerStatus.includes('pending') || lowerStatus.includes('معلق')) {
                    statusBg = 'bg-amber-500/15 text-amber-400 border border-amber-500/30';
                }

                let fStatus = c.followupStatus || 'غير محدد';
                let companyDisplay = c.company && c.company !== '' ? c.company : 'بدون';

                tbody.innerHTML += `
                    <tr class="border-b border-slate-800/60 hover:bg-slate-900/40 transition">
                        <td class="p-3 font-bold text-slate-400">${index + 1}</td>
                        <td class="p-3 tracking-widest text-xs font-mono" dir="ltr">${nidDisplay}</td>
                        <td class="p-3 font-bold text-white max-w-[150px] truncate" title="${c.name}">${c.name}</td>
                        <td class="p-3 tracking-wider font-mono" dir="ltr">${phoneDisplay}</td>
                        <td class="p-3">
                            <span class="inline-block px-3 py-1 rounded-xl text-xs font-bold whitespace-nowrap ${statusBg}">${c.status}</span>
                        </td>
                        <td class="p-3 text-slate-300 font-semibold">${companyDisplay}</td>
                        <td class="p-3 text-slate-400 text-xs max-w-[180px] truncate" title="${c.notes}">${c.notes}</td>
                        
                        <!-- Followup Status Options -->
                        <td class="p-3 bg-slate-900/30 text-center">
                            <div class="flex items-center justify-center gap-1.5">
                                <button onclick="updateCustomerFollowup('${c.firebaseKey}', 'followupStatus', 'مهتم')" 
                                        class="px-2.5 py-1 rounded-lg text-xs font-bold transition ${fStatus === 'مهتم' ? 'bg-emerald-600 text-white shadow-lg shadow-emerald-600/30 ring-2 ring-emerald-400' : 'bg-slate-800 text-slate-400 hover:bg-slate-700'}">
                                    مهتم
                                </button>
                                <button onclick="updateCustomerFollowup('${c.firebaseKey}', 'followupStatus', 'بيفكر')" 
                                        class="px-2.5 py-1 rounded-lg text-xs font-bold transition ${fStatus === 'بيفكر' ? 'bg-amber-600 text-white shadow-lg shadow-amber-600/30 ring-2 ring-amber-400' : 'bg-slate-800 text-slate-400 hover:bg-slate-700'}">
                                    بيفكر
                                </button>
                                <button onclick="updateCustomerFollowup('${c.firebaseKey}', 'followupStatus', 'مش مهتم')" 
                                        class="px-2.5 py-1 rounded-lg text-xs font-bold transition ${fStatus === 'مش مهتم' ? 'bg-rose-600 text-white shadow-lg shadow-rose-600/30 ring-2 ring-rose-400' : 'bg-slate-800 text-slate-400 hover:bg-slate-700'}">
                                    مش مهتم
                                </button>
                            </div>
                        </td>

                        <!-- Comment Box, Employee ID & Date -->
                        <td class="p-3 bg-slate-900/30 min-w-[200px]">
                            <div class="space-y-1">
                                <input type="text" value="${c.followupComment || ''}" 
                                       onchange="updateCustomerFollowup('${c.firebaseKey}', 'followupComment', this.value)"
                                       placeholder="اكتب تعليقك هنا..." 
                                       class="glass-input w-full rounded-lg px-3 py-1.5 text-xs">
                                <div class="flex justify-between items-center text-[10px] text-slate-400 font-mono">
                                    <span class="text-sky-400 font-bold">ID: ${c.employeeId || '-'}</span>
                                    <span><i class="fa-regular fa-clock ml-1"></i>${c.followupDate || 'غير مسجل'}</span>
                                </div>
                            </div>
                        </td>

                        <!-- WhatsApp Button -->
                        <td class="p-3 text-center">
                            <button onclick="openWhatsApp('${c.phone}', '${c.name}')" 
                                    class="bg-emerald-600 hover:bg-emerald-500 text-white px-3.5 py-2 rounded-xl font-bold text-xs shadow-lg shadow-emerald-600/20 transition flex items-center justify-center gap-1.5 mx-auto">
                                <i class="fa-brands fa-whatsapp text-sm"></i>
                                <span>رسالة</span>
                            </button>
                        </td>
                    </tr>
                `;
            });
        }

        function filterTable() {
            const query = document.getElementById('searchInput').value.toLowerCase().trim();
            const selectedStatus = document.getElementById('statusFilter').value;
            const selectedCompany = document.getElementById('companyFilter').value;
            
            currentFilteredCustomers = allCustomers.filter(c => {
                const matchQuery = (c.name && c.name.toLowerCase().includes(query)) ||
                                   (c.phone && c.phone.toLowerCase().includes(query)) ||
                                   (c.nationalId && c.nationalId.toLowerCase().includes(query));
                
                const matchStatus = selectedStatus === 'all' || c.status === selectedStatus;
                
                let matchCompany = true;
                if (selectedCompany !== 'all') {
                    if (selectedCompany === 'بدون') {
                        matchCompany = !c.company || c.company === 'بدون' || c.company === '-' || c.company === 'غير محددة';
                    } else {
                        matchCompany = c.company === selectedCompany;
                    }
                }

                const matchFollowup = currentFollowupFilter === 'all' || c.followupStatus === currentFollowupFilter;

                return matchQuery && matchStatus && matchCompany && matchFollowup;
            });

            renderTable(currentFilteredCustomers);
        }

        function setFollowupFilter(status) {
            currentFollowupFilter = status;
            
            ['all', 'مهتم', 'بيفكر', 'مش مهتم'].forEach(s => {
                const btn = document.getElementById(`ff-${s}`);
                if(btn) {
                    if(s === status) {
                        btn.className = 'bg-sky-600 text-white px-3 py-1.5 rounded-xl text-xs font-bold transition';
                    } else {
                        btn.className = 'bg-slate-800 text-slate-300 hover:bg-slate-700 px-3 py-1.5 rounded-xl text-xs font-bold transition';
                    }
                }
            });

            filterTable();
        }

        // بناء الرسالة بحيث تأخذ النص المكتوب في الخانة حرفياً مع إضافة اسم العميل بشكل لائق فقط
        function getFinalWhatsAppMessage(customerName) {
            let userMessage = document.getElementById('offersBox').value.trim();
            if(!userMessage) {
                userMessage = "أهلاً بك معنا!";
            }
            // إضافة اسم العميل بشكل منسق في الأعلى ثم رسالتك دون أي تعديل إضافي
            return `أهلاً بك يا أ/ *${customerName}* 👋\n\n${userMessage}`;
        }

        let whatsappWindow = null;
        function openWhatsApp(phone, customerName) {
            if(!phone || phone.length < 8 || phone === '-' || phone === 'undefined') {
                showToast('رقم هاتف هذا العميل غير متوفر أو غير صالح', 'error');
                return;
            }

            let cleanPhone = phone.replace(/[^0-9]/g, '');
            if(cleanPhone.startsWith('01')) {
                cleanPhone = '20' + cleanPhone;
            }

            let message = getFinalWhatsAppMessage(customerName);
            let encodedUrl = `https://wa.me/${cleanPhone}?text=${encodeURIComponent(message)}`;
            
            if (!whatsappWindow || whatsappWindow.closed) {
                whatsappWindow = window.open(encodedUrl, 'whatsapp_chat');
            } else {
                whatsappWindow.location.href = encodedUrl;
                whatsappWindow.focus();
            }
        }

        function sendWhatsAppToAll() {
            if(currentFilteredCustomers.length === 0) {
                showToast('لا توجد عملاء للإرسال إليهم', 'error');
                return;
            }

            if(confirm(`هل تريد إرسال الرسالة لـ ${currentFilteredCustomers.length} عميل في القائمة الحالية؟`)) {
                let count = 0;
                currentFilteredCustomers.forEach((c, index) => {
                    if(c.phone && c.phone.length >= 8 && c.phone !== '-' && c.phone !== 'undefined') {
                        setTimeout(() => {
                            let cleanPhone = c.phone.replace(/[^0-9]/g, '');
                            if(cleanPhone.startsWith('01')) cleanPhone = '20' + cleanPhone;
                            
                            let message = getFinalWhatsAppMessage(c.name);
                            let url = `https://wa.me/${cleanPhone}?text=${encodeURIComponent(message)}`;
                            window.open(url, '_blank');
                        }, index * 900);
                        count++;
                    }
                });
                showToast(`جاري فتح محادثات الواتساب لـ ${count} عميل...`);
            }
        }

        function clearAllCustomers() {
            if(confirm('هل أنت متأكد من رغبتك في حذف جميع بيانات العملاء؟')) {
                db.ref('customers').remove().then(() => {
                    showToast('تم مسح جميع العملاء بنجاح');
                });
            }
        }

        function showToast(message, type = 'success') {
            const toast = document.getElementById('toast');
            document.getElementById('toastMsg').textContent = message;
            toast.style.borderLeftColor = type === 'success' ? '#10b981' : '#f43f5e';
            document.getElementById('toastIcon').className = type === 'success' ? 'fa-solid fa-circle-check text-emerald-400 text-xl' : 'fa-solid fa-circle-xmark text-rose-400 text-xl';
            toast.classList.remove('bottom-[-100px]');
            toast.classList.add('bottom-6');
            setTimeout(() => {
                toast.classList.remove('bottom-6');
                toast.classList.add('bottom-[-100px]');
            }, 3000);
        }
    </script>
</body>
</html>
