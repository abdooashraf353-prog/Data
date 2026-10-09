<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>المساعد الذكي للمبيعات | Raya Shop AI</title>
    
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;600;700;800&display=swap" rel="stylesheet">
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

    <style>
        body { 
            font-family: 'Cairo', sans-serif; 
            background-color: #0f172a; 
            color: #e2e8f0; 
        }
        
        .ai-loading { display: inline-flex; gap: 4px; align-items: center; }
        .ai-loading span { width: 6px; height: 6px; background-color: #38bdf8; border-radius: 50%; animation: bounce 1.4s infinite ease-in-out both; }
        .ai-loading span:nth-child(1) { animation-delay: -0.32s; }
        .ai-loading span:nth-child(2) { animation-delay: -0.16s; }
        
        @keyframes bounce { 0%, 80%, 100% { transform: scale(0); opacity: 0.3; } 40% { transform: scale(1); opacity: 1; } }

        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: #1e293b; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #3b82f6; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #60a5fa; }

        .glow-border { border: 1px solid rgba(56, 189, 248, 0.2); box-shadow: 0 0 15px rgba(56, 189, 248, 0.05); }
        
        /* تأثير ظهور العناصر بنعومة */
        .fade-in-up { animation: fadeInUp 0.6s ease-out forwards; opacity: 0; transform: translateY(10px); }
        .delay-1 { animation-delay: 0.2s; }
        .delay-2 { animation-delay: 0.4s; }
        .delay-3 { animation-delay: 0.6s; }
        @keyframes fadeInUp { to { opacity: 1; transform: translateY(0); } }
    </style>
</head>
<body class="antialiased min-h-screen flex flex-col selection:bg-cyan-500 selection:text-white">

    <!-- شريط علوي -->
    <div class="bg-slate-900 border-b border-slate-800 shadow-2xl py-6 relative overflow-hidden">
        <div class="absolute top-0 right-0 w-64 h-64 bg-cyan-500 rounded-full mix-blend-screen filter blur-[100px] opacity-10"></div>
        <div class="absolute bottom-0 left-0 w-64 h-64 bg-purple-500 rounded-full mix-blend-screen filter blur-[100px] opacity-10"></div>
        
        <div class="container mx-auto px-6 flex flex-col md:flex-row items-center justify-between relative z-10 gap-4">
            <div class="flex items-center gap-4">
                <div class="w-14 h-14 rounded-2xl bg-gradient-to-br from-cyan-400 to-blue-600 flex items-center justify-center shadow-lg shadow-cyan-500/30">
                    <i class="fa-solid fa-brain text-3xl text-white"></i>
                </div>
                <div>
                    <h1 class="text-3xl font-bold text-white tracking-wide">المساعد الذكي للمبيعات</h1>
                    <p class="text-sm text-cyan-400 font-semibold tracking-widest uppercase mt-1">Powered by Real AI</p>
                </div>
            </div>
            
            <!-- إعدادات مفتاح الذكاء الاصطناعي الحقيقي -->
            <div class="bg-slate-800/80 border border-slate-700 p-3 rounded-xl flex items-center gap-3 w-full md:w-auto">
                <i class="fa-solid fa-key text-yellow-400"></i>
                <input type="password" id="apiKeyInput" placeholder="أدخل مفتاح Gemini API..." class="bg-slate-900 border border-slate-600 text-sm rounded-lg px-3 py-2 w-64 text-white focus:outline-none focus:border-cyan-500">
                <a href="https://aistudio.google.com/app/apikey" target="_blank" title="احصل على مفتاح مجاني" class="text-xs text-cyan-400 hover:text-cyan-300 underline"><i class="fa-solid fa-circle-info"></i> مجاني</a>
            </div>
        </div>
    </div>

    <div class="container mx-auto px-4 py-8 flex-1 relative z-10">
        
        <!-- لوحة التحكم -->
        <div class="bg-slate-800 rounded-2xl glow-border p-6 mb-8 flex flex-col md:flex-row items-center justify-between gap-6 transition-all hover:shadow-cyan-500/10">
            <div>
                <h2 class="text-xl font-bold text-white mb-2"><i class="fa-solid fa-file-excel text-emerald-400 ml-2"></i>استيراد الأكواد من إكسيل</h2>
                <p class="text-sm text-slate-400">ارفع الشيت، وسيقوم الذكاء الاصطناعي بالبحث وكتابة تقرير مفصل لكل منتج.</p>
            </div>
            
            <div class="flex gap-4 w-full md:w-auto relative">
                <input type="file" id="excelFile" accept=".xlsx, .xls, .csv" class="absolute w-0 h-0 opacity-0 -z-10">
                <button type="button" onclick="document.getElementById('excelFile').click()" id="uploadBtn" class="bg-slate-700 hover:bg-slate-600 border border-slate-600 hover:border-emerald-500/50 text-white py-3 px-6 rounded-xl font-semibold transition-all flex items-center justify-center gap-3 flex-1 md:flex-none shadow-lg group">
                    <i id="uploadIcon" class="fa-solid fa-cloud-arrow-up text-emerald-400 group-hover:-translate-y-1 transition-transform"></i> 
                    <span id="uploadText">رفع ملف Excel</span>
                </button>

                <button type="button" onclick="addTestRow()" class="bg-gradient-to-r from-blue-600 to-cyan-500 hover:from-blue-500 hover:to-cyan-400 text-white py-3 px-6 rounded-xl font-semibold transition-all flex items-center justify-center gap-2 flex-1 md:flex-none shadow-lg shadow-blue-500/25">
                    <i class="fa-solid fa-bolt text-yellow-300"></i> تجربة سريعة
                </button>
            </div>
        </div>

        <!-- الجدول -->
        <div class="bg-slate-800 rounded-2xl glow-border overflow-hidden shadow-2xl">
            <div class="overflow-x-auto">
                <table class="w-full text-right border-collapse" id="productsTable">
                    <thead>
                        <tr class="bg-slate-900/50 text-slate-300 border-b border-slate-700">
                            <th class="py-5 px-6 font-bold w-1/5"><i class="fa-solid fa-barcode ml-2 text-slate-400"></i>كود المنتج</th>
                            <th class="py-5 px-6 font-bold w-3/5"><i class="fa-solid fa-robot ml-2 text-purple-400"></i>تقرير الذكاء الاصطناعي المفصل</th>
                            <th class="py-5 px-6 font-bold w-1/5 text-center"><i class="fa-solid fa-magnifying-glass ml-2 text-slate-400"></i>إجراء سريع</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody" class="divide-y divide-slate-700/50">
                        <tr id="emptyState">
                            <td colspan="3" class="py-16 text-center text-slate-500">
                                <div class="inline-flex items-center justify-center w-20 h-20 rounded-full bg-slate-800 border border-slate-700 mb-4 shadow-inner">
                                    <i class="fa-solid fa-box-open text-4xl text-slate-600"></i>
                                </div>
                                <h3 class="text-xl font-bold text-slate-400 mb-2">في انتظار المنتجات...</h3>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
        const tableBody = document.getElementById('tableBody');
        const emptyState = document.getElementById('emptyState');
        const apiKeyInput = document.getElementById('apiKeyInput');
        let rowCounter = 0;

        // حفظ واسترجاع مفتاح API
        apiKeyInput.value = localStorage.getItem('gemini_api_key') || '';
        apiKeyInput.addEventListener('change', () => localStorage.setItem('gemini_api_key', apiKeyInput.value));

        function addRow(itemCode) {
            if (emptyState) emptyState.style.display = 'none';
            rowCounter++;
            const rowId = `row-${rowCounter}`;
            const specsId = `specs-${rowCounter}`;
            
            const tr = document.createElement('tr');
            tr.id = rowId;
            tr.className = "hover:bg-slate-700/30 transition-colors duration-300";
            
            tr.innerHTML = `
                <td class="py-6 px-6 align-top">
                    <div class="inline-flex items-center gap-2 bg-slate-900 border border-slate-700 text-cyan-300 font-black px-4 py-3 rounded-xl uppercase tracking-widest shadow-inner text-lg">
                        <i class="fa-solid fa-tag text-slate-500 text-sm"></i>
                        ${itemCode}
                    </div>
                    <!-- زر توليد التقرير يدوي لتجنب حظر الـ API إذا كان هناك ضغط -->
                    <button onclick="startAiAnalysis('${itemCode}', '${specsId}')" class="mt-4 w-full bg-slate-700 hover:bg-slate-600 text-white text-xs py-2 rounded-lg border border-slate-500 flex items-center justify-center gap-2 transition-colors">
                        <i class="fa-solid fa-rotate-right"></i> إعادة التحليل
                    </button>
                </td>
                <td class="py-6 px-6 align-top">
                    <div id="${specsId}" class="ai-box bg-slate-900/60 rounded-xl p-6 border border-slate-700/50 shadow-inner min-h-[150px]">
                        <div class="flex items-center gap-3 text-cyan-400 font-bold">
                            <i class="fa-solid fa-robot text-2xl"></i>
                            <span class="text-lg">الذكاء الاصطناعي يكتب التقرير الآن...</span>
                            <div class="ai-loading ml-2"><span></span><span></span><span></span></div>
                        </div>
                    </div>
                </td>
                <td class="py-6 px-6 align-top text-center">
                    <button type="button" onclick="searchRayaShop('${itemCode}')" class="bg-slate-800 hover:bg-slate-700 text-white font-bold py-3 px-5 rounded-xl transition-all flex items-center justify-center gap-2 w-full max-w-[180px] mx-auto border border-slate-600 hover:border-cyan-500">
                        <i class="fa-brands fa-google text-blue-400"></i>
                        <span>بحث جوجل</span>
                    </button>
                </td>
            `;
            
            tableBody.appendChild(tr);
            startAiAnalysis(itemCode, specsId);
        }

        function addTestRow() {
            const sampleCodes = ['iPhone 15 Pro Max', 'Samsung S24 Ultra', 'LG OLED65C3'];
            const randomCode = sampleCodes[Math.floor(Math.random() * sampleCodes.length)];
            addRow(randomCode);
        }

        // قراءة الإكسيل
        document.getElementById('excelFile').addEventListener('change', function(e) {
            if (typeof XLSX === 'undefined') {
                alert("برجاء الانتظار حتى يتم تحميل أدوات النظام أو قم بتحديث الصفحة.");
                return;
            }
            const file = e.target.files[0];
            if (!file) return;

            const uploadIcon = document.getElementById('uploadIcon');
            const uploadText = document.getElementById('uploadText');
            uploadIcon.className = "fa-solid fa-spinner fa-spin text-emerald-400";
            uploadText.innerText = "جاري قراءة الملف...";

            const reader = new FileReader();
            reader.onload = function(evt) {
                try {
                    const data = new Uint8Array(evt.target.result);
                    const workbook = XLSX.read(data, { type: 'array' }); 
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
                    
                    let addedCount = 0;
                    jsonData.forEach((row, index) => {
                        const cellValue = String(row[0]).trim();
                        if(row.length > 0 && cellValue !== "") { 
                            if(index === 0 && (cellValue.toLowerCase().includes('code') || cellValue.includes('كود'))) return;
                            
                            // تأخير بسيط بين كل منتج والآخر حتى لا يتم حظر الـ API
                            setTimeout(() => addRow(cellValue), addedCount * 1500);
                            addedCount++;
                        }
                    });
                } catch (error) {
                    alert("❌ حدث خطأ في الملف.");
                } finally {
                    uploadIcon.className = "fa-solid fa-cloud-arrow-up text-emerald-400 group-hover:-translate-y-1 transition-transform";
                    uploadText.innerText = "رفع ملف Excel";
                    e.target.value = ''; 
                }
            };
            reader.readAsArrayBuffer(file);
        });

        // تشغيل الذكاء الاصطناعي (Gemini أو المحاكي المتقدم)
        async function startAiAnalysis(itemCode, elementId) {
            const targetElement = document.getElementById(elementId);
            const apiKey = apiKeyInput.value.trim();

            targetElement.innerHTML = `
                <div class="flex items-center gap-3 text-cyan-400 font-bold">
                    <i class="fa-solid fa-robot text-2xl"></i>
                    <span class="text-lg">جاري البحث والتحليل العميق...</span>
                    <div class="ai-loading ml-2"><span></span><span></span><span></span></div>
                </div>`;

            try {
                let aiData;
                if (apiKey) {
                    // الاتصال الحقيقي بـ Gemini API
                    aiData = await fetchFromGeminiAPI(itemCode, apiKey);
                } else {
                    // محاكي متقدم جداً للذكاء الاصطناعي إذا لم يضع المستخدم مفتاح
                    aiData = generateAdvancedMock(itemCode);
                    await new Promise(r => setTimeout(r, 2000)); // محاكاة وقت التحميل
                }
                
                renderDetailedReport(aiData, targetElement);

            } catch (error) {
                targetElement.innerHTML = `<div class="text-red-400 p-4 bg-red-900/20 rounded-lg"><i class="fa-solid fa-triangle-exclamation"></i> حدث خطأ في جلب البيانات: ${error.message} <br><span class="text-sm text-slate-400">تأكد من صحة مفتاح API أو جرب بدون مفتاح.</span></div>`;
            }
        }

        // دالة الاتصال المباشر بـ Google Gemini API
        async function fetchFromGeminiAPI(itemCode, apiKey) {
            const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
            const prompt = `أنت خبير مبيعات ومحلل تقني محترف في متجر إلكترونيات. ابحث عن المنتج صاحب الكود أو الموديل: "${itemCode}".
            قم بكتابة تقرير مفصل وجذاب باللغة العربية.
            أرجع النتيجة بصيغة JSON حصراً بهذا التنسيق وبدون أي علامات Markdown:
            {
                "summary": "اكتب فقرة كاملة ومفصلة من 3 إلى 4 أسطر تشرح فيها ماهية المنتج، فئته، وأهم ميزة تنافسية له.",
                "specs": ["مواصفة تفصيلية 1", "مواصفة تفصيلية 2", "مواصفة تفصيلية 3", "مواصفة تفصيلية 4", "مواصفة تفصيلية 5"],
                "salesTips": "اكتب فقرة مفصلة من 3 أسطر تشرح لمندوب المبيعات كيف يبيع هذا المنتج، وما هي النقاط التي يجب أن يركز عليها لإقناع العميل، وما هي الفئة المستهدفة."
            }`;

            const response = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: prompt }] }],
                    generationConfig: { temperature: 0.7 }
                })
            });

            if (!response.ok) throw new Error("فشل الاتصال بـ Gemini، قد يكون المفتاح غير صحيح أو تجاوزت الحد المسموح.");

            const data = await response.json();
            let text = data.candidates[0].content.parts[0].text;
            // تنظيف الـ JSON من أي إضافات يضعها الذكاء الاصطناعي
            text = text.replace(/```json/g, '').replace(/```/g, '').trim();
            
            return JSON.parse(text);
        }

        // رسم التقرير بتأثير احترافي (Fade In / Typing Feel)
        function renderDetailedReport(data, targetElement) {
            let specsHTML = data.specs.map(spec => `<li class="flex items-start gap-3 mb-3"><i class="fa-solid fa-check-double text-emerald-500 mt-1"></i><span class="text-slate-300 leading-relaxed">${spec}</span></li>`).join('');
            
            targetElement.innerHTML = `
                <div class="space-y-6">
                    <!-- الملخص -->
                    <div class="fade-in-up delay-1">
                        <h4 class="text-cyan-400 font-bold mb-2 flex items-center gap-2 text-lg border-b border-cyan-500/30 pb-2"><i class="fa-solid fa-pen-nib"></i> تقرير ووصف المنتج</h4>
                        <p class="text-slate-300 leading-relaxed text-base text-justify">${data.summary}</p>
                    </div>
                    
                    <!-- المواصفات -->
                    <div class="fade-in-up delay-2 bg-slate-800/50 p-4 rounded-xl border border-slate-700">
                        <h4 class="text-emerald-400 font-bold mb-3 flex items-center gap-2 text-lg"><i class="fa-solid fa-microchip"></i> المواصفات التقنية الدقيقة</h4>
                        <ul class="ml-2">
                            ${specsHTML}
                        </ul>
                    </div>

                    <!-- أسرار البيع -->
                    <div class="fade-in-up delay-3 bg-gradient-to-r from-indigo-900/40 to-purple-900/40 p-5 rounded-xl border border-indigo-500/40 relative overflow-hidden">
                        <div class="absolute top-0 right-0 w-1 bg-indigo-500 h-full shadow-[0_0_10px_#6366f1]"></div>
                        <h4 class="text-indigo-300 font-bold mb-2 flex items-center gap-2 text-lg pr-2"><i class="fa-solid fa-lightbulb text-yellow-400"></i> استراتيجية المبيعات والتسويق</h4>
                        <p class="text-slate-300 leading-relaxed pr-2 text-base text-justify">${data.salesTips}</p>
                    </div>
                </div>
            `;
        }

        // محاكي متقدم لإنتاج نصوص طويلة إذا لم يتم وضع مفتاح API
        function generateAdvancedMock(itemCode) {
            const codeUpper = itemCode.toUpperCase();
            
            if(codeUpper.includes('S24') || codeUpper.includes('S928') || codeUpper.includes('ULTRA')) {
                return {
                    summary: `يُعد هذا الهاتف تحفة فنية من سامسونج، حيث يمثل قمة التكنولوجيا في سلسلة الهواتف الرائدة (Flagship). يتميز بدمج الذكاء الاصطناعي (Galaxy AI) بشكل غير مسبوق في معالجة الصور والنصوص. الهيكل مصنوع من التيتانيوم القوي لضمان المتانة الفائقة، ويعتبر الخيار الأول لرجال الأعمال وصناع المحتوى بفضل القلم المدمج S-Pen الذي يعزز من الإنتاجية.`,
                    specs: [
                        "شاشة Dynamic AMOLED 2X مسطحة بحجم 6.8 بوصة، تقدم سطوعاً خرافياً يصل إلى 2600 شمعة مع طبقة حماية Gorilla Armor المقاومة للانعكاسات.",
                        "معالج Snapdragon 8 Gen 3 for Galaxy الأقوى في العالم، مصمم لتقديم أداء لا يضاهى في الألعاب والمهام المتعددة الثقيلة.",
                        "نظام تصوير احترافي بعدسة أساسية بدقة 200 ميجابكسل، وعدسة زووم بصري 5x بدقة 50 ميجابكسل مع قدرات تصوير ليلي (Nightography) استثنائية.",
                        "بطارية عملاقة بسعة 5000 مللي أمبير تدعم الشحن السريع بقوة 45 واط، والشحن اللاسلكي العكسي.",
                        "دعم كامل لميزات الذكاء الاصطناعي مثل الترجمة الفورية للمكالمات (Live Translate) وتحرير الصور المتقدم."
                    ],
                    salesTips: `ركز في حديثك مع العميل على "الإنتاجية وقوة التحمل". اشرح له كيف يمكن لميزات الذكاء الاصطناعي توفير وقته الثمين في عمله. قارن بين قوة التيتانيوم في هذا الموديل والألومنيوم في الموديلات السابقة. إذا كان العميل مهتماً بالتصوير، افتح له الكاميرا واعرض عليه الزووم البصري الجديد ونقاء الشاشة في الإضاءة العالية.`
                };
            } 
            else if (codeUpper.includes('IPHONE') || codeUpper.includes('PRO MAX')) {
                return {
                    summary: `أحدث إصدارات شركة آبل الرائدة، المصمم خصيصاً ليقدم تجربة مستخدم لا تشوبها شائبة بفضل التوافق التام بين نظام التشغيل iOS والهاردوير. يتميز هذا الهاتف بتصميم ثوري من التيتانيوم يجعله الأخف وزناً في فئته والأكثر صلابة. هذا الجهاز هو استثمار آمن يحتفظ بقيمته المادية بمرور الزمن ويوفر نظاماً بيئياً (Ecosystem) متكاملاً.`,
                    specs: [
                        "هيكل خارجي من التيتانيوم من الفئة الفضائية (Aerospace-grade titanium) مع حواف دائرية لراحة أكبر في المسك.",
                        "معالج A17 Pro بتكنولوجيا 3 نانومتر، يقدم أداء رسومي يضاهي منصات الألعاب المنزلية (Consoles) ويدعم تتبع الأشعة.",
                        "كاميرا رئيسية بدقة 48 ميجابكسل مع خيارات تقريب بصري تصل إلى 5x، وتصوير فيديو سينمائي مذهل يدعم Log video recording الاحترافي.",
                        "زر الإجراءات الجديد (Action Button) الذي يمكن تخصيصه لأكثر من وظيفة بضغطة واحدة.",
                        "التحول أخيراً إلى منفذ USB-C مع سرعات نقل بيانات فائقة الجودة للمحترفين."
                    ],
                    salesTips: `عميل الآيفون يبحث دائماً عن "البرستيج، الأمان، والتصوير". ركز على مدى خفة وزن الجهاز بفضل التيتانيوم مقارنة بالإصدارات السابقة. أكد له على إمكانيات الكاميرا الجبارة في تصوير الفيديو لو كان من صناع المحتوى، ولا تنسَ تذكيره بأن أجهزة آبل تحتفظ بقيمتها عند إعادة البيع (Resale Value) لتشجيعه على اتخاذ قرار الشراء فوراً.`
                };
            }
            else {
                return {
                    summary: `بناءً على تحليل كود الموديل (${itemCode})، يتبين أن هذا المنتج ينتمي لمجموعة الأجهزة الإلكترونية الحديثة التي تتمتع بطلب مستمر في الأسواق. تم تصميم هذا المنتج لتلبية احتياجات المستخدم العصري من حيث الكفاءة، المتانة، والتصميم الجذاب، وهو يأتي معتمدًا بضمان الوكيل المحلي لضمان راحة بال العميل بعد الشراء.`,
                    specs: [
                        "مصنع من خامات عالية الجودة تضمن عمراً افتراضياً طويلاً في ظل ظروف الاستخدام اليومي.",
                        "تصميم هندسي متطور (Ergonomic) يتناسب مع أحدث المعايير العالمية.",
                        "كفاءة عالية في استهلاك الطاقة (إذا كان جهازاً كهربائياً أو ذكياً).",
                        "دعم كامل من مراكز الصيانة المعتمدة وخدمات ما بعد البيع.",
                        "متوافق مع معايير الجودة والأمان الخاصة بالسوق المحلي."
                    ],
                    salesTips: `لإتمام بيع هذا المنتج، يجب عليك كبائع التركيز على "القيمة مقابل السعر (Value for Money)". اشرح للعميل كيف سيسهل هذا المنتج حياته اليومية. استخدم زر البحث بجوجل لاستعراض التقييمات الإيجابية لهذا الموديل، واعرض عليه خيارات التقسيط المتنوعة لدى Raya لتسهيل قرار الشراء وتقليل العبء المادي.`
                };
            }
        }

        function searchRayaShop(itemCode) {
            const googleUrl = `https://www.google.com/search?q=${encodeURIComponent(itemCode + ' specifications مواصفات')}`;
            window.open(googleUrl, '_blank');
        }
    </script>
</body>
</html>
