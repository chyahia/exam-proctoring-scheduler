// =================================================================================
// --- Global Variables and Initial Setup ---
// =================================================================================

let eventSource = null;
let selectedProfessorForAssign = null;
let selectedSubjectsForAssign = [];
let availableLevels = [];
let availableHalls = [];
let availableProfessors = [];
let availableSubjects = [];
let examDayCounter = 0;
let lastGeneratedSchedule = null;
let workloadChartInstance = null;
let customTargetPatterns = []; 

// --- دالة جديدة لإظهار رسائل التأكيد ---
function showNotification(message, type = 'success') {
    const notification = document.getElementById('notification-area');
    if (!notification) return;

    notification.textContent = message;
    notification.className = type; // 'success' or 'error'

    // إظهار الرسالة
    setTimeout(() => {
        notification.classList.add('show');
    }, 10); // تأخير بسيط لبدء الانتقال

    // إخفاء الرسالة بعد 3 ثوانٍ
    setTimeout(() => {
        notification.classList.remove('show');
    }, 3000);
}

document.addEventListener('DOMContentLoaded', () => {
    setupHeaderButtons();
    setupFormListeners();
    loadInitialDataForSettings();
    setupExamScheduleBuilder();
    setupGenerationListener();
    setupBackupRestoreListeners();
    setupCustomTargetListeners();
    setupDataImportExportListeners();
    setupBalancingStrategyListener(); 
});

function setupHeaderButtons() {
    document.getElementById('about-button').addEventListener('click', () => {
        alert('(GPLv3) # Copyright (C) 2025 CHAIB YAHIA');
    });

    document.getElementById('help-button').addEventListener('click', () => {
        const helpText = `
نظرة عامة على مراحل عمل البرنامج:

المرحلة 1: إدخال البيانات الأساسية
- إضافة كل البيانات (أساتذة، قاعات، مواد...) سواء بشكل يدوي أو عبر استيراد ملف Excel.

المرحلة 2: عرض وإدارة البيانات
- مراجعة وتعديل وحذف كل البيانات التي تم إدخالها.

المرحلة 3: إعداد قيود الحراسة
- ربط الأساتذة بالمواد، وتخصيص القاعات للمستويات، وتحديد أنماط وأيام غياب الأساتذة.

المرحلة 4: إعداد جدول الامتحانات
- تحديد تواريخ وفترات الامتحانات لكل مستوى دراسي.

المرحلة 5: إنشاء وتصدير الجداول
- ضبط إعدادات الخوارزمية النهائية ثم إنشاء وتصدير جداول الحراسة.

المرحلة 6: النسخ الاحتياطي والاستعادة
- حفظ كل بيانات البرنامج في ملف واحد أو استعادتها.
        `;
        alert(helpText.trim());
    });

document.getElementById('shutdown-server-btn').addEventListener('click', () => {
        if (confirm("هل أنت متأكد من أنك تريد إيقاف الخادم؟ سيتم إغلاق البرنامج بالكامل.")) {
            fetch('/shutdown', { method: 'POST' })
                .then(() => {
                    alert("تم إرسال أمر الإيقاف. سيتم إغلاق البرنامج الآن.");
                    window.close();
                }).catch(error => {
                    console.error('Could not send shutdown signal:', error);
                    alert('فشل إرسال إشارة الإيقاف إلى الخادم.');
                });
        }
    });
}

function setupBalancingStrategyListener() {
    const advancedSwapLabel = document.getElementById('swap-attempts-label');
    const polishingSwapLabel = document.getElementById('polishing-swaps-label');
    const annealingParamsLabel = document.getElementById('annealing-params-label');
    const solverTimelimitLabel = document.getElementById('solver-timelimit-label');
    const geneticParamsLabel = document.getElementById('genetic-params-label'); // <<< إضافة جديدة
    const tabuParamsLabel = document.getElementById('tabu-params-label');
    const lnsParamsLabel = document.getElementById('lns-params-label');
    const vnsParamsLabel = document.getElementById('vns-params-label');
    const radioButtons = document.querySelectorAll('input[name="balancing_strategy"]');

    function toggleInputs(strategy) {
        // إخفاء كل الخانات الإضافية أولاً
        advancedSwapLabel.style.display = 'none';
        polishingSwapLabel.style.display = 'none';
        annealingParamsLabel.style.display = 'none';
        solverTimelimitLabel.style.display = 'none';
        geneticParamsLabel.style.display = 'none'; // <<< إضافة جديدة
        tabuParamsLabel.style.display = 'none';
        lnsParamsLabel.style.display = 'none';
        vnsParamsLabel.style.display = 'none';

        // إظهار الخانة المناسبة بناءً على الاختيار
        if (strategy === 'advanced') {
            advancedSwapLabel.style.display = 'block';
        } else if (strategy === 'phased_polished') {
            polishingSwapLabel.style.display = 'block';
        } else if (strategy === 'annealing') {
            annealingParamsLabel.style.display = 'block';
        } else if (strategy === 'constraint_solver') {
            solverTimelimitLabel.style.display = 'block';
        } else if (strategy === 'genetic') { // <<< إضافة جديدة
            geneticParamsLabel.style.display = 'block';
        } else if (strategy === 'tabu_search') {
            tabuParamsLabel.style.display = 'block';
        } else if (strategy === 'lns') {
            lnsParamsLabel.style.display = 'block';
        } else if (strategy === 'vns') {
            vnsParamsLabel.style.display = 'block';
        }
    }

    radioButtons.forEach(radio => {
        radio.addEventListener('change', (event) => {
            toggleInputs(event.target.value);
        });
    });

    // تعيين الحالة الأولية عند تحميل الصفحة
    const currentStrategy = document.querySelector('input[name="balancing_strategy"]:checked');
    if (currentStrategy) {
        toggleInputs(currentStrategy.value);
    }
}


// =================================================================================
// --- Section 1: Data Entry Form Listeners ---
// =================================================================================
function setupFormListeners() {
    // --- نموذج إضافة الأساتذة مع التحقق ---
    document.getElementById('add-professors-form').addEventListener('submit', e => {
        e.preventDefault();
        const names = document.getElementById('professor-names').value.split('\n').map(n => n.trim()).filter(Boolean);
        if (names.length === 0) return alert('الرجاء إدخال اسم أستاذ واحد على الأقل.');

        const existingProfNames = new Set(availableProfessors.map(p => p.name));
        for (const name of names) {
            if (existingProfNames.has(name)) {
                return alert(`خطأ: الأستاذ '${name}' موجود بالفعل.`);
            }
        }

        handleFormSubmit(e.target, '/api/professors/bulk', { names });
    });
    
    // --- نموذج إضافة القاعات مع التحقق ---
    document.getElementById('add-halls-form').addEventListener('submit', e => {
        e.preventDefault();
        const type = document.getElementById('hall-type-bulk').value;
        const names = document.getElementById('hall-names').value.split('\n').map(n => n.trim()).filter(Boolean);
        if (names.length === 0) return alert('الرجاء إدخال اسم قاعة واحدة على الأقل.');

        const existingHallNames = new Set(availableHalls.map(h => h.name));
        for (const name of names) {
            if (existingHallNames.has(name)) {
                return alert(`خطأ: القاعة '${name}' موجودة بالفعل.`);
            }
        }

        const halls = names.map(name => ({ name, type }));
        handleFormSubmit(e.target, '/api/halls/bulk', { halls });
    });

    // --- نموذج إضافة المستويات مع التحقق ---
    document.getElementById('add-levels-form').addEventListener('submit', e => {
        e.preventDefault();
        const names = document.getElementById('level-names').value.split('\n').map(n => n.trim()).filter(Boolean);
        if (names.length === 0) return alert('الرجاء إدخال اسم مستوى واحد على الأقل.');
        
        const existingLevelNames = new Set(availableLevels);
        for (const name of names) {
            if (existingLevelNames.has(name)) {
                return alert(`خطأ: المستوى '${name}' موجود بالفعل.`);
            }
        }
        
        handleFormSubmit(e.target, '/api/levels/bulk', { names });
    });
    
    // --- نموذج إضافة المواد مع التحقق ---
    document.getElementById('add-subjects-form').addEventListener('submit', e => {
        e.preventDefault();
        const level = document.getElementById('subject-level-bulk').value;
        const names = document.getElementById('subject-names-bulk').value.split('\n').map(n => n.trim()).filter(Boolean);
        if (!level || names.length === 0) return alert('الرجاء اختيار مستوى وإدخال مادة واحدة على الأقل.');

        const existingSubjects = new Set(availableSubjects.map(s => `${s.name}_${s.level}`));
        for (const name of names) {
            const potentialId = `${name}_${level}`;
            if (existingSubjects.has(potentialId)) {
                return alert(`خطأ: المادة '${name}' موجودة بالفعل في المستوى '${level}'.`);
            }
        }

        handleFormSubmit(e.target, '/api/subjects/bulk', { level, subjects: names });
    });

    document.getElementById('professor-search-assign').addEventListener('input', e => filterList('professors-list-assign', e.target.value));
    document.getElementById('subject-search-assign').addEventListener('input', e => filterList('subjects-list-assign', e.target.value));
}

function handleFormSubmit(formElement, url, body) {
    fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
    })
    .then(handleResponse)
    .then(() => {
        formElement.reset();
        loadInitialDataForSettings();
    }).catch(handleError);
}


// =================================================================================
// --- Section 2: Settings UI and Logic ---
// =================================================================================

function loadInitialDataForSettings() {
    Promise.all([
        fetch('/api/professors').then(res => res.json()),
        fetch('/api/subjects').then(res => res.json()),
        fetch('/api/levels').then(res => res.json()),
        fetch('/api/halls').then(res => res.json()),
        fetch('/api/assignments').then(res => res.json())
    ])
    .then(([professors, subjects, levels, halls, assignments]) => {
        availableProfessors = professors;
        availableSubjects = subjects;
        availableLevels = levels.sort();
        availableHalls = halls; 
        
        populateManagementLists(professors, halls, levels, subjects);
        populateLevelDropdowns();
        populateAssignmentLists(professors, subjects, assignments);
        populateDutyPatternTable(professors);
        populateUnavailabilitySettings(professors);
        populateLevelHallAssignments();
        
        loadAndApplySettings();
    }).catch(error => console.error('فشل في جلب البيانات الأولية للإعدادات:', error));
}

function populateManagementLists(professors, halls, levels, subjects) {
    const lists = {
        'manage-professors-list': { data: professors, type: 'professors', nameKey: 'name' },
        'manage-halls-list': { data: halls, type: 'halls', nameKey: 'name' },
        'manage-levels-list': { data: levels, type: 'levels' },
        'manage-subjects-list': { data: subjects, type: 'subjects', nameKey: 'name' } // << السطر بعد التصحيح
    };

    for (const [listId, config] of Object.entries(lists)) {
        const ul = document.getElementById(listId);
        ul.innerHTML = '';
        config.data.forEach(item => {
            const li = document.createElement('li');
            const name = config.nameKey ? item[config.nameKey] : item;
            let text = name;
            let originalType = '';
            if (config.type === 'halls') {
                text += ` (${item.type})`;
                originalType = item.type;
            }
            if (config.type === 'subjects') {
                text = `${item.name} (${item.level})`;
            }
            
            li.innerHTML = `
                <span>${text}</span>
                <div class="item-actions">
                    <button class="edit-btn" title="تعديل">📝</button>
                    <button class="delete-btn" title="حذف">&times;</button>
                </div>`;
            
            const deleteBtn = li.querySelector('.delete-btn');
            const editBtn = li.querySelector('.edit-btn');

            deleteBtn.addEventListener('click', () => handleDeleteClick({
                type: config.type,
                name: name,
                level: config.type === 'subjects' ? item.level : undefined
            }));

            editBtn.addEventListener('click', () => handleEditClick({
                type: config.type,
                name: name,
                level: config.type === 'subjects' ? item.level : undefined,
                hallType: originalType
            }));

            ul.appendChild(li);
        });
    }
}

function handleDeleteClick(item) {
    if (confirm(`هل أنت متأكد من أنك تريد حذف: "${item.name}"؟`)) {
        let body = { name: item.name };
        if (item.type === 'subjects') {
            body.level = item.level;
        }
        
        fetch(`/api/${item.type}`, {
            method: 'DELETE',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        })
        .then(handleResponse)
        .then(() => loadInitialDataForSettings())
        .catch(handleError);
    }
}

function handleEditClick(item) {
    const { type, name, level, hallType } = item;
    let newValue = prompt(`تعديل "${name}":`, name);

    if (newValue && newValue.trim() !== '' && newValue.trim() !== name) {
        const body = { old_name: name, new_name: newValue.trim() };
        let endpoint = `/api/${type}/edit`; // Endpoint for professor, level, subject

        if (type === 'subjects') {
            const newLevel = prompt(`تعديل مستوى المادة "${newValue.trim()}":\n(إذا غيرت المستوى، سيتم تحديثه في كل مكان آخر)`, level);
            if (!newLevel || newLevel.trim() === '') return alert('المستوى لا يمكن أن يكون فارغاً.');
            body.old_level = level;
            body.new_level = newLevel.trim();
        }

        if (type === 'halls') {
            const newHallType = prompt(`تعديل نوع القاعة (صغيرة, متوسطة, كبيرة):`, hallType);
            if (!newHallType || !['صغيرة', 'متوسطة', 'كبيرة'].includes(newHallType.trim())) {
                return alert('الرجاء إدخال نوع قاعة صالح: صغيرة, متوسطة, كبيرة.');
            }
            body.new_type = newHallType.trim();
            // The endpoint for halls is also /api/halls/edit
        }

        // For levels, the name is changed directly. The endpoint is /api/levels/edit
        if(type === 'levels'){
             // The body is already correct {old_name, new_name}
        }


        fetch(endpoint, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        })
        .then(handleResponse)
        .then(() => loadInitialDataForSettings())
        .catch(handleError);
    }
}


function populateLevelDropdowns() {
    const bulkLevelSelect = document.getElementById('subject-level-bulk');
    bulkLevelSelect.innerHTML = '<option value="" disabled selected>اختر المستوى الدراسي</option>';
    availableLevels.forEach(level => {
        const option = document.createElement('option');
        option.value = level;
        option.textContent = level;
        bulkLevelSelect.appendChild(option);
    });
}

function populateAssignmentLists(professors, subjects, assignments) {
    const profList = document.getElementById('professors-list-assign');
    const subjList = document.getElementById('subjects-list-assign');
    profList.innerHTML = '';
    subjList.innerHTML = '';

    const assignedUniqueSubjects = new Set();
    Object.values(assignments).forEach(subjArray => {
        subjArray.forEach(subj => assignedUniqueSubjects.add(subj));
    });

    professors.forEach(prof => {
        const li = document.createElement('li');
        li.dataset.professorName = prof.name;
        const assignedCourses = assignments[prof.name] || [];
        const hasAssignments = assignedCourses.length > 0;
        
        li.innerHTML = `
            <div class="list-entry">
                <span class="item-name">${prof.name} ${hasAssignments ? `<span class="item-count">(${assignedCourses.length})</span>` : ''}</span>
                ${hasAssignments ? '<button class="courses-dropdown-btn">▼</button>' : ''}
            </div>
            ${hasAssignments ? `<ul class="courses-dropdown-list">${assignedCourses.map(c => `<li>${c}</li>`).join('')}</ul>` : ''}
        `;
        if (hasAssignments) li.classList.add('assigned-prof');
        profList.appendChild(li);
    });

    subjects.forEach(subj => {
        const uniqueId = `${subj.name} (${subj.level})`;
        const li = document.createElement('li');
        li.textContent = uniqueId;
        li.dataset.uniqueId = uniqueId;

        if (assignedUniqueSubjects.has(uniqueId)) {
            li.classList.add('assigned-subj');
            li.addEventListener('click', handleUnassignClick);
        } else {
            li.addEventListener('click', handleSubjectSelect);
        }
        subjList.appendChild(li);
    });
    
    profList.addEventListener('click', handleProfessorInteraction);
    document.getElementById('assign-subjects-button').onclick = assignSubjectsToProfessor;
}

function assignSubjectsToProfessor() {
    if (!selectedProfessorForAssign || selectedSubjectsForAssign.length === 0) return;
    const uniqueIds = selectedSubjectsForAssign.map(s => s.uniqueId);

    fetch('/api/assign-subjects/bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            professor: selectedProfessorForAssign,
            subjects: uniqueIds,
        }),
    })
    .then(handleResponse)
    .then(() => {
        selectedProfessorForAssign = null;
        selectedSubjectsForAssign = [];
        loadInitialDataForSettings();
    })
    .catch(handleError);
}

function handleUnassignClick(event) {
    if (confirm(`هل أنت متأكد من أنك تريد إلغاء إسناد المادة: "${event.currentTarget.dataset.uniqueId}"؟`)) {
        fetch('/api/unassign-subject', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ subject: event.currentTarget.dataset.uniqueId }),
        })
        .then(handleResponse)
        .then(() => loadInitialDataForSettings())
        .catch(handleError);
    }
}


function populateLevelHallAssignments() {
    const container = document.getElementById('level-halls-assignment-container');
    container.innerHTML = '';
    const table = document.createElement('table');
    table.className = 'settings-table';
    table.innerHTML = `<thead><tr><th>المستوى الدراسي</th><th>القاعات المخصصة للامتحان</th></tr></thead><tbody></tbody>`;
    const tbody = table.querySelector('tbody');

    availableLevels.forEach(level => {
        const row = document.createElement('tr');
        row.dataset.levelName = level;
        const levelCell = document.createElement('td');
        levelCell.textContent = level;
        const hallsCell = document.createElement('td');
        const checkboxGroup = document.createElement('div');
        checkboxGroup.className = 'halls-checkbox-group';
        availableHalls.forEach(hall => {
            const label = document.createElement('label');
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = hall.name;
            label.appendChild(checkbox);
            label.append(` ${hall.name} (${hall.type})`);
            checkboxGroup.appendChild(label);
        });
        hallsCell.appendChild(checkboxGroup);
        row.appendChild(levelCell);
        row.appendChild(hallsCell);
        tbody.appendChild(row);
    });
    container.appendChild(table);
}

function populateDutyPatternTable(professors) {
    const container = document.getElementById('professors-duty-pattern-container');
    container.innerHTML = ''; 
    const table = document.createElement('table');
    table.className = 'settings-table';
    table.innerHTML = `<thead><tr><th>الأستاذ</th><th>نمط أيام الحراسة</th></tr></thead><tbody></tbody>`;
    const tbody = table.querySelector('tbody');
    professors.forEach(prof => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${prof.name}</td>
            <td>
                <label>
                    <input type="radio" name="pattern-${prof.name}" value="one_day_only">
                    يوم واحد فقط
                </label>
                <label>
                    <input type="radio" name="pattern-${prof.name}" value="flexible_2_days" checked>
                    مرن (يومان)
                </label>
                <label>
                    <input type="radio" name="pattern-${prof.name}" value="consecutive_strict">
                    يومان متتاليان (إلزامي)
                </label>
                <label>
                    <input type="radio" name="pattern-${prof.name}" value="flexible_3_days">
                    مرن (يومان أو 3 أيام)
                </label>
            </td>
        `;
        tbody.appendChild(row);
    });
    container.appendChild(table);
}

function populateUnavailabilitySettings(professors) {
    const container = document.getElementById('professors-unavailability-container');
    container.innerHTML = ''; 
    const table = document.createElement('table');
    table.className = 'settings-table';
    table.innerHTML = `<thead><tr><th>الأستاذ</th><th>الأيام غير المتاحة</th></tr></thead><tbody></tbody>`;
    const tbody = table.querySelector('tbody');

    professors.forEach(prof => {
        const row = document.createElement('tr');
        row.dataset.profName = prof.name;
        row.innerHTML = `
            <td>${prof.name}</td>
            <td class="unavailable-days-cell">
                <span>الرجاء تحديد أيام الامتحانات في المرحلة 5 أولاً</span>
            </td>
        `;
        tbody.appendChild(row);
    });
    container.appendChild(table);
}

function updateUnavailabilityDateOptions() {
    const dates = new Set();
    document.querySelectorAll('#exam-days-container input[type="date"]').forEach(dateInput => {
        if (dateInput.value) {
            dates.add(dateInput.value);
        }
    });

    const sortedDates = [...dates].sort();

    document.querySelectorAll('.unavailable-days-cell').forEach(cell => {
        const profName = cell.closest('tr').dataset.profName;
        const previouslySelected = [];
        cell.querySelectorAll('input:checked').forEach(cb => previouslySelected.push(cb.value));
        
        cell.innerHTML = ''; 

        if (sortedDates.length === 0) {
            cell.innerHTML = '<span>الرجاء تحديد أيام الامتحانات في المرحلة 5 أولاً</span>';
        } else {
            sortedDates.forEach(date => {
                const isChecked = previouslySelected.includes(date);
                const label = document.createElement('label');
                label.className = 'unavailable-day-label';
                
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.value = date;
                checkbox.name = `unavailable-${profName}`;
                checkbox.checked = isChecked;
                
                checkbox.addEventListener('change', (e) => {
                    const parentLabel = e.target.closest('.unavailable-day-label');
                    if(e.target.checked) {
                        parentLabel.classList.add('selected');
                    } else {
                        parentLabel.classList.remove('selected');
                    }
                });
                
                if(isChecked) {
                    label.classList.add('selected');
                }

                const span = document.createElement('span');
                span.textContent = date;

                label.appendChild(checkbox);
                label.appendChild(span);
                cell.appendChild(label);
            });
        }
    });
}

function setupExamScheduleBuilder() { 
    document.getElementById('add-exam-day-button').addEventListener('click', () => {
        addExamDayUI();
        updateUnavailabilityDateOptions(); 
    });
    
    document.getElementById('exam-days-container').addEventListener('change', (e) => {
        if (e.target && e.target.matches('input[type="date"]')) {
            updateUnavailabilityDateOptions();
        }
    });
}

function addExamDayUI() {
    examDayCounter++;
    const container = document.getElementById('exam-days-container');
    const dayDiv = document.createElement('div');
    dayDiv.className = 'exam-day';
    
    dayDiv.innerHTML = `
        <div class="exam-day-header">
            <h4>يوم الامتحان رقم ${examDayCounter}</h4>
            <input type="date" required>
            <button class="duplicate-day-btn" title="تكرار هذا اليوم مع فتراته">🔄</button>
            <button class="remove-day-btn" title="حذف هذا اليوم">&times;</button>
        </div>
        <div class="time-slots-container"></div>
        <button class="add-timeslot-button action-button">+ إضافة فترة زمنية</button>`;
    
    dayDiv.querySelector('.add-timeslot-button').addEventListener('click', e => addTimeSlotUI(e.target.previousElementSibling));
    dayDiv.querySelector('.duplicate-day-btn').addEventListener('click', e => {
        duplicateDay(e.currentTarget.closest('.exam-day'));
        updateUnavailabilityDateOptions();
    });
    dayDiv.querySelector('.remove-day-btn').addEventListener('click', e => {
        e.currentTarget.closest('.exam-day').remove();
        updateUnavailabilityDateOptions();
    });
    container.appendChild(dayDiv);
    return dayDiv;
}

// +++ START: Updated Function with a remove button +++
function addTimeSlotUI(container) {
    const slotDiv = document.createElement('div');
    slotDiv.className = 'time-slot';

    const levelsContainer = document.createElement('div');
    levelsContainer.className = 'time-slot-levels levels-checkbox-group';
    
    availableLevels.forEach(level => {
        const label = document.createElement('label');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = level;
        
        label.appendChild(checkbox);
        label.append(` ${level}`);
        levelsContainer.appendChild(label);
    });
    
    const slotTypeSelect = `
        <select class="slot-type-select">
            <option value="primary" selected>فترة أساسية</option>
            <option value="reserve">فترة احتياطية</option>
        </select>
    `;
    
    const inputsAndButtonDiv = document.createElement('div');
    inputsAndButtonDiv.className = 'time-slot-inputs-container';
    inputsAndButtonDiv.innerHTML = `
        <div class="time-slot-inputs">
            <input type="time" required value="09:00">
            <input type="time" required value="10:30">
            ${slotTypeSelect}
        </div>
        <button class="remove-timeslot-btn" title="حذف الفترة">&times;</button>
    `;

    slotDiv.appendChild(inputsAndButtonDiv);
    slotDiv.appendChild(levelsContainer);
    
    // Add event listener to the new remove button
    slotDiv.querySelector('.remove-timeslot-btn').addEventListener('click', (e) => {
        e.currentTarget.closest('.time-slot').remove();
    });

    container.appendChild(slotDiv);
    return slotDiv;
}
// +++ END: Updated Function with a remove button +++

function duplicateDay(sourceDayDiv) {
    const newDayDiv = addExamDayUI();
    
    const sourceTimeSlots = sourceDayDiv.querySelectorAll('.time-slot');
    const newTimeSlotsContainer = newDayDiv.querySelector('.time-slots-container');
    
    sourceTimeSlots.forEach(sourceSlot => {
        const newSlotDiv = addTimeSlotUI(newTimeSlotsContainer);
        
        const sourceStartTime = sourceSlot.querySelector('input[type="time"]:nth-of-type(1)').value;
        const sourceEndTime = sourceSlot.querySelector('input[type="time"]:nth-of-type(2)').value;
        const sourceType = sourceSlot.querySelector('.slot-type-select').value;
        const sourceSelectedLevels = Array.from(sourceSlot.querySelectorAll('.levels-checkbox-group input:checked')).map(cb => cb.value);
        
        newSlotDiv.querySelector('input[type="time"]:nth-of-type(1)').value = sourceStartTime;
        newSlotDiv.querySelector('input[type="time"]:nth-of-type(2)').value = sourceEndTime;
        newSlotDiv.querySelector('.slot-type-select').value = sourceType;
        
        Array.from(newSlotDiv.querySelectorAll('.levels-checkbox-group input')).forEach(cb => {
            if (sourceSelectedLevels.includes(cb.value)) {
                cb.checked = true;
            }
        });
    });
}

function setupGenerationListener() { 
    document.getElementById('generate-schedule-button').addEventListener('click', collectAndSendData);
    document.getElementById('save-settings-button').addEventListener('click', manualSaveSettings);
}

// استبدل الدالة القديمة بالكامل بهذه الدالة
function collectAllData() {
    const dutyPatterns = {};
    document.querySelectorAll('#professors-duty-pattern-container tbody tr').forEach(row => {
        const profName = row.cells[0].textContent.trim();
        const pattern = row.querySelector('input[type="radio"]:checked').value;
        dutyPatterns[profName] = pattern;
    });

    const levelHallAssignments = {};
    document.querySelectorAll('#level-halls-assignment-container tbody tr').forEach(row => {
        const levelName = row.dataset.levelName;
        const assignedHalls = [];
        row.querySelectorAll('input[type="checkbox"]:checked').forEach(checkbox => {
            assignedHalls.push(checkbox.value);
        });
        if (assignedHalls.length > 0) {
            levelHallAssignments[levelName] = assignedHalls;
        }
    });

    const examSchedule = {};
    document.querySelectorAll('.exam-day').forEach((dayDiv, index) => {
        const date = dayDiv.querySelector('input[type="date"]').value;
        if (!date) return;
        examSchedule[date] = [];
        dayDiv.querySelectorAll('.time-slot').forEach(slotDiv => {
            const times = slotDiv.querySelectorAll('input[type="time"]');
            const type = slotDiv.querySelector('.slot-type-select').value;
            const selectedLevels = [];
            slotDiv.querySelectorAll('.levels-checkbox-group input:checked').forEach(cb => {
                selectedLevels.push(cb.value);
            });

            if (times[0].value && times[1].value && selectedLevels.length > 0) {
                examSchedule[date].push({ 
                    time: `${times[0].value}-${times[1].value}`, 
                    levels: selectedLevels,
                    type: type 
                });
            }
        });
    });

    const unavailableDays = {};
    document.querySelectorAll('#professors-unavailability-container tbody tr').forEach(row => {
        const profName = row.dataset.profName;
        const selectedDates = [];
        row.querySelectorAll('input[type="checkbox"]:checked').forEach(cb => {
            selectedDates.push(cb.value);
        });
        if (selectedDates.length > 0) {
            unavailableDays[profName] = selectedDates;
        }
    });

    const assignOwnerAsGuard = document.getElementById('assign-owner-as-guard-checkbox').checked;
    const maxShifts = document.querySelector('input[name="max-shifts-limit"]:checked').value;
    const maxLargeHallShifts = document.querySelector('input[name="max-large-hall-shifts"]:checked').value;
    const intensiveSearch = document.getElementById('intensive-search-checkbox').checked;
    const groupSubjects = document.getElementById('group-subjects-checkbox').checked;
    const iterations = document.getElementById('iterations-count').value;
    const largeHallWeight = document.getElementById('large-hall-weight').value;
    const otherHallWeight = document.getElementById('other-hall-weight').value;
    const guardsLargeHall = document.getElementById('guards-large-hall').value;
    const guardsMediumHall = document.getElementById('guards-medium-hall').value;
    const guardsSmallHall = document.getElementById('guards-small-hall').value;
    const lastDayRestriction = document.querySelector('input[name="last_day_restriction"]:checked').value; 
    const balancingStrategy = document.querySelector('input[name="balancing_strategy"]:checked').value;
    const swapAttempts = document.getElementById('swap-attempts-count').value;
    const polishingSwaps = document.getElementById('polishing-swaps-count').value;
    const enableCustomTargets = document.getElementById('enable-custom-targets-checkbox').checked;
    const annealingTemp = document.getElementById('annealing-temp').value;
    const annealingCooling = document.getElementById('annealing-cooling').value;
    const annealingIterations = document.getElementById('annealing-iterations').value;
    const solverTimelimit = document.getElementById('solver-timelimit').value;

    // --- الجزء الذي تم تصحيحه ---
    const geneticPopulation = document.getElementById('genetic-population').value;
    const geneticGenerations = document.getElementById('genetic-generations').value;
    const geneticElitism = document.getElementById('genetic-elitism').value; // <-- السطر الجديد
    const geneticMutation = document.getElementById('genetic-mutation').value;
    const tabuIterations = document.getElementById('tabu-iterations').value;
    const tabuTenure = document.getElementById('tabu-tenure').value;
    const tabuNeighborhoodSize = document.getElementById('tabu-neighborhood-size').value;
    const professorPartnerships = currentProfessorPartnerships;
    const lnsIterations = document.getElementById('lns-iterations').value;
    const lnsDestroyFraction = document.getElementById('lns-destroy-fraction').value;
    const vnsIterations = document.getElementById('vns-iterations').value;
    const vnsMaxK = document.getElementById('vns-max-k').value;

    return { 
        dutyPatterns, levelHallAssignments, examSchedule, unavailableDays,
        assignOwnerAsGuard, maxShifts, maxLargeHallShifts, intensiveSearch, groupSubjects, iterations,
        lastDayRestriction,
        largeHallWeight, otherHallWeight, guardsLargeHall,
        guardsMediumHall, guardsSmallHall, enableCustomTargets, customTargetPatterns,
        balancingStrategy,
        swapAttempts,
        polishingSwaps,
        annealingTemp,
        annealingCooling,
        annealingIterations,
        solverTimelimit,
        geneticPopulation,
        geneticGenerations,
        geneticElitism, // <-- إضافة المتغير الجديد هنا
        geneticMutation,
        tabuIterations,
        tabuTenure,
        tabuNeighborhoodSize,
        professorPartnerships,
        lnsIterations,
        lnsDestroyFraction,
        vnsIterations,
        vnsMaxK
    };
}

async function manualSaveSettings() {
    const settings = collectAllData();
    try {
        const response = await fetch('/api/settings', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(settings),
        });
        const data = await handleResponse(response);
        // <<< تعديل: استخدام الدالة الجديدة بدل alert >>>
        showNotification(data.message || 'تم حفظ الإعدادات بنجاح.');
    } catch (error) {
        handleError(error);
    }
}

async function autoSaveSettings() {
    const settings = collectAllData();
    try {
        await fetch('/api/settings', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(settings),
        });
    } catch (error) {
        console.error("فشل الحفظ التلقائي:", error);
    }
}

// في ملف script.js، استبدل هذه الدالة بالكامل

// في ملف script.js، استبدل هذه الدالة بالكامل
async function collectAndSendData() {
    await autoSaveSettings();
    const resultsContainer = document.getElementById('results-container');
    const generateBtn = document.getElementById('generate-schedule-button');
    const logContainer = document.getElementById('live-log-container');
    const logOutput = document.getElementById('live-log-output');

    // --- إضافة جديدة: جلب عناصر شريط التقدم ---
    const progressBarContainer = document.getElementById('progress-bar-container');
    const progressBarInner = document.getElementById('progress-bar-inner');

    if (eventSource && eventSource.readyState !== EventSource.CLOSED) {
        eventSource.close();
    }

    logContainer.classList.remove('hidden');
    // --- إضافة جديدة: إظهار شريط التقدم وإعادة تعيينه ---
    progressBarContainer.classList.remove('hidden');
    progressBarInner.style.width = '0%';
    progressBarInner.textContent = '0%';

    logOutput.textContent = 'بدء الاتصال بالخادم...\n';
    resultsContainer.style.display = 'block';
    resultsContainer.innerHTML = '<h3>جاري إنشاء الجدول، يرجى الانتظار...</h3>';
    generateBtn.disabled = true;
    generateBtn.textContent = '... جاري البحث عن حل';

    const allSettings = collectAllData();
    if (Object.keys(allSettings.examSchedule).length === 0) {
        alert('الرجاء إعداد جدول الامتحانات أولاً.');
        resultsContainer.innerHTML = '';
        generateBtn.disabled = false;
        generateBtn.textContent = '🚀 إنشاء جدول الحراسة الآن';
        logContainer.classList.add('hidden');
        // --- إضافة جديدة: إخفاء شريط التقدم عند الخطأ ---
        progressBarContainer.classList.add('hidden');
        return;
    }

    eventSource = new EventSource('/stream-logs');
    
    eventSource.onmessage = function(event) {
        // --- تعديل مهم: التحقق من نوع الرسالة ---
        if (event.data.startsWith("PROGRESS:")) {
            const progress = event.data.split(':')[1];
            progressBarInner.style.width = progress + '%';
            progressBarInner.textContent = progress + '%';
            return; // لا تطبع رسالة التقدم في الصندوق الأسود
        }
        
        if (event.data.startsWith("DONE")) {
            eventSource.close();
            // --- إضافة جديدة: إخفاء شريط التقدم عند الانتهاء ---
            progressBarContainer.classList.add('hidden');

            const jsonString = event.data.substring(4); 
            
            if (jsonString) {
                try {
                    const finalData = JSON.parse(jsonString);
                    if (finalData.success) {
                        lastGeneratedSchedule = finalData.schedule; 
                        displayResults(finalData.schedule);
                        displayReports(finalData);
                        if (finalData.chart_data) displayWorkloadChart(finalData.chart_data);
                        if (finalData.balance_report) displayBalanceReport(finalData.balance_report);
                        if (finalData.stats_dashboard) displayStatsDashboard(finalData.stats_dashboard);
                    } else {
                        resultsContainer.innerHTML = `<p class="failure-message">فشل إنشاء الجدول: ${finalData.message}</p>`;
                    }
                } catch (e) {
                    console.error("خطأ في تحليل JSON من الخادم:", e);
                    resultsContainer.innerHTML = `<p class="failure-message">فشل في تحليل الاستجابة النهائية من الخادم.</p>`;
                }
            } else {
                 resultsContainer.innerHTML = `<p class="failure-message">انتهت العملية ولكن لم يتم استلام بيانات.</p>`;
            }

            generateBtn.disabled = false;
            generateBtn.textContent = '🚀 إنشاء جدول الحراسة الآن';
            return;
        }
        
        logOutput.textContent += event.data + '\n';
        logOutput.scrollTop = logOutput.scrollHeight;
    };

    eventSource.onerror = function(err) {
        logOutput.textContent += 'انقطع الاتصال بالخادم.\n';
        eventSource.close();
        generateBtn.disabled = false;
        generateBtn.textContent = '🚀 إنشاء جدول الحراسة الآن';
        // --- إضافة جديدة: إخفاء شريط التقدم عند الخطأ ---
        progressBarContainer.classList.add('hidden');
    };

    fetch('/api/generate-guard-schedule', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(allSettings),
    })
    .then(response => {
        if (!response.ok) {
           throw new Error('فشل الخادم في بدء عملية إنشاء الجدول.');
        }
        console.log("تم إرسال طلب إنشاء الجدول بنجاح.");
    })
    .catch(error => {
        resultsContainer.innerHTML = `<p class="failure-message">حدث خطأ في الاتصال الأولي بالخادم. الرجاء التأكد من أن الخادم يعمل.</p>`;
        handleError(error);
        if (eventSource) eventSource.close();
        generateBtn.disabled = false;
        generateBtn.textContent = '🚀 إنشاء جدول الحراسة الآن';
        // --- إضافة جديدة: إخفاء شريط التقدم عند الخطأ ---
        progressBarContainer.classList.add('hidden');
    });
}

function displayReports(data) {
    const reportsContainer = document.getElementById('reports-display-area');
    if (!reportsContainer) return;
    reportsContainer.innerHTML = ''; 

    const hasSchedulingReport = data.scheduling_report && data.scheduling_report.length > 0;
    const hasFailures = data.failures && data.failures.length > 0;
    const hasProfReport = data.prof_report && data.prof_report.length > 0;

    if (!hasSchedulingReport && !hasFailures && !hasProfReport) {
        return;
    }
    
    if (hasSchedulingReport) {
        const section = document.createElement('div');
        section.className = 'report-section';
        const title = document.createElement('h4');
        title.textContent = 'تقرير جدولة المواد وملاحظات الحراسة';
        const list = document.createElement('ul');
        data.scheduling_report.forEach(item => {
            const li = document.createElement('li');
            li.textContent = `• ${item.subject}${item.level ? ` (${item.level})` : ''} -> ${item.reason}`;
            list.appendChild(li);
        });
        section.appendChild(title);
        section.appendChild(list);
        reportsContainer.appendChild(section);
    }

    if (hasFailures) {
        const section = document.createElement('div');
        section.className = 'report-section';
        const title = document.createElement('h4');
        title.textContent = 'تقرير أخطاء قيود الأساتذة';
        const list = document.createElement('ul');
        data.failures.forEach(fail => {
            const li = document.createElement('li');
            li.className = 'failure';
            li.textContent = `• الأستاذ: ${fail.name} (${fail.reason})`;
            list.appendChild(li);
        });
        section.appendChild(title);
        section.appendChild(list);
        reportsContainer.appendChild(section);
    }
    
    if (hasProfReport) {
        const section = document.createElement('div');
        section.className = 'report-section';
        const title = document.createElement('h4');
        title.textContent = 'تقرير إجمالي حصص الحراسة للأساتذة';
        const list = document.createElement('ul');
        data.prof_report.forEach(line => {
            const li = document.createElement('li');
            li.textContent = line;
            list.appendChild(li);
        });
        section.appendChild(title);
        section.appendChild(list);
        reportsContainer.appendChild(section);
    }
}

function displayResults(schedule) {
    const container = document.getElementById('results-container');
    container.innerHTML = `
    <div class="final-buttons-container export-buttons-container">
        <button id="export-schedule-button" class="action-button export-btn">تصدير جداول الامتحانات (Excel)</button>
        <button id="export-prof-button" class="action-button export-btn">تصدير جداول الأساتذة (Excel)</button>
        <button id="export-schedule-word-button" class="action-button export-btn" style="background-color: #2a5599;">تصدير جداول الامتحانات (Word)</button>
        <button id="export-prof-word-button" class="action-button export-btn" style="background-color: #2a5599;">تصدير جداول الأساتذة (Word)</button>
        <button id="export-prof-anonymous-word-button" class="action-button export-btn" style="background-color: #5a9955;">تصدير جداول الأساتذة (مُبسَّط)</button>
    </div>

    <div id="balance-report-area"></div>
    <div id="stats-dashboard-container" class="stats-dashboard-container" style="display: none;">
        <h3>لوحة المعلومات الإحصائية</h3>
        <div id="stats-dashboard" class="stats-dashboard"></div>
    </div>
    <div id="chart-container" style="width: 90%; margin: 40px auto; display: none;">
         <h3 style="text-align: center;">رسم بياني لتوزيع عبء الحراسة</h3>
         <canvas id="workload-chart"></canvas>
    </div>
    
    <div id="results-search-container">
        <input type="text" id="results-search-input" placeholder="ابحث عن اسم أستاذ أو مادة لتظليلها في الجداول...">
    </div>

    <div id="schedule-tables-container"></div>
    <div id="reports-display-area"></div>`;

    document.getElementById('export-schedule-button').addEventListener('click', exportSchedule);
    document.getElementById('export-prof-button').addEventListener('click', exportProfSchedule);
    document.getElementById('export-schedule-word-button').addEventListener('click', exportScheduleWord);
    document.getElementById('export-prof-word-button').addEventListener('click', exportProfScheduleWord);
    document.getElementById('export-prof-anonymous-word-button').addEventListener('click', exportProfScheduleAnonymous);
    
    setupResultsSearch();

    const tablesContainer = document.getElementById('schedule-tables-container');
    tablesContainer.innerHTML = '';

    try {
        let allExams = [];
        const allDates = Object.keys(schedule).sort();
        const allLevels = new Set();
        const allTimes = new Set();

        allDates.forEach(date => {
            Object.keys(schedule[date]).sort().forEach(time => {
                allTimes.add(time);
                schedule[date][time].forEach(exam => {
                    allExams.push({ ...exam, date, time });
                    allLevels.add(exam.level);
                });
            });
        });

        const sortedLevels = [...allLevels].sort();
        const sortedTimes = [...allTimes].sort();
        const dayNames = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"];

        sortedLevels.forEach(level => {
            const levelExams = allExams.filter(exam => exam.level === level);
            if (levelExams.length === 0) return;

            const levelContainer = document.createElement('div');
            levelContainer.className = 'level-schedule-container';
            
            const title = document.createElement('h4');
            title.className = 'level-schedule-title';
            title.textContent = `جدول امتحانات: ${level}`;
            levelContainer.appendChild(title);

            const table = document.createElement('table');
            table.className = 'results-grid-table';
            
            const thead = table.createTHead();
            const headerRow = thead.insertRow();
            headerRow.innerHTML = '<th>الفترة</th>';
            allDates.forEach(dateStr => {
                const dateObj = new Date(dateStr);
                const utcDate = new Date(dateObj.valueOf() + dateObj.getTimezoneOffset() * 60000);
                const dayName = dayNames[utcDate.getDay()];
                headerRow.innerHTML += `<th>${dayName}<br>${dateStr}</th>`;
            });

            const tbody = table.createTBody();
            sortedTimes.forEach(time => {
                const row = tbody.insertRow();
                row.insertCell().innerHTML = `<strong>${time}</strong>`;

                allDates.forEach(date => {
                    const cell = row.insertCell();
                    const exam = levelExams.find(ex => ex.date === date && ex.time === time);
                    
                    if (exam) {
                        const examCellDiv = document.createElement('div');
                        examCellDiv.className = 'exam-cell';
                        if (exam.guards_incomplete) {
                            examCellDiv.classList.add('guards-incomplete');
                        }
                    
                        let guardsCopy = [...exam.guards];
                        const hallsByType = { كبيرة: [], متوسطة: [], صغيرة: [] };
                        (exam.halls || []).forEach(h => {
                            if(hallsByType[h.type] !== undefined) {
                                hallsByType[h.type].push(h.name);
                            }
                        });

                        let hallHtml = '';
                        
                        const processHalls = (type, title, guardsPerHall) => {
                            if (hallsByType[type].length > 0) {
                                const names = hallsByType[type].join(', ');
                                const count = guardsPerHall * hallsByType[type].length;
                                const hallGuards = guardsCopy.splice(0, count);

                                // --- بداية التعديل ---
                                // تحويل كل اسم حارس إلى عنصر HTML، مع تنسيق خاص لكلمة "نقص"
                                const styledGuards = hallGuards.map(guard => {
                                    if (guard.includes('**نقص**')) {
                                        return `<span class="guard-shortage">${guard}</span>`;
                                    }
                                    return guard;
                                }).join('<br>');
                                // --- نهاية التعديل ---

                                return `<div class="hall-group"><span class="hall-type-title">${title}: ${names}</span><div class="hall-guards-list">${styledGuards}</div></div>`;
                            }
                            return '';
                        };
                        
                        hallHtml += processHalls('كبيرة', 'القاعة الكبيرة', 4);
                        hallHtml += processHalls('متوسطة', 'القاعات المتوسطة', 2);
                        hallHtml += processHalls('صغيرة', 'القاعات الصغيرة', 1);

                        examCellDiv.innerHTML = `
                            <div class="exam-subject">${exam.subject}</div>
                            <div class="exam-professor">أستاذ المادة: ${exam.professor}</div>
                            <div class="exam-guards-section">
                                <strong class="guards-title">الحراسة:</strong>
                                ${hallHtml}
                            </div>
                        `;
                        cell.appendChild(examCellDiv);
                    }
                });
            });
            
            levelContainer.appendChild(table);
            tablesContainer.appendChild(levelContainer);
        });
    } catch (e) {
        console.error("خطأ فادح في دالة displayResults:", e);
        tablesContainer.innerHTML = `<p class="failure-message">فشل عرض النتائج بسبب خطأ في الجافاسكريبت. الرجاء التحقق من نافذة Console لمزيد من التفاصيل.</p>`;
    }
}

async function exportSchedule() {
    if (!lastGeneratedSchedule) {
        alert("يرجى إنشاء جدول أولاً قبل التصدير.");
        return;
    }
    const button = document.getElementById('export-schedule-button');
    button.disabled = true;
    button.textContent = 'جاري التصدير...';

    try {
        const response = await fetch('/api/export-schedule', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(lastGeneratedSchedule)
        });

        if (!response.ok) throw new Error('فشل التصدير من الخادم');

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'الجداول_المجمعة.xlsx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

    } catch (err) {
        alert('حدث خطأ أثناء تصدير الملف.');
        console.error(err);
    } finally {
        button.disabled = false;
        button.textContent = 'تصدير جداول الامتحانات (Excel)';
    }
}

async function exportProfSchedule() {
    if (!lastGeneratedSchedule) {
        alert("يرجى إنشاء جدول أولاً قبل التصدير.");
        return;
    }
    const button = document.getElementById('export-prof-button');
    button.disabled = true;
    button.textContent = 'جاري التصدير...';

    try {
        const response = await fetch('/api/export-prof-schedules', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(lastGeneratedSchedule)
        });

        if (!response.ok) throw new Error('فشل التصدير من الخادم');

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'جداول_الحراسة_للأساتذة.xlsx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

    } catch (err) {
        alert('حدث خطأ أثناء تصدير الملف.');
        console.error(err);
    } finally {
        button.disabled = false;
        button.textContent = 'تصدير جداول الأساتذة (Excel)';
    }
}

async function loadAndApplySettings() {
    try {
        const response = await fetch('/api/settings');
        const settings = await handleResponse(response);
        if (settings && Object.keys(settings).length > 0) {
            populateUIWithSettings(settings);
            console.log('تم تحميل الإعدادات المحفوظة بنجاح.');
        }
    } catch (error) {
        console.log('لم يتم العثور على ملف إعدادات محفوظ، سيبدأ البرنامج بإعدادات فارغة.');
    }
}

function populateUIWithSettings(settings) {
    if (settings.levelHallAssignments) {
        for (const [level, halls] of Object.entries(settings.levelHallAssignments)) {
            const row = document.querySelector(`#level-halls-assignment-container tr[data-level-name="${level}"]`);
            if (row) {
                halls.forEach(hallName => {
                    const checkbox = row.querySelector(`input[value="${hallName}"]`);
                    if (checkbox) checkbox.checked = true;
                });
            }
        }
    }

    if (settings.dutyPatterns) {
        for (const [prof, pattern] of Object.entries(settings.dutyPatterns)) {
            const radio = document.querySelector(`#professors-duty-pattern-container input[name="pattern-${prof}"][value="${pattern}"]`);
            if(radio) radio.checked = true;
        }
    }

    if (settings.examSchedule) {
        document.getElementById('exam-days-container').innerHTML = '';
        examDayCounter = 0;
        const sortedDates = Object.keys(settings.examSchedule).sort();

        for (const date of sortedDates) {
            const daySlots = settings.examSchedule[date];
            const dayDiv = addExamDayUI();
            dayDiv.querySelector('input[type="date"]').value = date;
            const slotsContainer = dayDiv.querySelector('.time-slots-container');

            daySlots.forEach(slotData => {
                const slotDiv = addTimeSlotUI(slotsContainer);
                const [startTime, endTime] = slotData.time.split('-');
                slotDiv.querySelector('input[type="time"]:nth-of-type(1)').value = startTime;
                slotDiv.querySelector('input[type="time"]:nth-of-type(2)').value = endTime;
                slotDiv.querySelector('.slot-type-select').value = slotData.type;

                slotData.levels.forEach(levelName => {
                    const checkbox = slotDiv.querySelector(`input[value="${levelName}"]`);
                    if (checkbox) checkbox.checked = true;
                });
            });
        }
        updateUnavailabilityDateOptions(); 
    }

    if (settings.unavailableDays) {
        for (const [prof, dates] of Object.entries(settings.unavailableDays)) {
            const row = document.querySelector(`#professors-unavailability-container tr[data-prof-name="${prof}"]`);
            if(row) {
                const checkboxes = row.querySelectorAll('input[type="checkbox"]');
                checkboxes.forEach(cb => {
                    if(dates.includes(cb.value)) {
                        cb.checked = true;
                        cb.closest('.unavailable-day-label').classList.add('selected');
                    }
                });
            }
        }
    }

    if (settings.assignOwnerAsGuard !== undefined) {
        document.getElementById('assign-owner-as-guard-checkbox').checked = settings.assignOwnerAsGuard;
    }
    if (settings.maxShifts !== undefined) {
        document.querySelector(`input[name="max-shifts-limit"][value="${settings.maxShifts}"]`).checked = true;
    }
    if (settings.maxLargeHallShifts !== undefined) {
        document.querySelector(`input[name="max-large-hall-shifts"][value="${settings.maxLargeHallShifts}"]`).checked = true;
    }
    if (settings.largeHallWeight !== undefined) document.getElementById('large-hall-weight').value = settings.largeHallWeight;
    if (settings.otherHallWeight !== undefined) document.getElementById('other-hall-weight').value = settings.otherHallWeight;
    if (settings.intensiveSearch !== undefined) document.getElementById('intensive-search-checkbox').checked = settings.intensiveSearch;
    if (settings.iterations !== undefined) document.getElementById('iterations-count').value = settings.iterations;

    // --- هذا هو التعديل الجديد ---
    if (settings.lastDayRestriction) {
        document.querySelector(`input[name="last_day_restriction"][value="${settings.lastDayRestriction}"]`).checked = true;
    }

    if (settings.guardsLargeHall) document.getElementById('guards-large-hall').value = settings.guardsLargeHall;
    if (settings.guardsMediumHall) document.getElementById('guards-medium-hall').value = settings.guardsMediumHall;
    if (settings.guardsSmallHall) document.getElementById('guards-small-hall').value = settings.guardsSmallHall;

    if (settings.balancingStrategy) {
        const radio = document.querySelector(`input[name="balancing_strategy"][value="${settings.balancingStrategy}"]`);
        if (radio) radio.checked = true;
    } else {
        document.querySelector('input[name="balancing_strategy"][value="advanced"]').checked = true;
    }

    const event = new Event('change', { bubbles: true });
    document.querySelector('input[name="balancing_strategy"]:checked').dispatchEvent(event);

    if (settings.swapAttempts !== undefined) {
        document.getElementById('swap-attempts-count').value = settings.swapAttempts;
    }
    if (settings.polishingSwaps !== undefined) {
        document.getElementById('polishing-swaps-count').value = settings.polishingSwaps;
    }
    if (settings.annealingTemp !== undefined) {
        document.getElementById('annealing-temp').value = settings.annealingTemp;
    }
    if (settings.annealingCooling !== undefined) {
        document.getElementById('annealing-cooling').value = settings.annealingCooling;
    }
    if (settings.annealingIterations !== undefined) {
        document.getElementById('annealing-iterations').value = settings.annealingIterations;
    }
    if (settings.solverTimelimit !== undefined) {
        document.getElementById('solver-timelimit').value = settings.solverTimelimit;
    }
    if (settings.geneticPopulation !== undefined) {
        document.getElementById('genetic-population').value = settings.geneticPopulation;
    }
    if (settings.geneticGenerations !== undefined) {
        document.getElementById('genetic-generations').value = settings.geneticGenerations;
    }
    if (settings.geneticElitism !== undefined) {
        document.getElementById('genetic-elitism').value = settings.geneticElitism;
    }
    if (settings.geneticMutation !== undefined) {
        document.getElementById('genetic-mutation').value = settings.geneticMutation;
    }

    if (settings.enableCustomTargets) {
        document.getElementById('enable-custom-targets-checkbox').checked = true;
        document.getElementById('custom-targets-controls').classList.remove('hidden');
    }
    if (settings.customTargetPatterns && Array.isArray(settings.customTargetPatterns)) {
        customTargetPatterns = settings.customTargetPatterns;
        renderCustomTargetsTable();
    }
    if (settings.tabuIterations !== undefined) {
         document.getElementById('tabu-iterations').value = settings.tabuIterations;
    }
    if (settings.tabuTenure !== undefined) {
        document.getElementById('tabu-tenure').value = settings.tabuTenure;
    }
    if (settings.tabuNeighborhoodSize !== undefined) {
        document.getElementById('tabu-neighborhood-size').value = settings.tabuNeighborhoodSize;
    }
    if (settings.lnsIterations !== undefined) {
        document.getElementById('lns-iterations').value = settings.lnsIterations;
    }
    if (settings.lnsDestroyFraction !== undefined) {
        document.getElementById('lns-destroy-fraction').value = settings.lnsDestroyFraction;
    }
    if (settings.vnsIterations !== undefined) {
        document.getElementById('vns-iterations').value = settings.vnsIterations;
    }
    if (settings.vnsMaxK !== undefined) {
        document.getElementById('vns-max-k').value = settings.vnsMaxK;
    }
    setupProfessorPartnershipsUI(settings, availableProfessors);
}

function filterList(listId, searchTerm) {
    const lowerCaseSearchTerm = searchTerm.toLowerCase();
    document.querySelectorAll(`#${listId} li`).forEach(li => {
        const itemText = li.textContent.toLowerCase();
        li.style.display = itemText.includes(lowerCaseSearchTerm) ? '' : 'none';
    });
}
function handleProfessorInteraction(event) {
    const li = event.target.closest('li');
    if (!li) return;
    if (event.target.classList.contains('courses-dropdown-btn')) {
        event.stopPropagation();
        const dropdown = li.querySelector('.courses-dropdown-list');
        if (dropdown) {
            document.querySelectorAll('.courses-dropdown-list').forEach(d => {
                if (d !== dropdown) d.style.display = 'none';
            });
            dropdown.style.display = dropdown.style.display === 'block' ? 'none' : 'block';
        }
        return;
    }
    document.querySelectorAll('#professors-list-assign li.selected').forEach(el => el.classList.remove('selected'));
    li.classList.add('selected');
    selectedProfessorForAssign = li.dataset.professorName;
    selectedSubjectsForAssign = [];
    document.querySelectorAll('#subjects-list-assign li.selected').forEach(el => el.classList.remove('selected'));
    updateAssignButtonState();
}
function handleSubjectSelect(event) {
    const li = event.currentTarget;
    if (li.classList.contains('assigned-subj')) return;
    li.classList.toggle('selected');
    
    const subjectInfo = { uniqueId: li.dataset.uniqueId };
    const index = selectedSubjectsForAssign.findIndex(s => s.uniqueId === subjectInfo.uniqueId);

    if (index > -1) {
        selectedSubjectsForAssign.splice(index, 1);
    } else {
        selectedSubjectsForAssign.push(subjectInfo);
    }
    updateAssignButtonState();
}
function updateAssignButtonState() {
    const assignButton = document.getElementById('assign-subjects-button');
    assignButton.disabled = !(selectedProfessorForAssign && selectedSubjectsForAssign.length > 0);
}
function handleResponse(response) {
    return response.json().then(data => {
        if (!response.ok) {
            // <<< تعديل: استدعاء showNotification مباشرة من handleError
            const error = new Error(data.error || data.message || 'Server error');
            handleError(error); 
            throw error; // استمر في طرح الخطأ لإيقاف العمليات اللاحقة
        }
        if(data.message && !data.success && response.status !== 201) {
             showNotification(data.message, 'error'); // يمكن استخدامها هنا أيضًا
        }
        return data;
    });
}
function handleError(error) {
    console.error('Error:', error);
    const errorMessage = error.message.includes('Failed to fetch') 
        ? 'فشل الاتصال بالخادم. يرجى التأكد من أن الخادم يعمل.'
        : error.message;
    showNotification(errorMessage, 'error'); // <<< تعديل: استخدام الدالة الجديدة للخطأ
}

function setupBackupRestoreListeners() {
    const backupBtn = document.getElementById('backup-btn');
    const restoreBtn = document.getElementById('restore-btn');
    const fileInput = document.getElementById('restore-file-input');
    const resetBtn = document.getElementById('reset-all-btn');

    backupBtn.addEventListener('click', async () => {
        try {
            const response = await fetch('/api/backup');
            if (!response.ok) throw new Error('فشل النسخ الاحتياطي من الخادم');
            
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
            a.href = url;
            a.download = `backup_${timestamp}.json`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            alert('تم تصدير النسخة الاحتياطية بنجاح.');
        } catch (error) {
            handleError(error);
            alert('حدث خطأ أثناء تصدير النسخة الاحتياطية.');
        }
    });

    restoreBtn.addEventListener('click', () => {
        fileInput.click();
    });

    resetBtn.addEventListener('click', () => {
        if (confirm("تحذير! هل أنت متأكد تماماً؟ سيؤدي هذا إلى حذف جميع الأساتذة والقاعات والمواد والإعدادات بشكل نهائي.")) {
            if(confirm("التأكيد الأخير: هل تريد حقاً مسح كل شيء؟ لا يمكن التراجع عن هذا الإجراء.")){
                 fetch('/api/reset-all', { method: 'POST' })
                    .then(handleResponse)
                    .then(res => {
                        alert(res.message);
                        location.reload();
                    })
                    .catch(handleError);
            }
        }
    });

    fileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) return;

        if (!confirm("هل أنت متأكد من أنك تريد استعادة البيانات من هذا الملف؟ سيتم الكتابة فوق جميع البيانات والإعدادات الحالية.")) {
            fileInput.value = '';
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = JSON.parse(e.target.result);
                fetch('/api/restore', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(data),
                })
                .then(handleResponse)
                .then(res => {
                    alert(res.message);
                    location.reload();
                })
                .catch(handleError);
            } catch (error) {
                alert('خطأ في قراءة الملف. يرجى التأكد من أنه ملف نسخة احتياطية صالح.');
                handleError(error);
            }
        };
        reader.readAsText(file);
        fileInput.value = '';
    });
}

function displayWorkloadChart(chartData) {
    const chartContainer = document.getElementById('chart-container');
    chartContainer.style.display = 'block';

    const ctx = document.getElementById('workload-chart').getContext('2d');

    if (workloadChartInstance) {
        workloadChartInstance.destroy();
    }

    workloadChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: chartData.labels,
            datasets: chartData.datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            scales: {
                x: {
                    stacked: true,
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    ticks: {
                        stepSize: 1
                    }
                }
            },
            plugins: {
                legend: {
                    position: 'top',
                },
                title: {
                    display: false
                }
            }
        }
    });
}

function displayBalanceReport(data) {
    const container = document.getElementById('balance-report-area');
    if (!data || !data.details) {
        container.innerHTML = '';
        return;
    };
    
    function generateDistributionRows(details) {
        if (!details) return '';
        return details.map(item => `
            <tr>
                <td>${item.pattern}</td>
                <td>${item.target_count}</td>
                <td>${item.actual_count}</td>
                <td style="color: ${item.deviation === 0 ? 'green' : 'red'};">${item.deviation > 0 ? '+' : ''}${item.deviation}</td>
            </tr>
        `).join('');
    }

    container.innerHTML = `
        <div class="target-distribution-report">
            <h4>تقرير توازن توزيع الحراسة</h4>
            <table class="distribution-table">
                <thead>
                    <tr>
                        <th>نمط التوزيع</th>
                        <th>العدد المستهدف من الأساتذة</th>
                        <th>العدد الفعلي</th>
                        <th>الانحراف</th>
                    </tr>
                </thead>
                <tbody>
                    ${generateDistributionRows(data.details)}
                </tbody>
            </table>
            <div class="balance-indicator">
                <span>مؤشر التوازن: </span>
                <div class="progress-bar">
                    <div class="progress" style="width: ${data.balance_score}%">
                        ${data.balance_score}%
                    </div>
                </div>
            </div>
        </div>
    `;
}

function setupCustomTargetListeners() {
    const enableCheckbox = document.getElementById('enable-custom-targets-checkbox');
    const controlsDiv = document.getElementById('custom-targets-controls');
    const addBtn = document.getElementById('add-custom-target-btn');
    const tableBody = document.querySelector('#custom-targets-table tbody');

    enableCheckbox.addEventListener('change', () => {
        controlsDiv.classList.toggle('hidden', !enableCheckbox.checked);
    });

    addBtn.addEventListener('click', () => {
        const largeInput = document.getElementById('custom-target-large');
        const otherInput = document.getElementById('custom-target-other');
        const countInput = document.getElementById('custom-target-prof-count');

        const large = parseInt(largeInput.value, 10) || 0;
        const other = parseInt(otherInput.value, 10) || 0;
        const count = parseInt(countInput.value, 10);

        if (isNaN(count) || count <= 0) {
            alert('الرجاء إدخال عدد صحيح وموجب للأساتذة.');
            return;
        }

        customTargetPatterns.push({ large, other, count });
        renderCustomTargetsTable();

        largeInput.value = '';
        otherInput.value = '';
        countInput.value = '';
    });

    tableBody.addEventListener('click', (e) => {
        if (e.target.classList.contains('delete-target-btn')) {
            const index = parseInt(e.target.dataset.index, 10);
            if (!isNaN(index)) {
                customTargetPatterns.splice(index, 1);
                renderCustomTargetsTable();
            }
        }
    });
}

function renderCustomTargetsTable() {
    const tableBody = document.querySelector('#custom-targets-table tbody');
    tableBody.innerHTML = '';
    let totalProfs = 0;

    customTargetPatterns.forEach((pattern, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${pattern.large} كبيرة + ${pattern.other} أخرى</td>
            <td>${pattern.count}</td>
            <td>
                <button type="button" class="delete-target-btn" data-index="${index}">&times;</button>
            </td>
        `;
        tableBody.appendChild(row);
        totalProfs += pattern.count;
    });

    const totalProfsP = document.getElementById('custom-target-prof-total');
    totalProfsP.textContent = `الإجمالي: ${totalProfs} أستاذًا في الأنماط المخصصة.`;

    const allProfsCount = availableProfessors.length;
    if (totalProfs > allProfsCount) {
        totalProfsP.style.color = 'red';
        totalProfsP.textContent += ` (تحذير: العدد يتجاوز إجمالي الأساتذة ${allProfsCount}!)`;
    } else if (totalProfs < allProfsCount) {
         totalProfsP.style.color = '#e0a800';
         totalProfsP.textContent += ` (ملاحظة: العدد أقل من إجمالي الأساتذة ${allProfsCount}. سيتم توزيع الباقي تلقائياً.)`;
    }
     else {
        totalProfsP.style.color = 'green';
    }
}

function setupDataImportExportListeners() {
    const exportBtn = document.getElementById('export-template-btn');
    const importBtn = document.getElementById('import-data-btn');
    const fileInput = document.getElementById('import-file-input');

    exportBtn.addEventListener('click', async () => {
        try {
            const response = await fetch('/api/data-template');
            if (!response.ok) throw new Error('فشل تصدير القالب من الخادم');
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'قالب_إدخال_البيانات.xlsx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
        } catch (error) {
            handleError(error);
            alert('حدث خطأ أثناء تصدير ملف القالب.');
        }
    });

    importBtn.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) return;

        if (!confirm("هل أنت متأكد من استيراد البيانات من هذا الملف؟ سيتم إضافة البيانات الجديدة فقط ولن يتم حذف البيانات الحالية.")) {
            fileInput.value = '';
            return;
        }

        const formData = new FormData();
        formData.append('file', file);

        fetch('/api/import-data', {
            method: 'POST',
            body: formData,
        })
        .then(handleResponse)
        .then(data => {
            alert(data.message);
            location.reload();
        })
        .catch(error => {})
        .finally(() => {
            fileInput.value = '';
        });
    });
}

// في ملف script.js، استبدل هذه الدالة بالكامل

// في ملف script.js، استبدل هذه الدالة بالكامل

function displayStatsDashboard(stats) {
    const container = document.getElementById('stats-dashboard');
    const containerWrapper = document.getElementById('stats-dashboard-container');
    if (!container || !stats) {
        containerWrapper.style.display = 'none';
        return;
    }

    let dashboardHTML = `
        <div class="stat-card">
            <h4>إجمالي الحصص الموزعة</h4>
            <p>${stats.total_duties}</p>
            <div class="sub-stat">قاعات كبيرة: ${stats.total_large_duties} | قاعات أخرى: ${stats.total_other_duties}</div>
        </div>
        <div class="stat-card">
            <h4>متوسط الحصص لكل أستاذ</h4>
            <p>${stats.avg_duties_per_prof.toFixed(2)}</p>
        </div>
        <div class="stat-card">
            <h4>اليوم الأكثر ازدحاماً</h4>
            <p>${stats.busiest_day.date || 'N/A'}</p>
            <div class="sub-stat">بمجموع ${stats.busiest_day.duties} حصص حراسة</div>
        </div>
        <div class="stat-card">
            <h4>الأكثر حراسة (حسب العبء)</h4>
            <ul>
                ${stats.most_burdened_profs.map(p => `<li><span>${p.name}:</span> ${p.workload.toFixed(1)} نقطة</li>`).join('')}
            </ul>
        </div>
         <div class="stat-card">
            <h4>الأقل حراسة (حسب العبء)</h4>
            <ul>
                ${stats.least_burdened_profs.map(p => `<li><span>${p.name}:</span> ${p.workload.toFixed(1)} نقطة</li>`).join('')}
            </ul>
        </div>
    `;

    // --- بداية المنطق الجديد والمحسن للمربع السادس ---
    const hasGuardShortages = stats.shortage_reports && stats.shortage_reports.length > 0;
    const hasUnscheduledSubjects = stats.unscheduled_subjects_report && stats.unscheduled_subjects_report.length > 0;

    let reportContentHTML = '';
    let reportCardClass = 'stat-card'; // التنسيق الافتراضي
    let reportTitle = 'تقارير الملاحظات'; // عنوان عام

    if (hasGuardShortages || hasUnscheduledSubjects) {
        reportCardClass = 'stat-card shortage-report'; // تطبيق تنسيق التحذير

        if (hasUnscheduledSubjects) {
            const subjectItems = stats.unscheduled_subjects_report.map(item => `<li>${item}</li>`).join('');
            reportContentHTML += `
                <div class="report-subsection">
                    <h5>مواد لم تتم جدولتها (${stats.unscheduled_subjects_report.length})</h5>
                    <ul>${subjectItems}</ul>
                </div>
            `;
        }

        if (hasGuardShortages) {
            const guardItems = stats.shortage_reports.map(item => `<li>${item}</li>`).join('');
            reportContentHTML += `
                <div class="report-subsection">
                    <h5>نقص في الحراسة (${stats.shortage_reports.length})</h5>
                    <ul>${guardItems}</ul>
                </div>
            `;
        }
    } else {
        // في حالة عدم وجود أي نقص على الإطلاق
        reportContentHTML = `
            <p class="no-shortage-message">
                ✅ لا يوجد نقص في المواد أو الحراسة. تم إنجاز الجدول بنجاح.
            </p>
        `;
    }
    
    const reportCardHTML = `
        <div class="${reportCardClass}">
            <h4>${reportTitle}</h4>
            ${reportContentHTML}
        </div>
    `;

    dashboardHTML += reportCardHTML;
    // --- نهاية المنطق الجديد ---

    container.innerHTML = dashboardHTML;
    containerWrapper.style.display = 'block';
}

function setupResultsSearch() {
    const searchInput = document.getElementById('results-search-input');
    if (!searchInput) return;

    searchInput.addEventListener('input', (e) => {
        const searchTerm = e.target.value.trim().toLowerCase();
        const allTables = document.querySelectorAll('.results-grid-table');

        if (searchTerm === '') {
            document.querySelectorAll('.highlight-search').forEach(cell => {
                cell.classList.remove('highlight-search');
            });
            return;
        }

        allTables.forEach(table => {
            const cells = table.getElementsByTagName('td');
            for (const cell of cells) {
                const cellText = cell.textContent.toLowerCase();
                if (cellText.includes(searchTerm)) {
                    cell.classList.add('highlight-search');
                } else {
                    cell.classList.remove('highlight-search');
                }
            }
        });
    });
}

// =======================================================================
// ========== بداية: منطق إدارة اشتراكات الأساتذة (للعمل معاً) ==========
// =======================================================================

// متغير للاحتفاظ بحالة الاشتراكات الحالية في الواجهة
let currentProfessorPartnerships = [];

/**
 * الدالة الرئيسية لإعداد واجهة اشتراكات الأساتذة.
 * @param {object} currentSettings - كائن الإعدادات الحالي المحمل من الخادم.
 * @param {Array} allProfessorData - مصفوفة الأساتذة الكاملة.
 */
function setupProfessorPartnershipsUI(currentSettings, allProfessorData) {
    const allProfessorNames = allProfessorData.map(p => p.name);
    // استخدام المفتاح الجديد 'professorPartnerships' من الإعدادات
    currentProfessorPartnerships = currentSettings.professorPartnerships || [];

    populatePartnershipDropdowns(allProfessorNames);
    renderPartnershipsList();

    // ربط الأحداث بالأزرار (مرة واحدة فقط لمنع التكرار)
    const addBtn = document.getElementById('add-pair-btn');
    if (!addBtn.hasAttribute('data-listener-attached')) {
        addBtn.addEventListener('click', handleAddPartnership);
        addBtn.setAttribute('data-listener-attached', 'true');
        
        document.getElementById('prof-pairs-list').addEventListener('click', handleDeletePartnership);
    }
}

/**
 * دالة لملء القوائم المنسدلة بالأساتذة المتاحين فقط للاشتراك.
 * @param {Array<string>} allProfessorNames - قائمة أسماء كل الأساتذة.
 */
function populatePartnershipDropdowns(allProfessorNames) {
    const select1 = document.getElementById('prof-select-1');
    const select2 = document.getElementById('prof-select-2');
    
    // قائمة الأساتذة الذين تم اختيارهم بالفعل في اشتراكات
    const partneredProfessors = currentProfessorPartnerships.flat();
    const availableProfessors = allProfessorNames.filter(p => !partneredProfessors.includes(p));

    // حفظ القيمة المختارة حالياً (إن وجدت)
    const selectedValue1 = select1.value;

    // مسح الخيارات القديمة
    select1.innerHTML = '<option value="">-- اختر الأستاذ الأول --</option>';
    select2.innerHTML = '<option value="">-- اختر الأستاذ الثاني --</option>';

    // ملء القائمة الأولى
    availableProfessors.forEach(prof => {
        select1.innerHTML += `<option value="${prof}">${prof}</option>`;
    });

    // ملء القائمة الثانية عند تغيير الأولى
    select1.onchange = () => {
        const selectedProf = select1.value;
        select2.innerHTML = '<option value="">-- اختر الأستاذ الثاني --</option>';
        availableProfessors.forEach(prof => {
            if (prof !== selectedProf) { // منع ظهور نفس الأستاذ
                select2.innerHTML += `<option value="${prof}">${prof}</option>`;
            }
        });
    };
    
    // استعادة القيمة المختارة إذا كانت لا تزال متاحة
    if (availableProfessors.includes(selectedValue1)) {
        select1.value = selectedValue1;
        select1.dispatchEvent(new Event('change')); // تفعيل حدث التغيير لتحديث القائمة الثانية
    }
}

/**
 * دالة لعرض قائمة الاشتراكات المحددة في الواجهة.
 */
function renderPartnershipsList() {
    const listElement = document.getElementById('prof-pairs-list'); // الـ ID لم يتغير في HTML
    const noPairsMsg = document.getElementById('no-pairs-msg');
    listElement.innerHTML = ''; 

    if (currentProfessorPartnerships.length === 0) {
        noPairsMsg.style.display = 'block';
    } else {
        noPairsMsg.style.display = 'none';
        currentProfessorPartnerships.forEach((partnership, index) => {
            const listItem = document.createElement('li');
            listItem.className = 'list-group-item d-flex justify-content-between align-items-center';
            listItem.innerHTML = `
                <span>${partnership[0]} &nbsp; <i class="fas fa-arrows-alt-h"></i> &nbsp; ${partnership[1]}</span>
                <button class="btn btn-danger btn-sm" data-index="${index}">حذف</button>
            `;
            listElement.appendChild(listItem);
        });
    }
}

/**
 * معالج حدث الضغط على زر "إضافة مشتركين".
 */
function handleAddPartnership() {
    const select1 = document.getElementById('prof-select-1');
    const select2 = document.getElementById('prof-select-2');
    const prof1 = select1.value;
    const prof2 = select2.value;

    if (!prof1 || !prof2) { alert("الرجاء اختيار أستاذين."); return; }
    
    currentProfessorPartnerships.push([prof1, prof2]);
    const allProfessorNames = availableProfessors.map(p => p.name);
    
    renderPartnershipsList();
    populatePartnershipDropdowns(allProfessorNames);
}

/**
 * معالج حدث الضغط على زر "حذف".
 */
function handleDeletePartnership(event) {
    if (event.target.tagName === 'BUTTON' && event.target.hasAttribute('data-index')) {
        const indexToDelete = parseInt(event.target.getAttribute('data-index'));
        
        currentProfessorPartnerships.splice(indexToDelete, 1);

        renderPartnershipsList();
        populatePartnershipDropdowns(availableProfessors.map(p => p.name));
    }
}

// =======================================================================
// ========== نهاية: منطق إدارة اشتراكات الأساتذة (للعمل معاً) ==========
// =======================================================================

// ============== بداية: منطق أداة حساب التوزيع العادل ==============

// ربط الحدث بالزر عند تحميل الصفحة
document.addEventListener('DOMContentLoaded', () => {
    const calcButton = document.getElementById('run-calculator-btn');
    if(calcButton) {
        calcButton.addEventListener('click', onCalculateDistributionClick);
    }
});

/**
 * دالة تُنفذ عند الضغط على زر "احسب التوزيع".
 */
function onCalculateDistributionClick() {
    try {
        // 1. قراءة القيم من حقول الإدخال
        const profs = parseInt(document.getElementById('calc-profs').value);
        const largeSlots = parseInt(document.getElementById('calc-large').value);
        const otherSlots = parseInt(document.getElementById('calc-other').value);
        const factor = parseFloat(document.getElementById('calc-factor').value);

        // التحقق من صحة المدخلات
        if (isNaN(profs) || isNaN(largeSlots) || isNaN(otherSlots) || isNaN(factor)) {
            showNotification("الرجاء ملء جميع الحقول بأرقام صحيحة.", 'error');
            return;
        }
        if (profs <= 0) {
            showNotification("عدد الأساتذة يجب أن يكون أكبر من صفر.", 'error');
            return;
        }

        // 2. استدعاء الدالة المنطقية للحصول على النتائج
        const results = suggestFairDistribution(profs, largeSlots, otherSlots, factor);

        // 3. عرض النتائج الجديدة في الجدول
        displayCalculationResults(results);

    } catch (e) {
        console.error("خطأ في حساب التوزيع:", e);
        showNotification("حدث خطأ غير متوقع أثناء الحساب.", 'error');
    }
}


/**
 * ترجمة للدالة المنطقية من بايثون إلى جافاسكريبت.
 * (النسخة المصححة لتطابق سلوك Python min() تمامًا)
 */
function suggestFairDistribution(totalProfs, largeHallSlots, otherHallSlots, workloadFactor) {
    if (totalProfs <= 0) return [];

    let professors = Array.from({ length: totalProfs }, (_, i) => ({
        id: i,
        large_halls: 0,
        other_halls: 0,
        workload: 0
    }));

    // --- بداية التعديل الجوهري: محاكاة دالة min() ---
    const findProfWithMinLoad = (profsArray) => {
        if (profsArray.length === 0) return null;
        let minProf = profsArray[0];
        for (let i = 1; i < profsArray.length; i++) {
            if (profsArray[i].workload < minProf.workload) {
                minProf = profsArray[i];
            }
        }
        return minProf;
    };
    // --- نهاية التعديل ---

    for (let i = 0; i < largeHallSlots; i++) {
        const profToUpdate = findProfWithMinLoad(professors);
        profToUpdate.large_halls += 1;
        profToUpdate.workload += workloadFactor;
    }

    for (let i = 0; i < otherHallSlots; i++) {
        const profToUpdate = findProfWithMinLoad(professors);
        profToUpdate.other_halls += 1;
        profToUpdate.workload += 1;
    }

    const distributionSummary = new Map();
    for (const p of professors) {
        const key = `${p.large_halls}-${p.other_halls}`;
        distributionSummary.set(key, (distributionSummary.get(key) || 0) + 1);
    }
    
    const results = [];
    for (const [plan, count] of distributionSummary.entries()) {
        const [largeDuties, otherDuties] = plan.split('-').map(Number);
        const workload = (largeDuties * workloadFactor) + (otherDuties * 1);
        results.push({
            "count": count,
            "large_duties": largeDuties,
            "other_duties": otherDuties,
            "workload": workload
        });
    }

    return results.sort((a, b) => b.workload - a.workload);
}



// script.js

// ============== بداية: منطق أداة حساب التوزيع العادل ==============

// ربط الأحداث بالأزرار عند تحميل الصفحة
document.addEventListener('DOMContentLoaded', () => {
    const calcButton = document.getElementById('run-calculator-btn');
    if (calcButton) {
        calcButton.addEventListener('click', onCalculateDistributionClick);
    }
    const autoFillButton = document.getElementById('autofill-calculator-btn');
    if(autoFillButton) {
        autoFillButton.addEventListener('click', autofillCalculatorFromSchedule);
    }
});

/**
 * دالة تُنفذ عند الضغط على زر "احسب التوزيع".
 */
function onCalculateDistributionClick() {
    try {
        const profs = parseInt(document.getElementById('calc-profs').value);
        const largeSlots = parseInt(document.getElementById('calc-large').value);
        const otherSlots = parseInt(document.getElementById('calc-other').value);
        const factor = parseFloat(document.getElementById('calc-factor').value);

        if (isNaN(profs) || isNaN(largeSlots) || isNaN(otherSlots) || isNaN(factor)) {
            showNotification("الرجاء ملء جميع الحقول بأرقام صحيحة.", 'error');
            return;
        }
        if (profs <= 0) {
            showNotification("عدد الأساتذة يجب أن يكون أكبر من صفر.", 'error');
            return;
        }

        const results = suggestFairDistribution(profs, largeSlots, otherSlots, factor);
        displayCalculationResults(results);

    } catch (e) {
        console.error("خطأ في حساب التوزيع:", e);
        showNotification("حدث خطأ غير متوقع أثناء الحساب.", 'error');
    }
}

/**
 * دالة تقوم بحساب البيانات من إعدادات البرنامج وتملأ حقول الأداة.
 */
function autofillCalculatorFromSchedule() {
    try {
        const profCount = availableProfessors.length;
        const guardsPerLarge = parseInt(document.getElementById('guards-large-hall').value) || 0;
        const guardsPerMedium = parseInt(document.getElementById('guards-medium-hall').value) || 0;
        const guardsPerSmall = parseInt(document.getElementById('guards-small-hall').value) || 0;

        const levelHallAssignments = {};
        document.querySelectorAll('#level-halls-assignment-container tbody tr').forEach(row => {
            const levelName = row.dataset.levelName;
            levelHallAssignments[levelName] = [];
            row.querySelectorAll('input[type="checkbox"]:checked').forEach(checkbox => {
                levelHallAssignments[levelName].push(checkbox.value);
            });
        });

        let totalLargeDuties = 0;
        let totalOtherDuties = 0;
        availableSubjects.forEach(subject => {
            const levelName = subject.level;
            const assignedHalls = levelHallAssignments[levelName] || [];
            assignedHalls.forEach(hallName => {
                const hallInfo = availableHalls.find(h => h.name === hallName);
                if (hallInfo) {
                    if (hallInfo.type === 'كبيرة') {
                        totalLargeDuties += guardsPerLarge;
                    } else if (hallInfo.type === 'متوسطة') {
                        totalOtherDuties += guardsPerMedium;
                    } else if (hallInfo.type === 'صغيرة') {
                        totalOtherDuties += guardsPerSmall;
                    }
                }
            });
        });

        document.getElementById('calc-profs').value = profCount;
        document.getElementById('calc-large').value = totalLargeDuties;
        document.getElementById('calc-other').value = totalOtherDuties;
        showNotification("تم حساب وملء الحقول بنجاح.", 'success');
    } catch (e) {
        console.error("خطأ في الحساب التلقائي:", e);
        showNotification("حدث خطأ أثناء محاولة الحساب التلقائي.", 'error');
    }
}

/**
 * ترجمة للدالة المنطقية من بايثون إلى جافاسكريبت.
 * (النسخة المصححة لتطابق سلوك Python min() تمامًا)
 */
function suggestFairDistribution(totalProfs, largeHallSlots, otherHallSlots, workloadFactor) {
    if (totalProfs <= 0) return [];

    let professors = Array.from({ length: totalProfs }, (_, i) => ({
        id: i,
        large_halls: 0,
        other_halls: 0,
        workload: 0
    }));

    const findProfWithMinLoad = (profsArray) => {
        if (profsArray.length === 0) return null;
        let minProf = profsArray[0];
        for (let i = 1; i < profsArray.length; i++) {
            if (profsArray[i].workload < minProf.workload) {
                minProf = profsArray[i];
            }
        }
        return minProf;
    };

    for (let i = 0; i < largeHallSlots; i++) {
        const profToUpdate = findProfWithMinLoad(professors);
        profToUpdate.large_halls += 1;
        profToUpdate.workload += workloadFactor;
    }

    for (let i = 0; i < otherHallSlots; i++) {
        const profToUpdate = findProfWithMinLoad(professors);
        profToUpdate.other_halls += 1;
        profToUpdate.workload += 1;
    }

    const distributionSummary = new Map();
    for (const p of professors) {
        const key = `${p.large_halls}-${p.other_halls}`;
        distributionSummary.set(key, (distributionSummary.get(key) || 0) + 1);
    }
    
    const results = [];
    for (const [plan, count] of distributionSummary.entries()) {
        const [largeDuties, otherDuties] = plan.split('-').map(Number);
        const workload = (largeDuties * workloadFactor) + (otherDuties * 1);
        results.push({
            "count": count,
            "large_duties": largeDuties,
            "other_duties": otherDuties,
            "workload": workload
        });
    }

    return results.sort((a, b) => b.workload - a.workload);
}

/**
 * دالة لعرض النتائج في جدول HTML ديناميكي.
 */
function displayCalculationResults(results) {
    const container = document.getElementById('calculator-results');
    if (results.length === 0) {
        container.innerHTML = "<p>لا توجد نتائج لعرضها.</p>";
        return;
    }

    let tableHTML = `
        <table class="distribution-table">
            <thead>
                <tr>
                    <th>عدد الأساتذة</th>
                    <th>حراسات (كبيرة)</th>
                    <th>حراسات (أخرى)</th>
                    <th>نقاط العبء للفرد</th>
                </tr>
            </thead>
            <tbody>
    `;

    results.forEach(row => {
        tableHTML += `
            <tr>
                <td>${row.count}</td>
                <td>${row.large_duties}</td>
                <td>${row.other_duties}</td>
                <td>${row.workload.toFixed(2)}</td>
            </tr>
        `;
    });

    tableHTML += `</tbody></table>`;
    container.innerHTML = tableHTML;
}

// ============== نهاية: منطق أداة حساب التوزيع العادل ==============

// أضف هذا الكود في نهاية ملف script.js

async function exportScheduleWord() {
    if (!lastGeneratedSchedule) {
        alert("يرجى إنشاء جدول أولاً قبل التصدير.");
        return;
    }
    const button = document.getElementById('export-schedule-word-button');
    button.disabled = true;
    button.textContent = 'جاري التصدير...';

    try {
        const response = await fetch('/api/export/word/all-exams', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(lastGeneratedSchedule)
        });

        if (!response.ok) throw new Error('فشل التصدير من الخادم');

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'جداول_الامتحانات.docx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

    } catch (err) {
        alert('حدث خطأ أثناء تصدير الملف.');
        console.error(err);
    } finally {
        button.disabled = false;
        button.textContent = 'تصدير جداول الامتحانات (Word)';
    }
}

async function exportProfScheduleWord() {
    if (!lastGeneratedSchedule) {
        alert("يرجى إنشاء جدول أولاً قبل التصدير.");
        return;
    }
    const button = document.getElementById('export-prof-word-button');
    button.disabled = true;
    button.textContent = 'جاري التصدير...';

    try {
        const response = await fetch('/api/export/word/all-profs', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(lastGeneratedSchedule)
        });

        if (!response.ok) throw new Error('فشل التصدير من الخادم');

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'جداول_الحراسة_للأساتذة.docx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

    } catch (err) {
        alert('حدث خطأ أثناء تصدير الملف.');
        console.error(err);
    } finally {
        button.disabled = false;
        button.textContent = 'تصدير جداول الأساتذة (Word)';
    }
}

async function exportProfScheduleAnonymous() {
    if (!lastGeneratedSchedule) {
        alert("يرجى إنشاء جدول أولاً قبل التصدير.");
        return;
    }
    const button = document.getElementById('export-prof-anonymous-word-button');
    button.disabled = true;
    button.textContent = 'جاري التصدير...';

    try {
        const response = await fetch('/api/export/word/all-profs-anonymous', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(lastGeneratedSchedule)
        });

        if (!response.ok) throw new Error('فشل التصدير من الخادم');

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'جداول_الحراسة_المبسطة.docx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

    } catch (err) {
        alert('حدث خطأ أثناء تصدير الملف.');
        console.error(err);
    } finally {
        button.disabled = false;
        button.textContent = 'تصدير جداول الأساتذة (مُبسَّط)';
    }
}