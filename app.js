// DOM Elements
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const uploadSection = document.getElementById('upload-section');
const dashboardSection = document.getElementById('dashboard-section');
const resetBtn = document.getElementById('reset-btn');
const themeToggle = document.getElementById('theme-toggle');

// Chart Instances
let ratioChartInst, group1ChartInst, group2ChartInst;

// Theme logic
themeToggle.addEventListener('click', () => {
    const currentTheme = document.body.getAttribute('data-theme') || 'dark';
    const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
    document.body.setAttribute('data-theme', newTheme);
    themeToggle.innerHTML = newTheme === 'dark' ? '<i class="fa-solid fa-sun"></i>' : '<i class="fa-solid fa-moon"></i>';
    
    // Update chart colors if they exist
    updateChartTheme();
});

function updateChartTheme() {
    const isLight = document.body.getAttribute('data-theme') === 'light';
    Chart.defaults.color = isLight ? '#64748b' : '#94a3b8';
    if(ratioChartInst) ratioChartInst.update();
    if(group1ChartInst) group1ChartInst.update();
    if(group2ChartInst) group2ChartInst.update();
}

// Drag & Drop Handlers
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
        handleFile(e.dataTransfer.files[0]);
    }
});

dropZone.addEventListener('click', () => {
    fileInput.click();
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) {
        handleFile(e.target.files[0]);
    }
});

resetBtn.addEventListener('click', () => {
    dashboardSection.classList.add('hidden');
    uploadSection.classList.add('active-section');
    fileInput.value = '';
    // Destroy charts
    if(ratioChartInst) ratioChartInst.destroy();
    if(group1ChartInst) group1ChartInst.destroy();
    if(group2ChartInst) group2ChartInst.destroy();
});

// File Processing
function handleFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            
            // Assume the first sheet is the one we want
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
            
            processData(json);
        } catch (error) {
            console.error(error);
            alert("Lỗi khi đọc file. Vui lòng đảm bảo đây là file Excel hợp lệ.");
        }
    };
    reader.readAsArrayBuffer(file);
}

// Department Mapping Definition
const GROUP_1 = {
    name: "Phát Triển Kinh Doanh & Sản Phẩm",
    departments: {
        "Phân tích kinh doanh & Báo cáo": ["phân tích kinh doanh", "báo cáo", "business analyst", "ba"],
        "Business Analytics": ["business analytics", "data analytics", "data analyst", "da"],
        "Product Owner & Design": ["product owner", "po", "design", "ui", "ux", "thiết kế"],
        "R&D (Research & Development)": ["r&d", "research", "development", "nghiên cứu"],
        "Project Management": ["project management", "pm", "quản lý dự án", "scrum"]
    }
};

const GROUP_2 = {
    name: "Kỹ Thuật & Vận Hành",
    departments: {
        "Tech Spec & CTO": ["tech spec", "cto", "chief tech", "kiến trúc sư"],
        "Phát triển Backend & Frontend": ["backend", "frontend", "dev", "lập trình", "software engineer", "se"],
        "QC (Quality Control)": ["qc", "quality control", "tester", "qa"],
        "Operation Logistics": ["operation", "logistics", "vận hành", "ops"]
    }
};

function normalizeString(str) {
    if (!str) return "";
    return String(str).toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function detectDepartmentColumn(row) {
    const keywords = ["phong ban", "bo phan", "department", "dept", "khoi", "chuyen mon", "to"];
    for (let key of Object.keys(row)) {
        const normKey = normalizeString(key);
        if (keywords.some(k => normKey.includes(k))) return key;
    }
    // If not found by name, guess by looking at values in the first few rows (not implemented for simplicity, fallback to generic approach)
    return null;
}

function processData(data) {
    if (data.length === 0) {
        alert("File không có dữ liệu.");
        return;
    }

    // Try to find the department column
    let deptCol = detectDepartmentColumn(data[0]);
    if (!deptCol) {
        // Fallback: ask user or just pick the 3rd/4th column if it's a typical HR file
        // Here we just pick the first string column that has diverse values, or prompt.
        const keys = Object.keys(data[0]);
        // Simple heuristic: column name containing 'phòng' or 'ban'
        deptCol = keys.find(k => k.toLowerCase().includes('phòng') || k.toLowerCase().includes('ban') || k.toLowerCase().includes('dept'));
        if(!deptCol) {
           // If we still can't find it, we just take the last column as a wild guess or ask user. 
           // For robust static app without prompt, let's take the 3rd column or last if length < 3
           deptCol = keys.length > 2 ? keys[2] : keys[keys.length - 1];
        }
    }

    let stats = {
        total: 0,
        group1: { total: 0, depts: {} },
        group2: { total: 0, depts: {} },
        other: 0
    };

    // Initialize departments
    Object.keys(GROUP_1.departments).forEach(d => stats.group1.depts[d] = 0);
    Object.keys(GROUP_2.departments).forEach(d => stats.group2.depts[d] = 0);

    data.forEach(row => {
        let rawDept = row[deptCol];
        if (!rawDept) return; // Skip empty rows
        stats.total++;

        const normDept = normalizeString(rawDept);
        let matched = false;

        // Check Group 1
        for (const [deptName, keywords] of Object.entries(GROUP_1.departments)) {
            if (keywords.some(kw => normDept.includes(normalizeString(kw)))) {
                stats.group1.depts[deptName]++;
                stats.group1.total++;
                matched = true;
                break;
            }
        }

        if (matched) return;

        // Check Group 2
        for (const [deptName, keywords] of Object.entries(GROUP_2.departments)) {
            if (keywords.some(kw => normDept.includes(normalizeString(kw)))) {
                stats.group2.depts[deptName]++;
                stats.group2.total++;
                matched = true;
                break;
            }
        }

        if (!matched) {
            stats.other++;
        }
    });

    updateUI(stats);
}

function updateUI(stats) {
    // Hide upload, show dashboard
    uploadSection.classList.remove('active-section');
    dashboardSection.classList.remove('hidden');

    // Animate numbers
    animateValue("total-staff", 0, stats.total, 1000);
    animateValue("total-group1", 0, stats.group1.total, 1000);
    animateValue("total-group2", 0, stats.group2.total, 1000);
    animateValue("total-other", 0, stats.other, 1000);

    // Render Charts
    renderCharts(stats);

    // Generate AI Report
    generateAIReport(stats);
}

function animateValue(id, start, end, duration) {
    if (start === end) {
        document.getElementById(id).innerHTML = end;
        return;
    }
    let range = end - start;
    let current = start;
    let increment = end > start ? 1 : -1;
    // Calculate appropriate step time based on range
    let stepTime = Math.abs(Math.floor(duration / range));
    if (stepTime < 5) stepTime = 5; // minimum interval

    let obj = document.getElementById(id);
    let timer = setInterval(function() {
        current += increment;
        obj.innerHTML = current;
        if (current == end) {
            clearInterval(timer);
        }
    }, stepTime);
}

function renderCharts(stats) {
    Chart.defaults.color = document.body.getAttribute('data-theme') === 'light' ? '#64748b' : '#94a3b8';
    Chart.defaults.font.family = "'Inter', sans-serif";

    // 1. Ratio Chart (Doughnut)
    const ctxRatio = document.getElementById('ratioChart').getContext('2d');
    ratioChartInst = new Chart(ctxRatio, {
        type: 'doughnut',
        data: {
            labels: ['PT Kinh Doanh & Sản Phẩm', 'Kỹ Thuật & Vận Hành', 'Khác'],
            datasets: [{
                data: [stats.group1.total, stats.group2.total, stats.other],
                backgroundColor: ['#8b5cf6', '#3b82f6', '#f59e0b'],
                borderWidth: 0,
                hoverOffset: 10
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom' }
            },
            cutout: '70%'
        }
    });

    // 2. Group 1 Details (Bar)
    const g1Labels = Object.keys(stats.group1.depts);
    const g1Data = Object.values(stats.group1.depts);
    const ctxG1 = document.getElementById('group1Chart').getContext('2d');
    group1ChartInst = new Chart(ctxG1, {
        type: 'bar',
        data: {
            labels: g1Labels,
            datasets: [{
                label: 'Số nhân sự',
                data: g1Data,
                backgroundColor: 'rgba(139, 92, 246, 0.7)',
                borderColor: '#8b5cf6',
                borderWidth: 1,
                borderRadius: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: { y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' } }, x: { grid: { display: false } } },
            plugins: { legend: { display: false } }
        }
    });

    // 3. Group 2 Details (Bar)
    const g2Labels = Object.keys(stats.group2.depts);
    const g2Data = Object.values(stats.group2.depts);
    const ctxG2 = document.getElementById('group2Chart').getContext('2d');
    group2ChartInst = new Chart(ctxG2, {
        type: 'bar',
        data: {
            labels: g2Labels,
            datasets: [{
                label: 'Số nhân sự',
                data: g2Data,
                backgroundColor: 'rgba(59, 130, 246, 0.7)',
                borderColor: '#3b82f6',
                borderWidth: 1,
                borderRadius: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: { y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' } }, x: { grid: { display: false } } },
            plugins: { legend: { display: false } }
        }
    });
}

function generateAIReport(stats) {
    const reportEl = document.getElementById('ai-report-content');
    reportEl.innerHTML = ''; // Clear loading
    
    if (stats.total === 0) {
        reportEl.innerHTML = '<p>Không có dữ liệu nhân sự để phân tích.</p>';
        return;
    }

    let reportHTML = '';

    // Calculate ratios
    const g1Ratio = (stats.group1.total / stats.total) * 100;
    const g2Ratio = (stats.group2.total / stats.total) * 100;
    
    reportHTML += `<div class="ai-insight">
        <strong>Tổng quan chung:</strong> Hệ thống đã phân tích tổng cộng ${stats.total} nhân sự. Trong đó, khối Kinh doanh & Sản phẩm chiếm ${g1Ratio.toFixed(1)}%, khối Kỹ thuật & Vận hành chiếm ${g2Ratio.toFixed(1)}%.
    </div>`;

    // Heuristics Check 1: Balance between Dev and QC
    const devs = stats.group2.depts["Phát triển Backend & Frontend"];
    const qcs = stats.group2.depts["QC (Quality Control)"];
    
    if (devs > 0 && qcs > 0) {
        const devQcRatio = devs / qcs;
        if (devQcRatio > 5) {
            reportHTML += `<div class="ai-insight ai-warning">
                <strong><i class="fa-solid fa-triangle-exclamation"></i> Rủi ro Chất lượng:</strong> Tỷ lệ Dev / QC đang là ${devQcRatio.toFixed(1)}:1 (Quá cao). Nguy cơ thắt cổ chai ở khâu kiểm thử. Đề xuất tuyển dụng thêm nhân sự QC để đảm bảo chất lượng sản phẩm.
            </div>`;
        } else if (devQcRatio < 2) {
             reportHTML += `<div class="ai-insight ai-warning">
                <strong><i class="fa-solid fa-circle-info"></i> Lưu ý nguồn lực:</strong> Tỷ lệ QC so với Dev khá cao (${devQcRatio.toFixed(1)}:1). Hãy đảm bảo hiệu suất làm việc của QC hoặc xem xét tự động hóa (Automation Test).
            </div>`;
        } else {
             reportHTML += `<div class="ai-insight ai-success">
                <strong><i class="fa-solid fa-check"></i> Cân bằng Kỹ thuật:</strong> Tỷ lệ Dev / QC (${devQcRatio.toFixed(1)}:1) đang ở mức lý tưởng, đảm bảo tốc độ phát triển và chất lượng.
            </div>`;
        }
    } else if (devs > 0 && qcs === 0) {
        reportHTML += `<div class="ai-insight ai-warning">
            <strong><i class="fa-solid fa-triangle-exclamation"></i> Báo động đỏ:</strong> Có đội ngũ phát triển nhưng không có nhân sự QC. Nguy cơ lỗi phần mềm khi release là cực kỳ cao. Cần tuyển QC gấp!
        </div>`;
    }

    // Heuristics Check 2: PO/PM vs Devs
    const pos = stats.group1.depts["Product Owner & Design"] + stats.group1.depts["Project Management"];
    if (pos > 0 && devs > 0) {
        const devPoRatio = devs / pos;
        if (devPoRatio > 8) {
            reportHTML += `<div class="ai-insight ai-warning">
                <strong><i class="fa-solid fa-circle-exclamation"></i> Thiếu hụt Quản lý dự án/PO:</strong> 1 PO/PM đang phải quản lý trung bình ${devPoRatio.toFixed(0)} Devs. Nguy cơ chậm trễ trong việc làm rõ requirement và backlog. Đề xuất bổ sung PO/BA.
            </div>`;
        }
    }

    // Heuristics Check 3: R&D focus
    const rnd = stats.group1.depts["R&D (Research & Development)"];
    if (rnd === 0 && stats.total > 50) {
        reportHTML += `<div class="ai-insight">
            <strong><i class="fa-solid fa-lightbulb"></i> Khuyến nghị dài hạn:</strong> Doanh nghiệp quy mô >50 nhân sự nhưng chưa ghi nhận nhân sự chuyên trách R&D. Để duy trì lợi thế cạnh tranh về công nghệ/sản phẩm mới, nên cân nhắc xây dựng lực lượng nòng cốt cho R&D.
        </div>`;
    }

    // Uncategorized check
    if (stats.other > stats.total * 0.3) {
        reportHTML += `<div class="ai-insight">
            <strong><i class="fa-solid fa-magnifying-glass"></i> Phân bổ chưa rõ ràng:</strong> Có tới ${stats.other} nhân sự (${((stats.other/stats.total)*100).toFixed(1)}%) thuộc các phòng ban khác chưa được đưa vào 2 khối cốt lõi. Hãy rà soát lại cơ cấu tổ chức nếu cần tối ưu chi phí vận hành.
        </div>`;
    }

    // Summary wrap up
    reportHTML += `<p><em>Báo cáo được tạo tự động bởi hệ thống dựa trên tập dữ liệu được cung cấp. Các đánh giá có tính tham khảo phục vụ ra quyết định chiến lược.</em></p>`;

    // Render with typing effect
    reportEl.innerHTML = reportHTML;
}
