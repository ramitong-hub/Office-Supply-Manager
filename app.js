// === Global Variables & State ===
const LOCAL_STORAGE_KEY = 'office_supplies_data';
const STOCK_STORAGE_KEY = 'office_stock_data';
let inventoryData = []; // Requisition History
let stockData = [];     // Current Stock
let currentEditId = null;
let currentStockEditId = null;

// === DOM Elements ===
// Requisition Form
const form = document.getElementById('requisitionForm');
const tableBody = document.getElementById('tableBody');
const noDataMessage = document.getElementById('noDataMessage');
const searchInput = document.getElementById('searchInput');
const filterMonth = document.getElementById('filterMonth');
const filterYear = document.getElementById('filterYear');
const btnClear = document.getElementById('btnClear');
const btnExport = document.getElementById('btnExport');
const itemNameInput = document.getElementById('itemName');
const itemCodeInput = document.getElementById('itemCode');
const unitInput = document.getElementById('unit');
const stockHint = document.getElementById('stockHint');

// Stock Form
const stockForm = document.getElementById('stockForm');
const stockTableBody = document.getElementById('stockTableBody');
const btnStockClear = document.getElementById('btnStockClear');

// Dashboard Elements
const statTotalRequests = document.getElementById('statTotalRequests');
const statTotalQty = document.getElementById('statTotalQty');
const statTopDept = document.getElementById('statTopDept');
const statLowStock = document.getElementById('statLowStock');
const chartCanvas = document.getElementById('summaryChart');

// === Initialization ===
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('yearDisplay').textContent = new Date().getFullYear();
    document.getElementById('reqDate').valueAsDate = new Date(); // Set default date to today
    
    // Initialize Filter Options (Months & Last 5 Years)
    initFilterOptions();
    
    // Load Data
    loadData();
    
    // Render UI
    refreshUI();
});

// === Data Management Functions ===

function loadData() {
    // Load Requisitions
    const storedData = localStorage.getItem(LOCAL_STORAGE_KEY);
    inventoryData = storedData ? JSON.parse(storedData) : [];

    // Load Stock
    const storedStock = localStorage.getItem(STOCK_STORAGE_KEY);
    stockData = storedStock ? JSON.parse(storedStock) : [];
}

function saveData() {
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(inventoryData));
    localStorage.setItem(STOCK_STORAGE_KEY, JSON.stringify(stockData));
    refreshUI();
}

function refreshUI() {
    // 1. Filter Data based on inputs
    const filteredData = getFilteredData();
    
    // 2. Render Tables
    renderTable(filteredData);
    renderStockTable();
    updateDatalist();
    
    // 3. Update Dashboard Stats & Chart
    updateDashboard(filteredData);
}

function getFilteredData() {
    const searchTerm = searchInput.value.toLowerCase().trim();
    const selectedMonth = filterMonth.value;
    const selectedYear = filterYear.value;

    return inventoryData.filter(item => {
        const itemDate = new Date(item.date);
        
        // Search Filter
        const matchesSearch = 
            item.requester.toLowerCase().includes(searchTerm) ||
            item.department.toLowerCase().includes(searchTerm) ||
            item.itemName.toLowerCase().includes(searchTerm);

        // Date Filter
        const matchesMonth = selectedMonth === 'all' || (itemDate.getMonth() + 1).toString() === selectedMonth;
        const matchesYear = selectedYear === 'all' || itemDate.getFullYear().toString() === selectedYear;

        return matchesSearch && matchesMonth && matchesYear;
    }).sort((a, b) => new Date(b.date) - new Date(a.date)); // Sort new to old
}

// === Stock Management Logic ===

// Add/Edit Stock
stockForm.addEventListener('submit', (e) => {
    e.preventDefault();

    const name = document.getElementById('stockName').value.trim();
    const code = document.getElementById('stockCode').value.trim();
    const qty = parseFloat(document.getElementById('stockQty').value);
    const unit = document.getElementById('stockUnit').value;

    if (currentStockEditId) {
        // Edit existing stock
        const index = stockData.findIndex(s => s.id === currentStockEditId);
        if (index !== -1) {
            stockData[index] = { ...stockData[index], name, code, quantity: qty, unit };
            alert('แก้ไขข้อมูลสต็อกเรียบร้อย');
        }
    } else {
        // Add new stock
        // Check duplicate name
        const exists = stockData.find(s => s.name.toLowerCase() === name.toLowerCase());
        if (exists) {
            alert('ชื่อพัสดุนี้มีในระบบแล้ว กรุณาแก้ไขรายการเดิม');
            return;
        }

        const newStock = {
            id: 'stk_' + Date.now(),
            name,
            code,
            quantity: qty,
            unit,
            updated: new Date().toISOString()
        };
        stockData.push(newStock);
        alert('เพิ่มพัสดุเข้าสต็อกเรียบร้อย');
    }

    saveData();
    resetStockForm();
});

function resetStockForm() {
    stockForm.reset();
    currentStockEditId = null;
    document.getElementById('stockEditId').value = '';
    document.getElementById('btnStockSave').textContent = 'เพิ่ม/ปรับปรุงสต็อก';
}

btnStockClear.addEventListener('click', resetStockForm);

function renderStockTable() {
    stockTableBody.innerHTML = '';
    let lowStockCount = 0;

    // Sort by Name
    const sortedStock = [...stockData].sort((a, b) => a.name.localeCompare(b.name));

    sortedStock.forEach(item => {
        const tr = document.createElement('tr');
        if (item.quantity < 5) {
            tr.classList.add('low-stock');
            lowStockCount++;
        }

        tr.innerHTML = `
            <td>${escapeHtml(item.name)}</td>
            <td>${escapeHtml(item.code || '-')}</td>
            <td><strong>${item.quantity}</strong></td>
            <td>${item.unit}</td>
            <td>
                <button class="btn-icon" onclick="editStock('${item.id}')" title="แก้ไข">✏️</button>
                <button class="btn-icon" onclick="deleteStock('${item.id}')" title="ลบ" style="color:red;">🗑️</button>
            </td>
        `;
        stockTableBody.appendChild(tr);
    });

    statLowStock.textContent = lowStockCount;
}

window.editStock = function(id) {
    const item = stockData.find(s => s.id === id);
    if (!item) return;

    currentStockEditId = id;
    document.getElementById('stockName').value = item.name;
    document.getElementById('stockCode').value = item.code;
    document.getElementById('stockQty').value = item.quantity;
    document.getElementById('stockUnit').value = item.unit;
    document.getElementById('btnStockSave').textContent = 'บันทึกการแก้ไขสต็อก';
    
    // Scroll to stock form
    document.querySelector('.stock-section').scrollIntoView({ behavior: 'smooth' });
};

window.deleteStock = function(id) {
    if (confirm('คุณแน่ใจหรือไม่ที่จะลบรายการนี้ออกจากสต็อก? (ประวัติการเบิกจะไม่ถูกลบ)')) {
        stockData = stockData.filter(s => s.id !== id);
        saveData();
    }
};

function updateDatalist() {
    const datalist = document.getElementById('stockItemsList');
    datalist.innerHTML = '';
    stockData.forEach(item => {
        const option = document.createElement('option');
        option.value = item.name;
        datalist.appendChild(option);
    });
}

// Auto-fill Item Details and Check Stock when typing in Requisition Form
itemNameInput.addEventListener('input', function() {
    const val = this.value.trim().toLowerCase();
    const stockItem = stockData.find(s => s.name.toLowerCase() === val);
    
    if (stockItem) {
        itemCodeInput.value = stockItem.code;
        unitInput.value = stockItem.unit;
        stockHint.textContent = `มีในสต็อก: ${stockItem.quantity} ${stockItem.unit}`;
        
        if (stockItem.quantity === 0) {
            stockHint.style.color = 'red';
            stockHint.textContent += ' (สินค้าหมด)';
        } else {
            stockHint.style.color = 'var(--primary-color)';
        }
    } else {
        stockHint.textContent = '';
    }
});

// === Requisition Form Handling ===

form.addEventListener('submit', (e) => {
    e.preventDefault();

    const reqName = document.getElementById('requesterName').value.trim();
    const dept = document.getElementById('department').value;
    const date = document.getElementById('reqDate').value;
    const itemName = document.getElementById('itemName').value.trim();
    const itemCode = document.getElementById('itemCode').value.trim();
    const qty = parseFloat(document.getElementById('quantity').value);
    const unit = document.getElementById('unit').value;
    const note = document.getElementById('note').value.trim();

    // Check Stock Availability
    const stockIndex = stockData.findIndex(s => s.name.toLowerCase() === itemName.toLowerCase());
    
    // Validation logic for stock
    if (stockIndex !== -1) {
        // Item exists in stock
        if (currentEditId) {
            // Editing existing requisition: 
            // Note: Simplification - We don't adjust stock automatically on Edit History to avoid logic errors
            // Alert user that stock needs manual adjustment if necessary, OR logic becomes too complex for Vanilla JS single file without transaction logs.
            // Current decision: Allow edit, but don't change stock.
        } else {
            // New Requisition
            if (stockData[stockIndex].quantity < qty) {
                alert(`จำนวนของในสต็อกไม่พอ (มี: ${stockData[stockIndex].quantity})`);
                return;
            }
            // Deduct Stock
            stockData[stockIndex].quantity -= qty;
        }
    } else {
        // Item not in stock system - Warning but allow?
        // Let's allow it for flexibility ("Ad-hoc purchase")
    }

    const record = {
        id: currentEditId || Date.now().toString(),
        requester: reqName,
        department: dept,
        date: date,
        itemName: itemName,
        itemCode: itemCode,
        quantity: qty,
        unit: unit,
        note: note,
        timestamp: new Date().toISOString()
    };

    if (currentEditId) {
        const index = inventoryData.findIndex(item => item.id === currentEditId);
        if (index !== -1) inventoryData[index] = record;
        alert('แก้ไขข้อมูลการเบิกเรียบร้อย (สต็อกไม่ได้ถูกปรับปรุงอัตโนมัติ)');
    } else {
        inventoryData.push(record);
        alert('บันทึกการเบิกและตัดสต็อกเรียบร้อย');
    }

    saveData();
    resetForm();
});

function resetForm() {
    form.reset();
    document.getElementById('reqDate').valueAsDate = new Date();
    currentEditId = null;
    document.getElementById('editId').value = '';
    stockHint.textContent = '';
    
    const btnSave = document.getElementById('btnSave');
    btnSave.textContent = 'บันทึกการเบิก';
    btnSave.classList.remove('btn-success');
    btnSave.classList.add('btn-primary');
}

btnClear.addEventListener('click', resetForm);

// === Table Rendering ===

function renderTable(data) {
    tableBody.innerHTML = '';
    
    if (data.length === 0) {
        noDataMessage.style.display = 'block';
        return;
    }
    
    noDataMessage.style.display = 'none';

    data.forEach((item, index) => {
        const tr = document.createElement('tr');
        const formattedDate = new Date(item.date).toLocaleDateString('th-TH', {
            year: 'numeric', month: '2-digit', day: '2-digit'
        });

        tr.innerHTML = `
            <td>${index + 1}</td>
            <td>${formattedDate}</td>
            <td>${escapeHtml(item.requester)}</td>
            <td><span class="badge">${item.department}</span></td>
            <td>${escapeHtml(item.itemName)}</td>
            <td>${escapeHtml(item.itemCode || '-')}</td>
            <td>${item.quantity} ${item.unit}</td>
            <td>${escapeHtml(item.note || '-')}</td>
            <td>
                <button class="btn-icon" onclick="editRecord('${item.id}')" title="แก้ไข">✏️</button>
                <button class="btn-icon" onclick="deleteRecord('${item.id}')" title="ลบ" style="color:red;">🗑️</button>
            </td>
        `;
        tableBody.appendChild(tr);
    });
}

// === CRUD Operations triggered from HTML ===

window.editRecord = function(id) {
    const item = inventoryData.find(d => d.id === id);
    if (!item) return;

    currentEditId = id;
    document.getElementById('requesterName').value = item.requester;
    document.getElementById('department').value = item.department;
    document.getElementById('reqDate').value = item.date;
    document.getElementById('itemName').value = item.itemName;
    document.getElementById('itemCode').value = item.itemCode;
    document.getElementById('quantity').value = item.quantity;
    document.getElementById('unit').value = item.unit;
    document.getElementById('note').value = item.note;

    // Trigger stock check visually
    const event = new Event('input');
    itemNameInput.dispatchEvent(event);

    const btnSave = document.getElementById('btnSave');
    btnSave.textContent = 'บันทึกการแก้ไข';
    btnSave.classList.remove('btn-primary');
    btnSave.classList.add('btn-success');
    
    document.querySelector('.form-section').scrollIntoView({ behavior: 'smooth' });
};

window.deleteRecord = function(id) {
    if (confirm('คุณแน่ใจหรือไม่ที่จะลบรายการเบิกนี้? (จำนวนจะถูกคืนเข้าสต็อก)')) {
        const itemToDelete = inventoryData.find(item => item.id === id);
        
        if (itemToDelete) {
            // Return to Stock
            const stockIndex = stockData.findIndex(s => s.name.toLowerCase() === itemToDelete.itemName.toLowerCase());
            if (stockIndex !== -1) {
                stockData[stockIndex].quantity += itemToDelete.quantity;
            }

            inventoryData = inventoryData.filter(item => item.id !== id);
            saveData();
            alert('ลบรายการและคืนยอดเข้าสต็อกแล้ว');
        }
    }
};

// === Dashboard & Chart ===

function updateDashboard(data) {
    // 1. Stats Calculation
    const totalRequests = data.length;
    const totalQty = data.reduce((sum, item) => sum + item.quantity, 0);
    
    // Find Top Department
    const deptCount = {};
    data.forEach(item => {
        deptCount[item.department] = (deptCount[item.department] || 0) + 1;
    });
    const topDept = Object.keys(deptCount).reduce((a, b) => deptCount[a] > deptCount[b] ? a : b, '-');

    // Update DOM
    statTotalRequests.textContent = totalRequests.toLocaleString();
    statTotalQty.textContent = totalQty.toLocaleString();
    statTopDept.textContent = topDept;

    // 2. Prepare Data for Chart (Top 5 Items by Frequency)
    const itemFreq = {};
    data.forEach(item => {
        itemFreq[item.itemName] = (itemFreq[item.itemName] || 0) + item.quantity;
    });

    const sortedItems = Object.entries(itemFreq)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);

    drawChart(sortedItems);
}

function drawChart(dataItems) {
    const ctx = chartCanvas.getContext('2d');
    const width = chartCanvas.width;
    const height = chartCanvas.height;
    const padding = 40;
    const chartHeight = height - padding * 2;
    const chartWidth = width - padding * 2;

    // Clear Canvas
    ctx.clearRect(0, 0, width, height);
    
    if (dataItems.length === 0) {
        ctx.font = "16px Sarabun";
        ctx.fillStyle = "#64748b";
        ctx.textAlign = "center";
        ctx.fillText("ไม่มีข้อมูลสำหรับแสดงกราฟ", width / 2, height / 2);
        return;
    }

    const maxVal = Math.max(...dataItems.map(i => i[1]));
    const barWidth = (chartWidth / dataItems.length) - 20;

    dataItems.forEach((item, index) => {
        const [name, value] = item;
        const barHeight = (value / maxVal) * (chartHeight - 30);
        const x = padding + index * (chartWidth / dataItems.length) + 10;
        const y = height - padding - barHeight;

        // Bar
        ctx.fillStyle = '#3b82f6';
        ctx.fillRect(x, y, barWidth, barHeight);

        // Value Label
        ctx.fillStyle = '#1e293b';
        ctx.textAlign = 'center';
        ctx.font = 'bold 14px sans-serif';
        ctx.fillText(value, x + barWidth / 2, y - 5);

        // Name Label
        ctx.fillStyle = '#64748b';
        ctx.font = '12px Sarabun';
        let displayName = name.length > 15 ? name.substring(0, 12) + '...' : name;
        ctx.fillText(displayName, x + barWidth / 2, height - padding + 20);
    });
    
    ctx.beginPath();
    ctx.moveTo(padding, height - padding);
    ctx.lineTo(width - padding, height - padding);
    ctx.strokeStyle = '#cbd5e1';
    ctx.stroke();
}

// === Utilities ===

function initFilterOptions() {
    const thaiMonths = [
        "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
        "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
    ];
    thaiMonths.forEach((m, i) => {
        const opt = document.createElement('option');
        opt.value = (i + 1).toString();
        opt.textContent = m;
        filterMonth.appendChild(opt);
    });

    const currentYear = new Date().getFullYear();
    for (let i = 0; i < 5; i++) {
        const y = currentYear - i;
        const opt = document.createElement('option');
        opt.value = y.toString();
        opt.textContent = y + 543; // Thai Year
        filterYear.appendChild(opt);
    }
}

function escapeHtml(text) {
    if (!text) return '';
    return text.replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

// === Event Listeners for Filters ===
searchInput.addEventListener('input', refreshUI);
filterMonth.addEventListener('change', refreshUI);
filterYear.addEventListener('change', refreshUI);

// === Export CSV ===
btnExport.addEventListener('click', () => {
    if (inventoryData.length === 0) {
        alert("ไม่มีข้อมูลสำหรับส่งออก");
        return;
    }

    let csvContent = "\uFEFF"; 
    
    // Headers updated to match new groups
    const headers = ["ID", "ผู้เบิก", "กลุ่ม", "วันที่", "รายการ", "รหัส", "จำนวน", "หน่วย", "หมายเหตุ"];
    csvContent += headers.join(",") + "\n";

    getFilteredData().forEach(item => {
        const row = [
            item.id,
            `"${item.requester.replace(/"/g, '""')}"`, 
            item.department,
            item.date,
            `"${item.itemName.replace(/"/g, '""')}"`,
            `"${item.itemCode || ''}"`,
            item.quantity,
            item.unit,
            `"${item.note || ''}"`
        ];
        csvContent += row.join(",") + "\n";
    });

    const encodedUri = encodeURI("data:text/csv;charset=utf-8," + csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `office_supplies_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
});