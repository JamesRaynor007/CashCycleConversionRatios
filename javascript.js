// Cash Cycle Analyzer - Fixed load XLSX & ratios calculation with modal popup
const DAYS = 365;

const METRICS = {
    'Account Receivable': { key: 'ar' },
    'Total Credit Sales': { key: 'sales' },
    'Average Inventory': { key: 'inventory' },
    'Cost of Goods Sold': { key: 'cogs' },
    'Average Account Payable': { key: 'payable' }
};

const OBJECTIVES = {
    DSO: 45,
    DIO: 60,
    DPO: 75,
    CCC: 30
};

const fileInput = document.getElementById('xlsxFile');
const loadBtn = document.getElementById('loadBtn');
const downloadTemplateBtn = document.getElementById('downloadTemplateBtn');
const ratiosBtn = document.getElementById('ratiosBtn');
const statusEl = document.getElementById('status');
const metricsContainer = document.getElementById('metricsContainer');

// Modal elements
const modal = document.createElement('div');
modal.id = 'ratiosModal';
modal.className = 'modal';
modal.style.cssText = `
    display: none;
    position: fixed;
    z-index: 1000;
    left: 0;
    top: 0;
    width: 100%;
    height: 100%;
    background-color: rgba(0,0,0,0.5);
`;
const modalContent = document.createElement('div');
modalContent.className = 'modal-content';
modalContent.style.cssText = `
    background: rgba(255,255,255,0.95);
    margin: 5% auto;
    border-radius: 15px;
    padding: 25px;
    box-shadow: 0 20px 40px rgba(0,0,0,0.3);
    width: 90%;
    max-width: 1000px;
    max-height: 80vh;
    overflow-y: auto;
    backdrop-filter: blur(10px);
`;
const closeBtn = document.createElement('span');
closeBtn.className = 'close';
closeBtn.innerHTML = '&times;';
closeBtn.style.cssText = `
    color: #aaa;
    float: right;
    font-size: 28px;
    font-weight: bold;
    cursor: pointer;
`;
closeBtn.onclick = () => modal.style.display = 'none';
modalContent.appendChild(closeBtn);
modal.appendChild(modalContent);
document.body.appendChild(modal);

loadBtn.addEventListener('click', async () => {
    const file = fileInput.files[0];
    if (!file) {
        statusEl.textContent = 'Please select an XLSX file';
        return;
    }
    statusEl.textContent = 'Loading...';
    try {
        const data = await file.arrayBuffer();
        const wb = XLSX.read(data, { type: 'array' });
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        const headers = json[0];
        const rows = json.slice(1).map(row => {
            const obj = {};
            headers.forEach((h, i) => {
                obj[h] = row[i];
            });
            return obj;
        });

        const parsed = parseCashCycle(rows);
        renderMetrics(parsed);
        window.currentData = parsed;
        statusEl.textContent = 'Data loaded successfully! Check console for parsed data.';
    } catch(e) {
        console.error(e);
        statusEl.textContent = 'Error loading file.';
        alert('Parse error: ' + e.message);
    }
});

function parseCashCycle(rows) {
    const data = {
        current: {},
        last: {}
    };

    rows.forEach(row => {
        const accName = (row['Account'] || row['account'] || row['Account Name'] || '').toString().trim().toLowerCase();

        if (!accName) return;

        const amountRaw = row['Amount'] || row['amount'] || 0;
        const amount = parseFloat(amountRaw) || 0;

        // Simple string match
        if (accName.includes('receivable') && accName.includes('current')) {
            data.current.ar = amount;
        } else if (accName.includes('receivable') && accName.includes('last')) {
            data.last.ar = amount;
        } else if (accName.includes('credit sales') && accName.includes('current')) {
            data.current.sales = amount;
        } else if (accName.includes('credit sales') && accName.includes('last')) {
            data.last.sales = amount;
        } else if (accName.includes('inventory') && accName.includes('current')) {
            data.current.inventory = amount;
        } else if (accName.includes('inventory') && accName.includes('last')) {
            data.last.inventory = amount;
        } else if (accName.includes('goods sold') && accName.includes('current')) {
            data.current.cogs = amount;
        } else if (accName.includes('goods sold') && accName.includes('last')) {
            data.last.cogs = amount;
        } else if (accName.includes('account payable') && accName.includes('current')) {
            data.current.payable = amount;
        } else if (accName.includes('account payable') && accName.includes('last')) {
            data.last.payable = amount;
        }
    });

    // Defaults
    Object.keys(METRICS).forEach(key => {
        if (data.current[key] === undefined) data.current[key] = 0;
        if (data.last[key] === undefined) data.last[key] = 0;
    });

    console.log('Parsed Cash Cycle data:', data);
    return data;
}

function renderMetrics(data) {
    const results = computeRatios(data);

    metricsContainer.innerHTML = `
        <div class="kpi-grid">
            <h2>Cash Cycle Metrics</h2>
            <table>
                <thead>
                    <tr>
                        <th>Metric (days)</th>
                        <th>current</th>
                        <th>objective</th>
                        <th>% vs OBJ</th>
                        <th>Last</th>
                        <th>Evolution</th>
                    </tr>
                </thead>
                <tbody>
                    ${results.map(r => `
                        <tr>
                            <td>${r.name}</td>
                            <td>${isFinite(r.current) ? r.current.toFixed(1) : 'NaN'}</td>
                            <td>${r.obj}</td>
                            <td class="${r.percentObj >= 0 ? 'good' : 'bad'}">${isFinite(r.percentObj) ? r.percentObj.toFixed(1) + '%' : 'NaN'}</td>
                            <td>${isFinite(r.last) ? r.last.toFixed(1) : 'NaN'}</td>
                            <td class="${r.evoClass}">${isFinite(r.percentEvo) ? r.percentEvo.toFixed(1) + '%' : 'NaN'}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;
    statusEl.textContent = 'Loaded and ratios calculated!';
}

function computeRatios(data) {
    const { current, last } = data;

    const safeCurrentSales = current.sales || 1;
    const safeCurrentCogs = current.cogs || 1;
    const safeLastSales = last.sales || 1;
    const safeLastCogs = last.cogs || 1;

    const currentDSO = (current.ar / safeCurrentSales) * DAYS;
    const currentDIO = (current.inventory / safeCurrentCogs) * DAYS;
    const currentDPO = (current.payable / safeCurrentCogs) * DAYS;
    const currentCCC = currentDSO + currentDIO - currentDPO;

    const lastDSO = (last.ar / safeLastSales) * DAYS;
    const lastDIO = (last.inventory / safeLastCogs) * DAYS;
    const lastDPO = (last.payable / safeLastCogs) * DAYS;
    const lastCCC = lastDSO + lastDIO - lastDPO;

    const results = [
        {
            name: 'DSO',
            current: currentDSO,
            last: lastDSO,
            obj: OBJECTIVES.DSO,
            percentObj: 100 * (1 - currentDSO / OBJECTIVES.DSO),
            percentEvo: lastDSO === 0 ? 0 : 100 * (currentDSO - lastDSO) / lastDSO,
            deltaClass: currentDSO <= OBJECTIVES.DSO ? 'good' : 'bad',
            evoClass: currentDSO >= lastDSO ? 'evo-bad' : 'evo-good'
        },
        {
            name: 'DIO',
            current: currentDIO,
            last: lastDIO,
            obj: OBJECTIVES.DIO,
            percentObj: 100 * (1 - currentDIO / OBJECTIVES.DIO),
            percentEvo: lastDIO === 0 ? 0 : 100 * (currentDIO - lastDIO) / lastDIO,
            deltaClass: currentDIO <= OBJECTIVES.DIO ? 'good' : 'bad',
            evoClass: currentDIO >= lastDIO ? 'evo-bad' : 'evo-good'
        },
        {
            name: 'DPO',
            current: currentDPO,
            last: lastDPO,
            obj: OBJECTIVES.DPO,
            percentObj: 100 * (currentDPO / OBJECTIVES.DPO),
            percentEvo: lastDPO === 0 ? 0 : 100 * (currentDPO - lastDPO) / lastDPO,
            deltaClass: currentDPO >= OBJECTIVES.DPO ? 'good' : 'bad',
            evoClass: currentDPO >= lastDPO ? 'evo-bad' : 'evo-good'
        },
        {
            name: 'CCC',
            current: currentCCC,
            last: lastCCC,
            obj: OBJECTIVES.CCC,
            percentObj: 100 * (1 - Math.abs(currentCCC) / OBJECTIVES.CCC),
            percentEvo: lastCCC === 0 ? 0 : 100 * (currentCCC - lastCCC) / Math.abs(lastCCC),
            deltaClass: Math.abs(currentCCC) <= OBJECTIVES.CCC ? 'good' : 'bad',
            evoClass: currentCCC >= lastCCC ? 'evo-bad' : 'evo-good'
        }
    ];

    return results;
}

// Template download with sample data
downloadTemplateBtn.addEventListener('click', () => {
    const wb = XLSX.utils.book_new();

    const templateData = [
        ['Account', 'Amount'],
        ['Account Receivable (current)', 10000],
        ['Total Credit Sales (current)', 100000],
        ['Average Inventory (current)', 20000],
        ['Cost of Goods Sold (current)', 80000],
        ['Average Account Payable (current)', 15000],
        [],
        ['Account Receivable (last)', 12000],
        ['Total Credit Sales (last)', 95000],
        ['Average Inventory (last)', 22000],
        ['Cost of Goods Sold (last)', 78000],
        ['Average Account Payable (last)', 14000]
    ];

    const ws = XLSX.utils.aoa_to_sheet(templateData);
    XLSX.utils.book_append_sheet(wb, ws, 'Cash Cycle');
    XLSX.writeFile(wb, 'cash-cycle-template.xlsx');
    statusEl.textContent = 'Template downloaded with sample data - modify numbers, load to see ratios!';
});

// Modal ratios popup
ratiosBtn.addEventListener('click', () => {
    if (!window.currentData) return alert('Load data first');
    const ratios = computeRatios(window.currentData);
    const modalTable = modalContent.querySelector('table') || document.createElement('table');
    modalTable.className = 'popup-table';
    modalTable.innerHTML = `
        <thead>
            <tr>
                <th style="text-align: center;">Metric</th>
                <th style="text-align: center;">Current</th>
                <th style="text-align: center;">Objective</th>
                <th style="text-align: center;">% to Obj</th>
            </tr>
        </thead>
        <tbody>
            ${ratios.map(r => `
                <tr>
                    <td style="text-align: center;">${r.name}</td>
                    <td style="text-align: center;">${isFinite(r.current) ? r.current.toFixed(1) : 'NaN'}</td>
                    <td style="text-align: center;">${r.obj}</td>
                    <td style="text-align: center;" class="${r.percentObj >= 0 ? 'good' : 'bad'}">${isFinite(r.percentObj) ? r.percentObj.toFixed(1) + '%' : 'NaN'}</td>
                </tr>
            `).join('')}
        </tbody>
    `;
    if (!modalContent.querySelector('h2')) {
        const title = document.createElement('h2');
        title.textContent = 'Detailed Ratios Analysis';
        title.style.cssText = 'text-align: center; color: #2c3e50; margin-bottom: 20px;';
        modalContent.insertBefore(title, closeBtn);
    }
    if (!modalContent.querySelector('table')) {
        modalContent.appendChild(modalTable);
    }
    modal.style.display = 'block';
});

// Close modal on outside click
window.onclick = (event) => {
    if (event.target === modal) modal.style.display = 'none';
}

// DOM ready
document.addEventListener('DOMContentLoaded', () => {
    statusEl.textContent = 'Ready - Download template, modify amounts, load XLSX for ratios';
    console.log('Cash Cycle Analyzer loaded');
});
