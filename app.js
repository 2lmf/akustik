const MATERIALS_DB = [
    { name: 'Armirani beton', rho: 2500, rw_lab: 0, dlw: 0, cat: 'masivni' },
    { name: 'Puna opeka', rho: 1800, rw_lab: 0, dlw: 0, cat: 'masivni' },
    { name: 'Porotherm 25 S', rho: 800, rw_lab: 52, dlw: 0, cat: 'masivni' },
    { name: 'Porotherm 20 AKU', rho: 1200, rw_lab: 56, dlw: 0, cat: 'masivni' },
    { name: 'Ytong 20', rho: 550, rw_lab: 44, dlw: 0, cat: 'masivni' },
    { name: 'EPS-T (zvučni)', rho: 15, rw_lab: 0, dlw: 28, cat: 'izolacija' },
    { name: 'EPS 100', rho: 20, rw_lab: 0, dlw: 0, cat: 'izolacija' },
    { name: 'XPS 300 kPa', rho: 33, rw_lab: 0, dlw: 0, cat: 'izolacija' },
    { name: 'XPS 500 kPa', rho: 42, rw_lab: 0, dlw: 0, cat: 'izolacija' },
    { name: 'XPS 700 kPa', rho: 48, rw_lab: 0, dlw: 0, cat: 'izolacija' },
    { name: 'Kamena vuna Floor', rho: 100, rw_lab: 0, dlw: 32, cat: 'izolacija' },
    { name: 'Čepasta ploča H30', rho: 20, rw_lab: 0, dlw: 0, cat: 'sistem' },
    { name: 'Cementni estrih', rho: 2200, rw_lab: 0, dlw: 0, cat: 'masivni' },
    { name: 'Anhidritni estrih', rho: 2400, rw_lab: 0, dlw: 0, cat: 'masivni' },
    { name: 'Keramika', rho: 2300, rw_lab: 0, dlw: 0, cat: 'obloga' },
    { name: 'Parket 14mm', rho: 700, rw_lab: 0, dlw: 5, cat: 'obloga' },
];

const PRESETS = {
    stambena: { rw: 52, lnw: 55 },
    skola: { rw: 47, lnw: 53 },
    bolnica: { rw: 55, lnw: 48 },
    ured: { rw: 42, lnw: 60 }
};

let layers = [
    { materialName: 'Armirani beton', thickness: 20 },
    { materialName: 'EPS-T (zvučni)', thickness: 3 },
    { materialName: 'Cementni estrih', thickness: 6 }
];

// --- GLOBALNE POSTAVKE I PODACI ---
let currentTab = 'akustika';
let settings = {
    armaturaRatio: 90, // kg/m3
    oplataLimit: 3.0    // m2
};

const DESCRIPTIONS_DB = [
    { title: 'Betonski zid C25/30', text: 'Dobava, doprema i ugradnja betona C25/30 u vertikalnu oplatu zida. U stavku uključeno vibriranje i njegovanje betona. Izračun po m3.' },
    { title: 'Oplata zida', text: 'Izrada, montaža i demontaža dvostrane glatke oplate zidova visine do 4m. U cijenu uključeno podupiranje i premazivanje oplatnim uljem. Izračun po m2.' },
    { title: 'Armatura B500B', text: 'Dobava, sječenje, savijanje i postavljanje armature B500B prema planovima armature. Izračun po kg.' }
];

let troskovnikElements = [
    { id: 1, name: 'Zid Z1', l: 10, h: 2.8, d: 20, description: DESCRIPTIONS_DB[0].text, openings: [] }
];

let currentElementId = null;

// --- INICIJALIZACIJA ---
function init() {
    loadSettings();
    renderLayers();
    renderTroskovnik();
    calculateAll();
}

function switchTab(tabId) {
    currentTab = tabId;
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));

    document.getElementById(`tab-${tabId}`).classList.add('active');
    document.querySelector(`.tab-btn[onclick*="${tabId}"]`).classList.add('active');

    // Update subtitle
    const subtitles = {
        akustika: 'Iskaznica Akustičkih Svojstava (NN 71/2025)',
        troskovnik: 'Građevinski Troškovnik (Beton / Oplata)',
        settings: 'Postavke sustava i izračuna'
    };
    document.getElementById('app-subtitle').innerText = subtitles[tabId];
}

function loadSettings() {
    const saved = localStorage.getItem('arch_toolkit_settings');
    if (saved) {
        settings = JSON.parse(saved);
        document.getElementById('setting-armatura').value = settings.armaturaRatio;
        document.getElementById('setting-oplata-limit').value = settings.oplataLimit;
    }
}

function saveSettings() {
    settings.armaturaRatio = parseFloat(document.getElementById('setting-armatura').value) || 90;
    settings.oplataLimit = parseFloat(document.getElementById('setting-oplata-limit').value) || 3;
    localStorage.setItem('arch_toolkit_settings', JSON.stringify(settings));
    calculateTroskovnik();
}

function resetSettings() {
    if (confirm("Vratiti postavke na zadane vrijednosti?")) {
        settings = { armaturaRatio: 90, oplataLimit: 3.0 };
        document.getElementById('setting-armatura').value = 90;
        document.getElementById('setting-oplata-limit').value = 3.0;
        saveSettings();
    }
}

function applyPreset() {
    const cat = document.getElementById('building-category').value;
    if (cat !== 'custom' && PRESETS[cat]) {
        document.getElementById('req-rw').value = PRESETS[cat].rw;
        document.getElementById('req-lnw').value = PRESETS[cat].lnw;
        calculateAll();
    }
}

function renderLayers() {
    const tbody = document.getElementById('layers-body');
    tbody.innerHTML = '';

    layers.forEach((layer, index) => {
        const tr = document.createElement('tr');

        const mat = findMaterial(layer.materialName);
        const mass = (mat.rho * layer.thickness / 100).toFixed(1);

        tr.innerHTML = `
            <td>${index + 1}</td>
            <td>
                <select onchange="updateLayerMaterial(${index}, this.value)">
                    ${MATERIALS_DB.map(m => `<option value="${m.name}" ${m.name === layer.materialName ? 'selected' : ''}>${m.name}</option>`).join('')}
                </select>
            </td>
            <td>
                <input type="number" value="${layer.thickness}" oninput="updateLayerThickness(${index}, this.value)">
            </td>
            <td style="color: var(--text-secondary)">${mass}</td>
            <td>
                <button class="btn btn-outline" style="padding: 0.25rem 0.5rem" onclick="removeLayer(${index})">×</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function findMaterial(name) {
    return MATERIALS_DB.find(m => m.name === name);
}

function addLayer() {
    layers.push({ materialName: 'Cementni estrih', thickness: 5 });
    renderLayers();
    calculateAll();
}

function removeLayer(index) {
    layers.splice(index, 1);
    renderLayers();
    calculateAll();
}

function updateLayerMaterial(index, name) {
    layers[index].materialName = name;
    renderLayers();
    calculateAll();
}

function updateLayerThickness(index, thickness) {
    layers[index].thickness = parseFloat(thickness) || 0;
    renderLayers();
    calculateAll();
}

function calculateAll() {
    let totalMass = 0;
    let baseMass = 0; // Mass of the main structural layer
    let totalDLw = 0;
    let maxRwLab = 0;

    const reqRw = parseFloat(document.getElementById('req-rw').value) || 0;
    const reqLnw = parseFloat(document.getElementById('req-lnw').value) || 0;

    layers.forEach(layer => {
        const mat = findMaterial(layer.materialName);
        const mass = (mat.rho * layer.thickness / 100);
        totalMass += mass;

        if (mat.cat === 'masivni') {
            baseMass = Math.max(baseMass, mass);

            // Lab Rw calculation
            let rw = 0;
            if (mat.rw_lab > 0) {
                rw = mat.rw_lab;
            } else {
                // ISO 12354 reduction for massive single leaf
                rw = 33.5 * Math.log10(mass) - 2;
            }
            maxRwLab = Math.max(maxRwLab, rw);
        }

        if (mat.dlw > 0) {
            totalDLw += mat.dlw;
        }
    });

    // AIRBORNE NOISE
    const rwInsitu = Math.round(maxRwLab - 3); // -3dB rule for flanking
    document.getElementById('val-total-mass').innerText = `${Math.round(totalMass)} kg/m²`;
    document.getElementById('val-rw').innerText = `${rwInsitu} dB`;

    const statusRw = document.getElementById('status-rw');
    if (rwInsitu >= reqRw) {
        statusRw.innerText = 'PROLAZI';
        statusRw.className = 'status-badge status-ok';
    } else {
        statusRw.innerText = 'NE PROLAZI';
        statusRw.className = 'status-badge status-fail';
    }

    // IMPACT NOISE
    const type = document.getElementById('const-type').value;
    const impactCard = document.getElementById('impact-card');

    if (type === 'floor') {
        impactCard.style.display = 'flex';
        // Simplified ISO 12354-2: Lnw = Ln,eq,0,w - dLw + K
        // Ln,eq,0,w approx for base slab
        const lnZero = 164 - 35 * Math.log10(baseMass || 1);
        const lnwInsitu = Math.round(lnZero - totalDLw + 2);

        document.getElementById('val-lnw').innerText = `${lnwInsitu} dB`;
        const statusLnw = document.getElementById('status-lnw');
        if (lnwInsitu <= reqLnw) {
            statusLnw.innerText = 'PROLAZI';
            statusLnw.className = 'status-badge status-ok';
        } else {
            statusLnw.innerText = 'NE PROLAZI';
            statusLnw.className = 'status-badge status-fail';
        }
    } else {
        impactCard.style.display = 'none';
    }
}

async function syncToDrive() {
    const email = document.getElementById('user-email').value;
    const buildingName = document.querySelector('input[placeholder="Npr. Stambena zgrada A"]').value;

    if (!email) {
        alert("Molimo unesite email adresu kako biste dobili kopiju izračuna.");
        return;
    }

    const payload = {
        email: email,
        buildingName: buildingName || "Novi projekt",
        layers: layers,
        resultRw: document.getElementById('val-rw').innerText,
        resultLnw: document.getElementById('val-lnw')?.innerText || "N/A",
        reqRw: document.getElementById('req-rw').value,
        reqLnw: document.getElementById('req-lnw').value,
        totalMass: document.getElementById('val-total-mass').innerText
    };

    console.log("Sending data to Cloud...", payload);

    // URL koji si mi poslao
    const GAS_URL = "https://script.google.com/macros/s/AKfycbz9pogDtFfoa3riTtnEk6Be5cx3Xt7_y1iZ55mxpaAHmKMgJ9SiDf8K-4jvzq_zh506TQ/exec";

    if (GAS_URL === "TVOJ_GOOGLE_APPS_SCRIPT_URL_OVDJE") {
        alert("Podaci su spremni, ali još niste postavili GAS URL u app.js! Pogledajte upute u walkthrough.md.");
        return;
    }

    try {
        const response = await fetch(GAS_URL, {
            method: 'POST',
            mode: 'no-cors', // GAS zahtijeva no-cors za jednostavne POST-ove
            body: JSON.stringify(payload)
        });
        // Maknut alert da ne smeta pri downloadu
        console.log("Sync uspješan!");
    } catch (err) {
        console.error("Sync greška: " + err.message);
    }
}

async function downloadExcelReport() {
    const buildingName = document.querySelector('input[placeholder="Npr. Stambena zgrada A"]').value || "Novi Projekt";

    // Provjera da li je knjižnica učitana iz index.html
    if (typeof ExcelJS === 'undefined') {
        alert("Sustav za generiranje Excela (ExcelJS) se nije ispravno učitao. Provjerite internet vezu.");
        return;
    }

    try {
        // AUTOMATSKI SYNC PRIJE DOWNLOADA
        await syncToDrive();

        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Iskaznica Akustike Web';
        workbook.created = new Date();

        // 1. DATA SHEET: BAZA_MATERIJALA
        const dataSheet = workbook.addWorksheet('BAZA_MATERIJALA');
        dataSheet.columns = [
            { header: 'Kategorija', key: 'cat', width: 20 },
            { header: 'Materijal', key: 'name', width: 35 },
            { header: 'Gustoća (kg/m3)', key: 'rho', width: 15 },
            { header: 'Rw_lab (dB)', key: 'rw_lab', width: 15 },
            { header: 'dLw (dB)', key: 'dlw', width: 10 },
            { header: 'Napomena', key: 'note', width: 30 },
        ];

        dataSheet.addRows(MATERIALS_DB.map(m => ({
            cat: m.cat.toUpperCase(),
            name: m.name,
            rho: m.rho,
            rw_lab: m.rw_lab,
            dlw: m.dlw,
            note: m.cat === 'masivni' ? (m.rw_lab > 0 ? 'Specijalni' : 'Računa se po masi') : ''
        })));

        // 2. PRORACUN SHEET
        const calcSheet = workbook.addWorksheet('PRORACUN');

        calcSheet.mergeCells('A1:G1');
        const titleCell = calcSheet.getCell('A1');
        titleCell.value = 'UNIVERZALNI AKUSTIČKI KALKULATOR (NN 71/2025)';
        titleCell.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
        titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF217346' } };
        titleCell.alignment = { horizontal: 'center' };

        const startRow = 3;
        calcSheet.getRow(startRow).values = ['Sloj br.', 'Materijal (Odaberi s liste)', 'Debljina (cm)', 'Masa (kg/m2)', 'Rw_lab', 'dLw', 'Napomena'];
        calcSheet.getRow(startRow).font = { bold: true };
        calcSheet.columns = [
            { width: 8 }, { width: 35 }, { width: 15 }, { width: 15 }, { width: 12 }, { width: 10 }, { width: 30 }
        ];

        let totalMasaCalc = 0;
        let maxRwLabVal = 0;
        let dLwSum = 0;
        let baseMassExcel = 0;

        for (let i = 1; i <= 8; i++) {
            const rowNum = startRow + i;
            const row = calcSheet.getRow(rowNum);
            const inputLayer = layers && layers[i - 1] ? layers[i - 1] : null;

            row.getCell(1).value = i;
            const valCell = row.getCell(2);
            if (inputLayer) {
                valCell.value = inputLayer.materialName;
                row.getCell(3).value = inputLayer.thickness;

                // Ručni proračun prema ISO 12354 (isto kao u aplikaciji)
                const mat = MATERIALS_DB.find(m => m.name === inputLayer.materialName);
                if (mat) {
                    const mass = mat.rho * (inputLayer.thickness / 100);
                    totalMasaCalc += mass;
                    if (mat.dlw > 0) dLwSum += mat.dlw;

                    if (mat.cat === 'masivni') {
                        baseMassExcel = Math.max(baseMassExcel, mass);
                        let rw = mat.rw_lab > 0 ? mat.rw_lab : (33.5 * Math.log10(mass) - 2);
                        maxRwLabVal = Math.max(maxRwLabVal, rw);
                    }
                }
            }

            valCell.dataValidation = {
                type: 'list',
                allowBlank: true,
                formulae: [`=BAZA_MATERIJALA!$B$2:$B$${MATERIALS_DB.length + 1}`]
            };

            // Formule i vrijednosti (kombinirano)
            row.getCell(4).value = {
                formula: `IF(B${rowNum}="",0,VLOOKUP(B${rowNum},BAZA_MATERIJALA!$B$2:$D$${MATERIALS_DB.length + 1},2,FALSE)*(C${rowNum}/100))`,
                result: inputLayer ? (MATERIALS_DB.find(m => m.name === inputLayer.materialName)?.rho * (inputLayer.thickness / 100)) : 0
            };
            row.getCell(5).value = {
                formula: `IF(B${rowNum}="",0,IF(VLOOKUP(B${rowNum},BAZA_MATERIJALA!$B$2:$E$${MATERIALS_DB.length + 1},3,FALSE)>0,VLOOKUP(B${rowNum},BAZA_MATERIJALA!$B$2:$E$${MATERIALS_DB.length + 1},3,FALSE),33.5*LOG10(MAX(1,D${rowNum}))-2))`
            };
            row.getCell(6).value = {
                formula: `IF(B${rowNum}="",0,VLOOKUP(B${rowNum},BAZA_MATERIJALA!$B$2:$E$${MATERIALS_DB.length + 1},4,FALSE))`,
                result: inputLayer ? (MATERIALS_DB.find(m => m.name === inputLayer.materialName)?.dlw) : 0
            };
        }

        const resRow = startRow + 10;
        calcSheet.getCell(`C${resRow}`).value = 'UKUPNA MASA [kg/m2]:';
        calcSheet.getCell(`D${resRow}`).value = { formula: `SUM(D4:D11)`, result: totalMasaCalc };
        calcSheet.getCell(`D${resRow}`).font = { bold: true };

        const rwRow = resRow + 1;
        calcSheet.getCell(`C${rwRow}`).value = 'RW_LAB (Teoretski):';
        calcSheet.getCell(`D${rwRow}`).value = { formula: `MAX(E4:E11)`, result: maxRwLabVal };

        const rprimeRow = rwRow + 1;
        calcSheet.getCell(`C${rprimeRow}`).value = "R'W (TERENSKI -3dB):";
        calcSheet.getCell(`D${rprimeRow}`).value = { formula: `D${rwRow}-3`, result: maxRwLabVal - 3 };
        calcSheet.getCell(`D${rprimeRow}`).font = { bold: true, color: { argb: 'FF0000FF' } };

        const lprimeRow = rprimeRow + 1;
        calcSheet.getCell(`C${lprimeRow}`).value = "L'NW (UDARNA BUKA):";
        // ISO 12354-2 formula: (164 - 35*log10(m_base)) - dLw + 2
        // U Excelu je teže automatski naći m_base iz liste, pa koristimo sumu dLw i aproksimaciju mase ako nemamo baseMassExcel
        const lnZeroExcel = 164 - 35 * Math.log10(baseMassExcel || 1);
        const finalLnw = Math.round(lnZeroExcel - dLwSum + 2);

        calcSheet.getCell(`D${lprimeRow}`).value = {
            formula: `IF(D${resRow}<50,"N/A",ROUND(164-35*LOG10(MAX(1,D${resRow}))-SUM(F4:F11)+2,0))`,
            result: finalLnw
        };

        // 3. ISKAZNICA SHEET (PRILOG E)
        const iskaznica = workbook.addWorksheet('ISKAZNICA');
        iskaznica.getColumn(1).width = 5;
        iskaznica.getColumn(2).width = 40;
        iskaznica.getColumn(3).width = 25;
        iskaznica.getColumn(4).width = 20;
        iskaznica.getColumn(5).width = 20;
        iskaznica.getColumn(6).width = 15;

        // Header
        iskaznica.mergeCells('A1:F1');
        iskaznica.getCell('A1').value = 'PRILOG E';
        iskaznica.getCell('A1').alignment = { horizontal: 'center' };

        iskaznica.mergeCells('A2:F2');
        const mainTitle = iskaznica.getCell('A2');
        mainTitle.value = 'ISKAZNICA O AKUSTIČKIM SVOJSTVIMA ZGRADE';
        mainTitle.font = { bold: true, size: 14 };
        mainTitle.alignment = { horizontal: 'center' };

        // Metadata section
        let currRow = 4;
        const addSection = (label, value) => {
            iskaznica.getCell(`A${currRow}`).value = label;
            iskaznica.getCell(`A${currRow}`).font = { bold: true };
            iskaznica.mergeCells(`B${currRow}:F${currRow}`);
            iskaznica.getCell(`B${currRow}`).value = value || '';
            iskaznica.getCell(`B${currRow}`).border = { bottom: { style: 'thin' } };
            currRow++;
        };

        addSection('1. INVESTITOR', document.getElementById('prj-investitor').value);
        addSection('2. OZNAKA PROJEKTA', document.getElementById('prj-oznaka').value);
        currRow++;

        iskaznica.getCell(`A${currRow}`).value = '3. OPIS ZGRADE';
        iskaznica.getCell(`A${currRow}`).font = { bold: true };
        currRow++;

        const addSubSection = (label, value) => {
            iskaznica.getCell(`A${currRow}`).value = label;
            iskaznica.mergeCells(`B${currRow}:F${currRow}`);
            iskaznica.getCell(`B${currRow}`).value = value || '';
            iskaznica.getCell(`B${currRow}`).border = { bottom: { style: 'thin' } };
            currRow++;
        };

        addSubSection('Naziv zgrade ili dijela zgrade', document.getElementById('prj-naziv').value || document.getElementById('project-name').value);
        addSubSection('Vrsta zgrade', document.getElementById('prj-vrsta').value);
        addSubSection('Namjena zgrade', document.getElementById('prj-namjena').value);
        addSubSection('Lokacija zgrade', document.getElementById('prj-lokacija').value);
        addSubSection('Mjesto, mjesec i godina', document.getElementById('prj-datum').value);
        currRow += 2;

        // Section 4
        iskaznica.getCell(`A${currRow}`).value = '4. ZAŠTITA OD VANJSKIH IZVORA BUKE';
        iskaznica.getCell(`A${currRow}`).font = { bold: true };
        currRow++;
        const temelj = document.getElementById('prj-temelj').value;
        iskaznica.mergeCells(`A${currRow}:F${currRow}`);
        iskaznica.getCell(`A${currRow}`).value = `Proračun se temelji na: ${temelj === 'a' ? '(a) mjerenju inicijalne buke' :
            temelj === 'b' ? '(b) proračunskoj procjeni stvarnog opterećenja' :
                '(c) najvišoj dopuštenoj razini vanjske buke'
            }`;
        currRow += 2;

        // Section 5 & 6 (Dynamic Tables)
        const addTableHeaders = (title) => {
            iskaznica.mergeCells(`A${currRow}:F${currRow}`);
            iskaznica.getCell(`A${currRow}`).value = title;
            iskaznica.getCell(`A${currRow}`).font = { bold: true };
            currRow++;

            const headRow = iskaznica.getRow(currRow);
            headRow.values = ['', 'Građevni dio ili pregrada', 'Oznaka vel.', 'Traženo', 'Projektirano', 'Izmjereno'];
            headRow.font = { bold: true };
            headRow.eachCell(c => c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } });
            currRow++;
        };

        const isWall = document.getElementById('const-type').value === 'wall';

        // Table 5 (External)
        addTableHeaders('5. ZVUČNA IZOLACIJA VANJSKIH GRAĐEVNIH DIJELOVA');
        if (isWall) {
            const r = iskaznica.getRow(currRow);
            r.values = ['1', document.getElementById('prj-naziv').value || 'Vanjski zid', "R'w [dB]", document.getElementById('req-rw').value, { formula: 'PRORACUN!D15', result: maxRwLabVal - 3 }, ''];
            r.eachCell(c => c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } });
            currRow++;
        }
        currRow += 2;

        // Table 6 (Internal)
        addTableHeaders('6. ZVUČNA IZOLACIJA UNUTARNJIH GRAĐEVNIH DIJELOVA');
        if (!isWall) {
            // First row for Rw
            let r1 = iskaznica.getRow(currRow);
            r1.values = ['1', document.getElementById('prj-naziv').value || 'Međukatna konstrukcija', "R'w [dB]", document.getElementById('req-rw').value, { formula: 'PRORACUN!D15', result: maxRwLabVal - 3 }, ''];
            r1.eachCell(c => c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } });
            currRow++;
            // Second row for Lnw
            let r2 = iskaznica.getRow(currRow);
            const lnZeroExcel = 164 - 35 * Math.log10(baseMassExcel || 1);
            const finalLnw = Math.round(lnZeroExcel - dLwSum + 2);
            r2.values = ['2', document.getElementById('prj-naziv').value || 'Međukatna konstrukcija', "L'nw [dB]", document.getElementById('req-lnw').value, { formula: 'PRORACUN!D16', result: finalLnw }, ''];
            r2.eachCell(c => c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } });
            currRow++;
        }
        currRow += 2;

        // Sections 7, 8 (Placeholders)
        const addEmptySection = (title) => {
            iskaznica.getCell(`A${currRow}`).value = title;
            iskaznica.getCell(`A${currRow}`).font = { bold: true };
            currRow += 2;
        };
        addEmptySection('7. ZVUČNA IZOLACIJA OD BUKE SERVISNE OPREME');
        addEmptySection('8. PROSTORNA AKUSTIKA I ZAŠTITA OD ODJEKA');

        // Section 9
        iskaznica.getCell(`A${currRow}`).value = '9. ODGOVORNOST ZA PROJEKTIRANE VRIJEDNOSTI';
        iskaznica.getCell(`A${currRow}`).font = { bold: true };
        currRow++;
        const signs = ['Projektant arhitektonskog dijela', 'Projektant građevinskog dijela', 'Projektant strojarskog dijela', 'Glavni projektant zgrade'];
        signs.forEach(s => {
            iskaznica.getCell(`A${currRow}`).value = s;
            iskaznica.mergeCells(`B${currRow}:D${currRow}`);
            iskaznica.getCell(`B${currRow}`).border = { bottom: { style: 'thin' } };
            iskaznica.getCell(`E${currRow}`).value = '(Potpis)';
            currRow++;
        });

        iskaznica.columns = [{ width: 35 }, { width: 35 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 }];

        // GENERIRANJE DATOTEKE (U MEMORIJI PREGLEDNIKA)
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        const downloadUrl = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = downloadUrl;
        a.download = `Akustika_Iskaznica_${buildingName.replace(/\s+/g, '_')}.xlsx`;
        document.body.appendChild(a);
        a.click();

        window.URL.revokeObjectURL(downloadUrl);
        a.remove();

    } catch (err) {
        alert("Greška pri kreiranju Excel datoteke u pregledniku: " + err.message);
    }
}



function downloadJSON() {
    const data = {
        projectName: document.getElementById('project-name').value,
        category: document.getElementById('building-category').value,
        type: document.getElementById('const-type').value,
        reqRw: document.getElementById('req-rw').value,
        reqLnw: document.getElementById('req-lnw').value,
        email: document.getElementById('user-email').value,
        // Prilog E metadata
        investitor: document.getElementById('prj-investitor').value,
        oznaka: document.getElementById('prj-oznaka').value,
        naziv: document.getElementById('prj-naziv').value,
        vrsta: document.getElementById('prj-vrsta').value,
        namjena: document.getElementById('prj-namjena').value,
        lokacija: document.getElementById('prj-lokacija').value,
        datum: document.getElementById('prj-datum').value,
        temelj: document.getElementById('prj-temelj').value,
        layers: layers
    };

    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Akustika_${data.projectName.replace(/\s+/g, '_') || 'Projekt'}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function triggerUploadJSON() {
    document.getElementById('json-upload-input').click();
}

function uploadJSON(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = JSON.parse(e.target.result);
            if (data.layers) layers = data.layers;
            if (data.projectName) document.getElementById('project-name').value = data.projectName;
            if (data.category) document.getElementById('building-category').value = data.category;
            if (data.type) document.getElementById('const-type').value = data.type;
            if (data.reqRw) document.getElementById('req-rw').value = data.reqRw;
            if (data.reqLnw) document.getElementById('req-lnw').value = data.reqLnw;
            if (data.email) document.getElementById('user-email').value = data.email;

            // Prilog E metadata
            if (data.investitor) document.getElementById('prj-investitor').value = data.investitor;
            if (data.oznaka) document.getElementById('prj-oznaka').value = data.oznaka;
            if (data.naziv) document.getElementById('prj-naziv').value = data.naziv;
            if (data.vrsta) document.getElementById('prj-vrsta').value = data.vrsta;
            if (data.namjena) document.getElementById('prj-namjena').value = data.namjena;
            if (data.lokacija) document.getElementById('prj-lokacija').value = data.lokacija;
            if (data.datum) document.getElementById('prj-datum').value = data.datum;
            if (data.temelj) document.getElementById('prj-temelj').value = data.temelj;

            renderLayers();
            calculateAll();
            alert("Projekt uspješno učitan!");
        } catch (err) {
            alert("Greška pri učitavanju JSON datoteke.");
        }
    };
    reader.readAsText(file);
}

// --- TROŠKOVNIK LOGIKA ---

function renderTroskovnik() {
    const tbody = document.getElementById('troskovnik-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    troskovnikElements.forEach((el, index) => {
        const totalOpeningArea = el.openings.reduce((sum, op) => sum + (op.w * op.h * op.count), 0);
        const openingPercent = el.l * el.h > 0 ? ((totalOpeningArea / (el.l * el.h)) * 100).toFixed(1) : 0;

        const tr = document.createElement('tr');
        tr.className = 'troskovnik-row';
        tr.innerHTML = `
            <td colspan="6">
                <div class="element-header">
                    <input type="text" class="input-name" value="${el.name}" oninput="updateElementName(${index}, this.value)">
                    <div class="element-dims">
                        L: <input type="number" value="${el.l}" step="0.1" oninput="updateElementDim(${index}, 'l', this.value)"> m | 
                        H: <input type="number" value="${el.h}" step="0.1" oninput="updateElementDim(${index}, 'h', this.value)"> m | 
                        d: <input type="number" value="${el.d}" oninput="updateElementDim(${index}, 'd', this.value)"> cm
                    </div>
                    <div class="element-actions">
                        <button class="btn btn-outline small" onclick="editOpenings(${el.id})">🔲 Otvori (${el.openings.length})</button>
                        <button class="btn btn-outline small danger" onclick="removeElement(${index})">×</button>
                    </div>
                </div>
                <div class="element-description">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.5rem;">
                        <span style="font-size: 0.8rem; font-weight: bold; color: var(--text-secondary);">Opis stavke:</span>
                        <select class="preset-select" onchange="applyDescriptionPreset(${index}, this.value)">
                            <option value="">-- Odaberi predložak --</option>
                            ${DESCRIPTIONS_DB.map(d => `<option value="${d.text}">${d.title}</option>`).join('')}
                        </select>
                    </div>
                    <textarea oninput="updateElementDescription(${index}, this.value)">${el.description || ''}</textarea>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });
    calculateTroskovnik();
}

function updateElementDescription(index, val) {
    troskovnikElements[index].description = val;
}

function applyDescriptionPreset(index, val) {
    if (!val) return;
    troskovnikElements[index].description = val;
    renderTroskovnik();
}

function addElement() {
    const newId = troskovnikElements.length > 0 ? Math.max(...troskovnikElements.map(e => e.id)) + 1 : 1;
    troskovnikElements.push({
        id: newId,
        name: `Element ${newId}`,
        l: 5,
        h: 2.8,
        d: 20,
        description: DESCRIPTIONS_DB[0].text,
        openings: []
    });
    renderTroskovnik();
}

function removeElement(index) {
    if (troskovnikElements[index].id === currentElementId) {
        document.getElementById('openings-editor').style.display = 'none';
        currentElementId = null;
    }
    troskovnikElements.splice(index, 1);
    renderTroskovnik();
}

function updateElementName(index, val) {
    troskovnikElements[index].name = val;
}

function updateElementDim(index, field, val) {
    troskovnikElements[index][field] = parseFloat(val) || 0;
    renderTroskovnik();
}

function editOpenings(id) {
    currentElementId = id;
    const el = troskovnikElements.find(e => e.id === id);
    document.getElementById('current-element-name').innerText = el.name;
    document.getElementById('openings-editor').style.display = 'block';
    renderOpenings();
}

function renderOpenings() {
    const el = troskovnikElements.find(e => e.id === currentElementId);
    const tbody = document.getElementById('openings-body');
    tbody.innerHTML = '';

    el.openings.forEach((op, index) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" value="${op.name}" oninput="updateOpening(${index}, 'name', this.value)"></td>
            <td>
                <select onchange="updateOpening(${index}, 'type', this.value)">
                    <option value="prozor" ${op.type === 'prozor' ? 'selected' : ''}>Prozor (4 str.)</option>
                    <option value="vrata" ${op.type === 'vrata' ? 'selected' : ''}>Vrata (3 str.)</option>
                </select>
            </td>
            <td><input type="number" value="${op.w}" step="0.1" oninput="updateOpening(${index}, 'w', this.value)"></td>
            <td><input type="number" value="${op.h}" step="0.1" oninput="updateOpening(${index}, 'h', this.value)"></td>
            <td><input type="number" value="${op.count}" oninput="updateOpening(${index}, 'count', this.value)"></td>
            <td>
                <button class="btn btn-outline" style="padding: 0.25rem 0.5rem" onclick="removeOpening(${index})">×</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function addOpening() {
    const el = troskovnikElements.find(e => e.id === currentElementId);
    el.openings.push({ name: 'Prozor', type: 'prozor', w: 1.2, h: 1.4, count: 1 });
    renderOpenings();
    renderTroskovnik();
}

function removeOpening(index) {
    const el = troskovnikElements.find(e => e.id === currentElementId);
    el.openings.splice(index, 1);
    renderOpenings();
    renderTroskovnik();
}

function updateOpening(index, field, val) {
    const el = troskovnikElements.find(e => e.id === currentElementId);
    el.openings[index][field] = field === 'name' ? val : (parseFloat(val) || 0);
    renderTroskovnik();
}

function calculateTroskovnik() {
    let totalBeton = 0;
    let totalOplata = 0;
    let totalSpalete = 0;

    troskovnikElements.forEach(el => {
        const grossArea = el.l * el.h;
        const grossVolume = grossArea * (el.d / 100);

        let netVolume = grossVolume;
        let netOplataArea = grossArea * 2; // Obostrano
        let spaleteArea = 0;

        el.openings.forEach(op => {
            const opArea = op.w * op.h * op.count;
            const opVolume = opArea * (el.d / 100);

            // Beton: uvijek se oduzima
            netVolume -= opVolume;

            // Oplata: oduzima se ako je pojedinačni otvor > limite
            if ((op.w * op.h) > settings.oplataLimit) {
                netOplataArea -= (opArea * 2);
            }

            // Špalete: oplata rubova otvora
            // Obujam (opseg) ovisi o tipu: prozor 4 strane, vrata 3 strane
            const perimeter = (op.type === 'vrata') ? (op.w + 2 * op.h) : (2 * op.w + 2 * op.h);
            spaleteArea += (perimeter * (el.d / 100)) * op.count;
        });

        totalBeton += netVolume;
        totalOplata += netOplataArea;
        totalSpalete += spaleteArea;
    });

    // Konačna oplata uključuje i špalete
    const finalOplata = totalOplata + totalSpalete;

    document.getElementById('val-total-beton').innerText = `${totalBeton.toFixed(2)} m³`;
    document.getElementById('val-total-oplata').innerText = `${finalOplata.toFixed(2)} m²`;
    document.getElementById('val-total-spalete').innerText = `${totalSpalete.toFixed(2)} m²`;
    document.getElementById('val-total-armatura').innerText = `${Math.round(totalBeton * settings.armaturaRatio)} kg`;
}

async function syncTroskovnik() {
    const email = document.getElementById('user-email').value;
    if (!email) {
        alert("Unesite email u tabu Akustika za sinkronizaciju.");
        switchTab('akustika');
        document.getElementById('user-email').focus();
        return;
    }

    const payload = {
        type: 'troskovnik',
        email: email,
        project: document.getElementById('project-name').value || "Novi Projekt",
        data: troskovnikElements,
        summary: {
            beton: document.getElementById('val-total-beton').innerText,
            oplata: document.getElementById('val-total-oplata').innerText,
            spalete: document.getElementById('val-total-spalete').innerText,
            armatura: document.getElementById('val-total-armatura').innerText
        }
    };

    const GAS_URL = "https://script.google.com/macros/s/AKfycbz9pogDtFfoa3riTtnEk6Be5cx3Xt7_y1iZ55mxpaAHmKMgJ9SiDf8K-4jvzq_zh506TQ/exec";

    try {
        await fetch(GAS_URL, { method: 'POST', mode: 'no-cors', body: JSON.stringify(payload) });
        alert("Troškovnik sinkroniziran s Google Sheets!");
    } catch (e) {
        alert("Greška pri sinkronizaciji: " + e.message);
    }
}

window.onload = init;

