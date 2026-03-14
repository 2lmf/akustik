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

function init() {
    renderLayers();
    calculateAll();
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
        alert("Zahtjev poslan! Ako je GAS ispravno konfiguriran, podaci su spremljeni i email je na putu.");
    } catch (err) {
        alert("Greška pri komunikaciji s Cloudom: " + err.message);
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

        const materials = [
            { cat: 'MASIVNI', name: 'Armirani beton', rho: 2500, rw_lab: 0, dlw: 0, note: 'Računa se po masi' },
            { cat: 'MASIVNI', name: 'Puna opeka', rho: 1800, rw_lab: 0, dlw: 0, note: 'Računa se po masi' },
            { cat: 'MASIVNI', name: 'Porotherm 25 S', rho: 800, rw_lab: 52, dlw: 0, note: 'Sa žbukom' },
            { cat: 'MASIVNI', name: 'Porotherm 20 AKU', rho: 1200, rw_lab: 56, dlw: 0, note: 'Specijalni AKU' },
            { cat: 'MASIVNI', name: 'Ytong 20', rho: 550, rw_lab: 44, dlw: 0, note: 'Lagani blok' },
            { cat: 'IZOLACIJA', name: 'EPS-T (zvučni)', rho: 15, rw_lab: 0, dlw: 28, note: 'Za udarnu buku' },
            { cat: 'IZOLACIJA', name: 'EPS 100', rho: 20, rw_lab: 0, dlw: 0, note: 'Podni EPS' },
            { cat: 'IZOLACIJA', name: 'XPS 300 kPa', rho: 33, rw_lab: 0, dlw: 0, note: 'Ekstrudirani' },
            { cat: 'IZOLACIJA', name: 'XPS 500 kPa', rho: 42, rw_lab: 0, dlw: 0, note: 'Vrlo jaki XPS' },
            { cat: 'IZOLACIJA', name: 'XPS 700 kPa', rho: 48, rw_lab: 0, dlw: 0, note: 'Ekstremni XPS' },
            { cat: 'IZOLACIJA', name: 'Kamena vuna Floor', rho: 100, rw_lab: 0, dlw: 32, note: 'Podna vuna' },
            { cat: 'SISTEM', name: 'Čepasta ploča H30', rho: 20, rw_lab: 0, dlw: 0, note: 'Podno grijanje' },
            { cat: 'SISTEM', name: 'Cementni estrih', rho: 2200, rw_lab: 0, dlw: 0, note: 'Mokri estrih' },
            { cat: 'SISTEM', name: 'Anhidritni estrih', rho: 2400, rw_lab: 0, dlw: 0, note: 'Tekući estrih' },
            { cat: 'OBLOGA', name: 'Keramika', rho: 2300, rw_lab: 0, dlw: 0, note: '' },
            { cat: 'OBLOGA', name: 'Parket 14mm', rho: 700, rw_lab: 0, dlw: 5, note: '' },
        ];

        dataSheet.addRows(materials);

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

        for (let i = 1; i <= 8; i++) {
            const rowNum = startRow + i;
            const row = calcSheet.getRow(rowNum);
            
            // Mapiramo lokalno spremljene slojeve u tablicu
            const inputLayer = layers && layers[i - 1] ? layers[i - 1] : null;

            row.getCell(1).value = i;
            
            const valCell = row.getCell(2);
            if (inputLayer) valCell.value = inputLayer.materialName;
            valCell.dataValidation = {
                type: 'list',
                allowBlank: true,
                formulae: [`=BAZA_MATERIJALA!$B$2:$B$${materials.length + 1}`]
            };

            if (inputLayer) row.getCell(3).value = inputLayer.thickness;

            // Formule
            row.getCell(4).value = { formula: `IF(B${rowNum}="",0,VLOOKUP(B${rowNum},BAZA_MATERIJALA!$B$2:$D$${materials.length + 1},2,FALSE)*(C${rowNum}/100))` };
            row.getCell(5).value = { formula: `IF(AND(B${rowNum}<>"",VLOOKUP(B${rowNum},BAZA_MATERIJALA!$B$2:$E$${materials.length + 1},3,FALSE)>0),VLOOKUP(B${rowNum},BAZA_MATERIJALA!$B$2:$E$${materials.length + 1},3,FALSE),IF(B${rowNum}<>"",MAX(10, 20*LOG10(D${rowNum}) + 5),0))` };
            row.getCell(6).value = { formula: `IF(B${rowNum}="",0,VLOOKUP(B${rowNum},BAZA_MATERIJALA!$B$2:$E$${materials.length + 1},4,FALSE))` };
        }

        const resRow = startRow + 10;
        calcSheet.getCell(`C${resRow}`).value = 'UKUPNA MASA [kg/m2]:';
        calcSheet.getCell(`D${resRow}`).value = { formula: `SUM(D4:D11)` };
        calcSheet.getCell(`D${resRow}`).font = { bold: true };

        const rwRow = resRow + 1;
        calcSheet.getCell(`C${rwRow}`).value = 'RW_LAB (Teoretski):';
        calcSheet.getCell(`D${rwRow}`).value = { formula: `MAX(E4:E11, 20*LOG10(D${resRow})+10)` };

        const rprimeRow = rwRow + 1;
        calcSheet.getCell(`C${rprimeRow}`).value = "R'W (TERENSKI -3dB):";
        calcSheet.getCell(`D${rprimeRow}`).value = { formula: `D${rwRow}-3` };
        calcSheet.getCell(`D${rprimeRow}`).font = { bold: true, color: { argb: 'FF0000FF' } };

        const lprimeRow = rprimeRow + 1;
        calcSheet.getCell(`C${lprimeRow}`).value = "L'NW (UDARNA BUKA):";
        calcSheet.getCell(`D${lprimeRow}`).value = { formula: `IF(D${rwRow}<40,"N/A",80 - SUM(F4:F11) - 0.1*D${resRow})` };

        // 3. ISKAZNICA SHEET
        const iskaznica = workbook.addWorksheet('ISKAZNICA');
        iskaznica.getCell('A1').value = 'ISKAZNICA AKUSTIČKIH SVOJSTAVA ZGRADE';
        iskaznica.getCell('A1').font = { bold: true, size: 16 };
        
        iskaznica.getCell('A3').value = 'Građevina:';
        iskaznica.getCell('B3').value = buildingName;
        iskaznica.getCell('B3').font = { bold: true };

        iskaznica.getCell('A4').value = 'Datum kreiranja:';
        iskaznica.getCell('B4').value = new Date().toLocaleDateString('hr-HR');

        iskaznica.getCell('A6').value = 'PRORAČUNATI ELEMENTI:';
        iskaznica.getRow(7).values = ['Element', 'Putanja', 'Zahtjev (dB)', 'Projektirano (dB)', 'STATUS'];
        iskaznica.getRow(7).font = { bold: true };
        
        iskaznica.getRow(8).values = ['Zid / Element', 'Proracun!D15', '52', { formula: 'PRORACUN!D15' }, { formula: 'IF(D8>=52,"PROLAZI","NE PROLAZI")' }];
        if (layers && layers.some(l => l.materialName.includes('estrih') || l.materialName.includes('EPS-T'))) {
            iskaznica.getRow(9).values = ['Međukatna konstr.', 'Proracun!D16', '55', { formula: 'PRORACUN!D16' }, { formula: 'IF(D9<=55,"PROLAZI","NE PROLAZI")' }];
        }

        iskaznica.columns = [{ width: 25 }, { width: 15 }, { width: 15 }, { width: 20 }, { width: 15 }];

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

function exportData() {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(layers, null, 2));
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", "iskaznica_akustike.json");
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
}

window.onload = init;
