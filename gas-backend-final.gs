/**
 * ARCH-TOOLKIT | GOOGLE APPS SCRIPT BACKEND (v3.1)
 * Automatsko razvrstavanje: AKUSTIKA i TROŠKOVNIK u zasebne tabove.
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Arch-Toolkit PWA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var ssId = "1EG_0RzwE7UsrFqJJu4Mwwr0WCtSpcNsXIv6j0uT3a1k";
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(ssId);
    
    if (data.type === 'troskovnik') {
      // --- SPREMANJE U TAB "Troškovnik" ---
      var sheet = ss.getSheetByName("Troškovnik") || ss.insertSheet("Troškovnik");
      if (sheet.getLastRow() == 0) {
        sheet.appendRow(["Datum", "Projekt/Pozicija", "Svi Podaci (JSON)", "Beton (m3)", "Oplata (m2)", "Špalete (m2)", "Armatura (kg)", "Email"]);
        sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#f1f5f9");
      }
      
      sheet.appendRow([
        new Date(),
        data.project || "Novi projekt",
        JSON.stringify(data.data),
        data.summary.beton,
        data.summary.oplata,
        data.summary.spalete,
        data.summary.armatura,
        data.email
      ]);
      
    } else {
      // --- SPREMANJE U TAB "Akustika" ---
      var sheet = ss.getSheetByName("Akustika") || ss.insertSheet("Akustika");
      if (sheet.getLastRow() == 0) {
        sheet.appendRow(["Datum", "Građevina", "Slojevi (JSON)", "Rezultat Rw", "Rezultat Lnw", "Cilj Rw", "Cilj Lnw", "Email"]);
        sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#e0f2fe");
      }
      
      var buildingName = data.buildingName || data.project || "Novi projekt";
      sheet.appendRow([
        new Date(),
        buildingName,
        JSON.stringify(data.layers),
        data.resultRw + " dB",
        (data.resultLnw || "N/A") + " dB",
        (data.reqRw || "52") + " dB",
        (data.reqLnw || "55") + " dB",
        data.email
      ]);
      
      if (data.email && data.email.includes("@")) {
        sendAcousticEmail(data);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({"status": "success"}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sendAcousticEmail(data) {
  var subject = "Izvještaj: Iskaznica Akustike - " + (data.buildingName || "Projekt");
  var htmlTable = "<table border='1' style='border-collapse: collapse; width: 100%; font-family: sans-serif;'>" +
                  "<tr style='background-color: #f2f2f2;'><th>#</th><th>Materijal</th><th>Debljina (cm)</th><th>Masa (kg/m²)</th></tr>";
  if (data.layers && data.layers.length > 0) {
    data.layers.forEach(function(l, i) {
      var rho = getRho(l.materialName);
      htmlTable += "<tr><td style='padding: 8px;'>" + (i+1) + "</td><td style='padding: 8px;'>" + l.materialName + "</td><td style='padding: 8px; text-align: center;'>" + l.thickness + "</td><td style='padding: 8px; text-align: center;'>" + Math.round(rho * l.thickness / 100) + "</td></tr>";
    });
  }
  htmlTable += "</table>";
  var resTable = "<table border='1' style='border-collapse: collapse; width: 100%; font-family: sans-serif; margin-top: 20px;'>" +
                 "<tr style='background-color: #e6f3ff;'><th>Parametar</th><th>Rezultat</th><th>Cilj</th></tr>" +
                 "<tr><td style='padding: 8px;'>Ukupna masa</td><td style='padding: 8px; text-align: center;'>" + (data.totalMass || "-") + "</td><td style='padding: 8px; text-align: center;'>-</td></tr>" +
                 "<tr><td style='padding: 8px;'><b>Zračna izolacija (Rw)</b></td><td style='padding: 8px; text-align: center;'><b>" + data.resultRw + "</b></td><td style='padding: 8px; text-align: center;'>≥ " + (data.reqRw || "52") + " dB</td></tr>";
  if (data.resultLnw && data.resultLnw !== "N/A" && data.resultLnw !== "0 dB") {
    resTable += "<tr><td style='padding: 8px;'><b>Udarna buka (Lnw)</b></td><td style='padding: 8px; text-align: center;'><b>" + data.resultLnw + "</b></td><td style='padding: 8px; text-align: center;'>≤ " + (data.reqLnw || "55") + " dB</td></tr>";
  }
  resTable += "</table>";
  var htmlBody = "<html><body>" + htmlTable + "<h4>Rezultati</h4>" + resTable + "</body></html>";
  MailApp.sendEmail({ to: data.email, subject: subject, htmlBody: htmlBody });
}

function getRho(n) {
  var db = {'Armirani beton': 2500, 'Puna opeka': 1800, 'Porotherm 25 S': 800, 'Porotherm 20 AKU': 1200, 'Ytong 20': 550, 'EPS-T (elastificirani)': 15, 'XPS 300 kPa': 33, 'XPS 500 kPa': 42, 'XPS 700 kPa': 48, 'Kamena vuna Floor': 100, 'Cementni estrih': 2200, 'Tekući estrih': 2400, 'Keramika / Pločice': 2300, 'Parket 14mm': 700};
  return db[n] || 0;
}
