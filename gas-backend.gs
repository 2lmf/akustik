/**
 * Google Apps Script (GAS) Backend za Iskaznicu Akustike
 */

function doPost(e) {
  try {
    var ssId = "1EG_0RzwE7UsrFqJJu4Mwwr0WCtSpcNsXIv6j0uT3a1k";
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = ss.getSheets()[0];
    
    if (sheet.getLastRow() == 0) {
      sheet.appendRow(["Datum", "Tip", "Projekt/Građevina", "Podaci (JSON)", "Rezultat 1", "Rezultat 2", "Email"]);
    }
    
    if (data.type === 'troskovnik') {
      sheet.appendRow([
        new Date(),
        "TROŠKOVNIK",
        data.project || "Novi projekt",
        JSON.stringify(data.data),
        data.summary.beton,
        data.summary.oplata + " | " + data.summary.armatura,
        data.email
      ]);
    } else {
      // Postojeći tip (Akustika)
      sheet.appendRow([
        new Date(),
        "AKUSTIKA",
        data.buildingName || "Novi projekt",
        JSON.stringify(data.layers),
        data.resultRw,
        data.resultLnw || "N/A",
        data.email
      ]);
      if (data.email) {
        sendAcousticEmail(data);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      "status": "success",
      "message": "Podaci spremljeni i HTML tablica poslana."
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function sendAcousticEmail(data) {
  var subject = "Izvještaj: Iskaznica Akustičkih Svojstava - " + (data.buildingName || "Projekt");
  
  var htmlTable = "<table border='1' style='border-collapse: collapse; width: 100%; font-family: sans-serif;'>" +
                  "<tr style='background-color: #f2f2f2;'><th>#</th><th>Materijal</th><th>Debljina (cm)</th><th>Masa (kg/m²)</th></tr>";
  
  data.layers.forEach(function(l, i) {
    var rho = getRho(l.materialName);
    var mass = Math.round(rho * l.thickness / 100);
    htmlTable += "<tr><td style='padding: 8px;'>" + (i + 1) + "</td>" +
                 "<td style='padding: 8px;'>" + l.materialName + "</td>" +
                 "<td style='padding: 8px; text-align: center;'>" + l.thickness + "</td>" +
                 "<td style='padding: 8px; text-align: center;'>" + mass + "</td></tr>";
  });
  htmlTable += "</table>";

  var resTable = "<table border='1' style='border-collapse: collapse; width: 100%; font-family: sans-serif; margin-top: 20px;'>" +
                 "<tr style='background-color: #e6f3ff;'><th>Parametar</th><th>Rezultat</th><th>Uvjet</th></tr>" +
                 "<tr><td style='padding: 8px;'>Ukupna masa</td><td style='padding: 8px; text-align: center;'>" + data.totalMass + "</td><td style='padding: 8px; text-align: center;'>-</td></tr>" +
                 "<tr><td style='padding: 8px;'><b>Zračna izolacija (Rw)</b></td><td style='padding: 8px; text-align: center;'><b>" + data.resultRw + "</b></td><td style='padding: 8px; text-align: center;'>≥ 52 dB</td></tr>";
  
  if (data.resultLnw !== "N/A" && data.resultLnw) {
    resTable += "<tr><td style='padding: 8px;'><b>Udarna buka (Lnw)</b></td><td style='padding: 8px; text-align: center;'><b>" + data.resultLnw + "</b></td><td style='padding: 8px; text-align: center;'>≤ 55 dB</td></tr>";
  }
  resTable += "</table>";

  var htmlBody = "<html><body style='font-family: sans-serif;'>" +
                 "<h3>Akustički Izvještaj: " + (data.buildingName || "Novi projekt") + "</h3>" +
                 "<p>U nastavku su rezultati proračuna. Tablice ispod možete direktno označiti i zalijepiti (Copy/Paste) u svoj Excel:</p>" +
                 "<h4>1. Sastav konstrukcije</h4>" + htmlTable +
                 "<h4>2. Finalni rezultati</h4>" + resTable +
                 "<p style='margin-top: 20px; color: #666; font-size: 0.9rem;'>Ugodan rad,<br>Vaš Iskaznica Akustike Web App</p>" +
                 "</body></html>";
  
  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    htmlBody: htmlBody
  });
}

function getRho(name) {
  var db = {
    'Armirani beton': 2500,
    'Puna opeka': 1800,
    'Porotherm 25 S': 800,
    'Porotherm 20 AKU': 1200, 
    'Ytong 20': 550,
    'EPS-T (zvučni)': 15,
    'EPS 100': 20,
    'XPS 300 kPa': 33,
    'XPS 500 kPa': 42,
    'XPS 700 kPa': 48,
    'Kamena vuna Floor': 100,
    'Čepasta ploča H30': 20,
    'Cementni estrih': 2200,
    'Anhidritni estrih': 2400,
    'Keramika': 2300,
    'Parket 14mm': 700
  };
  return db[name] || 0;
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Iskaznica Akustike - Web App')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
