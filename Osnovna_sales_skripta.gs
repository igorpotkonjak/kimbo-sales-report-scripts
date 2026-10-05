function posaljiJutarnjiIzvestajSaExcelPrilogom() {
  var pocetnoVreme = new Date().getTime();
  Logger.log("=== POČETAK IZVRŠAVANJA SKRIPTE ===");
  
  // Lista glavnih primaoca (To) - dodati servis@kimbo.rs i srdjan.mladenovic@kimbo.rs
  var primaociTo = "nikola.rakic@kimbo.rs, zoran.jaric@kimbo.rs, nebojsa.petrovic@kimbo.rs, srecko.tocakovic@kimbo.rs, vukasin.aleksic@kimbo.rs, servis@kimbo.rs, srdjan.mladenovic@kimbo.rs";
  
  // Lista primaoca u kopiji (CC)
  var primaociCc = "veran.adamovic@mojipartneri.rs, djordje.despotovic@mojipartneri.rs, suzana.lazarevic@kimbo.rs, filip.gavrilovic@kimbo.rs, igor.potkonjak@gmail.com, natasa@vinteam.rs";
  
  var naslov = "Sales izvestaj";
  
  // ID dokumenta iz koga se preuzimaju tabele za prikaz u mejlu
  var spreadsheetIdZaTabele = "1diIpCdTWRfJKFBSPBvSn4f7nfJiVRRuqYNYRGcdEh74"; 
  var ssTabele = SpreadsheetApp.openById(spreadsheetIdZaTabele);

  // --- KORAK ZA PROVERU I FORSIRANJE IMPORTRANGE-A U SVA 4 SHEET-A ---
  Logger.log("Osvežavam IMPORTRANGE u svim listovima (A1)...");
  var imenaSheetova = ["Za slanje 1", "Za slanje 2", "Za slanje 3", "Za slanje 4"];
  
  for (var i = 0; i < imenaSheetova.length; i++) {
    var sheetProvera = ssTabele.getSheetByName(imenaSheetova[i]);
    if (sheetProvera) {
      var celijaA1 = sheetProvera.getRange("A1");
      var postojecaFormula = celijaA1.getFormula();
      
      // Ponovno upisujemo formulu u A1 da nateramo Google na osvežavanje
      if (postojecaFormula && postojecaFormula.toUpperCase().indexOf("IMPORTRANGE") !== -1) {
        celijaA1.setFormula(postojecaFormula);
      }
    }
  }
  
  // Forsiramo primenu izmena i čekamo da Google povuče sve podatke sa mreže
  SpreadsheetApp.flush();
  Utilities.sleep(4000); 
  SpreadsheetApp.flush();
  
  // Provera statusa u prvom sheet-u da li se i dalje prikazuje učitavanje
  var sheet1Provera = ssTabele.getSheetByName("Za slanje 1");
  if (sheet1Provera) {
    var trenutniPrikaz = sheet1Provera.getRange("A1").getDisplayValue();
    if (trenutniPrikaz === "#LOADING..." || trenutniPrikaz === "#N/A" || trenutniPrikaz === "#REF!") {
      Logger.log("IMPORTRANGE još učitava (" + trenutniPrikaz + "). Čekam još 5 sekundi...");
      Utilities.sleep(5000);
      SpreadsheetApp.flush();
    }
  }
  // -----------------------------------------------------------------

  // ID originalnog dokumenta koji se konvertuje u Excel i šalje kao prilog (attachment)
  var spreadsheetIdZaPrilog = "1xfkoP0qFZSw3LD0jRrYnMnZqGUsqwleHubGg2xUXhrs";
  var ssPrilog = SpreadsheetApp.openById(spreadsheetIdZaPrilog);
  
  // Optimizovana funkcija za brzo kreiranje HTML tabele korišćenjem skupnih poziva (Batching)
  function generisiHtmlTabeluBrzo(sheet, startRow, endRow, startCol, endCol) {
    if (!sheet) return "";
    
    var numRows = endRow - startRow + 1;
    var numCols = endCol - startCol + 1;
    var range = sheet.getRange(startRow, startCol, numRows, numCols);
    
    // Izvlačimo sve podatke odjednom u memoriju
    var displayValues = range.getDisplayValues();
    var backgrounds = range.getBackgrounds();
    var fontColors = range.getFontColors();
    var fontWeights = range.getFontWeights();
    var textStyles = range.getTextStyles();
    var alignments = range.getHorizontalAlignments();
    
    var html = '<table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); width: auto;">';
    
    for (var r = 0; r < numRows; r++) {
      var mehanickiRed = startRow + r;
      if (sheet.isRowHiddenByUser(mehanickiRed)) {
        continue; // Preskače skrivene redove
      }
      
      html += '<tr style="height: ' + sheet.getRowHeight(mehanickiRed) + 'px;">';
      
      for (var c = 0; c < numCols; c++) {
        var mehanickaKolona = startCol + c;
        if (sheet.isColumnHiddenByUser(mehanickaKolona)) {
          continue; // Preskače skrivene kolone
        }
        
        var text = displayValues[r][c];
        var bgColor = backgrounds[r][c];
        var fontColor = fontColors[r][c];
        var fontWeight = fontWeights[r][c];
        var textStyle = textStyles[r][c];
        var align = alignments[r][c];
        
        if (bgColor == "#ffffff" || bgColor == "") {
          bgColor = "transparent";
        }
        
        var stil = 'padding: 6px 8px; ' +
                   'background-color: ' + bgColor + '; ' +
                   'color: ' + fontColor + '; ' +
                   'font-weight: ' + fontWeight + '; ' +
                   'text-align: ' + align + '; ' +
                   'border: 1px solid #ddd; ' + 
                   'white-space: nowrap; ';    
                   
        if (textStyle.isItalic()) stil += 'font-style: italic; ';
        if (textStyle.isStrikethrough()) stil += 'text-decoration: line-through; ';
        
        html += '<td style="' + stil + '">' + text + '</td>';
      }
      html += '</tr>';
    }
    html += '</table>';
    return html;
  }

  var prolaznoVreme = new Date().getTime();
  Logger.log("Učitavanje dokumenta trajalo: " + (prolaznoVreme - pocetnoVreme) + " ms");
  var poslednjeVreme = prolaznoVreme;

  // 1. GENERISANJE TABELE 1 ("Za slanje 1" - A4 do AR28, kolone 1-44)
  var sheet1 = ssTabele.getSheetByName("Za slanje 1");
  var htmlTabela1 = generisiHtmlTabeluBrzo(sheet1, 4, 28, 1, 44);
  
  prolaznoVreme = new Date().getTime();
  Logger.log("Generisanje TABELE 1 (Za slanje 1 - A4:AR28) trajalo: " + (prolaznoVreme - poslednjeVreme) + " ms");
  poslednjeVreme = prolaznoVreme;

  // 2. GENERISANJE TABELE 2 ("Za slanje 2" - A4 do AR28, kolone 1-44)
  var sheet2 = ssTabele.getSheetByName("Za slanje 2");
  var htmlTabela2 = generisiHtmlTabeluBrzo(sheet2, 4, 28, 1, 44);
  
  prolaznoVreme = new Date().getTime();
  Logger.log("Generisanje TABELE 2 (Za slanje 2 - A4:AR28) trajalo: " + (prolaznoVreme - poslednjeVreme) + " ms");
  poslednjeVreme = prolaznoVreme;

  // 3. GENERISANJE TABELE 3 ("Za slanje 3" - A4 do AL17, kolone 1-38)
  var sheet3 = ssTabele.getSheetByName("Za slanje 3");
  var htmlTabela3 = generisiHtmlTabeluBrzo(sheet3, 4, 17, 1, 38);
  
  prolaznoVreme = new Date().getTime();
  Logger.log("Generisanje TABELE 3 (Za slanje 3 - A4:AL17) trajalo: " + (prolaznoVreme - poslednjeVreme) + " ms");
  poslednjeVreme = prolaznoVreme;

  // 4. GENERISANJE TABELE 4 ("Za slanje 4" - A4 do AL17, kolone 1-38)
  var sheet4 = ssTabele.getSheetByName("Za slanje 4");
  var htmlTabela4 = generisiHtmlTabeluBrzo(sheet4, 4, 17, 1, 38);
  
  prolaznoVreme = new Date().getTime();
  Logger.log("Generisanje TABELE 4 (Za slanje 4 - A4:AL17) trajalo: " + (prolaznoVreme - poslednjeVreme) + " ms");
  poslednjeVreme = prolaznoVreme;

  // 5. SKLAPANJE TELA MEJLA
  var htmlTelo = '<div style="font-family: Arial, sans-serif; color: #333; line-height: 1.5;">' +
                 "Dobar dan,<br><br>" +
                 
                 "Tracking volume po kategoriji:<br><br>" +
                 '<div style="overflow-x: auto;">' + htmlTabela1 + '</div><br><br>' + 
                 
                 "Tracking volume po zaposlenom<br><br>" + 
                 '<div style="overflow-x: auto;">' + htmlTabela3 + '</div><br><br>' + 
                 
                 "Tracking value po kategoriji.<br><br>" + 
                 '<div style="overflow-x: auto;">' + htmlTabela2 + '</div><br><br>' + 
                 
                 "Tracking value po zaposlenom<br><br>" + 
                 '<div style="overflow-x: auto;">' + htmlTabela4 + '</div><br><br>' + 
                 
                 "U prilogu mejla možete preuzeti ceo Excel dokument (.xlsx) sa svim listovima.<br><br>" +
                 "Srdačan pozdrav" +
                 '</div>';

  // 6. PRIPREMA CELOG DOKUMENTA KAO XLSX ATTACHMENT (iz originalnog fajla)
  var excelUrl = "https://docs.google.com/spreadsheets/d/" + spreadsheetIdZaPrilog + "/export?format=xlsx";
  
  var excelOdgovor = UrlFetchApp.fetch(excelUrl, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });
  
  var excelBlob = excelOdgovor.getBlob().setName(ssPrilog.getName() + ".xlsx");
  
  prolaznoVreme = new Date().getTime();
  Logger.log("Priprema Excel priloga (.xlsx) trajala: " + (prolaznoVreme - poslednjeVreme) + " ms");
  poslednjeVreme = prolaznoVreme;
                 
  // 7. SLANJE MEJLA SA TO I CC PRIMAOCIMA
  GmailApp.sendEmail(primaociTo, naslov, "", {
    cc: primaociCc,
    htmlBody: htmlTelo,
    attachments: [excelBlob]
  });
  
  prolaznoVreme = new Date().getTime();
  Logger.log("Slanje mejla preko GmailApp trajalo: " + (prolaznoVreme - poslednjeVreme) + " ms");
  
  var ukupnoVreme = (prolaznoVreme - pocetnoVreme) / 1000;
  Logger.log("=== UKUPNO VREME IZVRŠAVANJA: " + ukupnoVreme.toFixed(2) + " sekundi ===");
}
