function posaljiSalesIzvestajSkuIKupci() {
  var pocetnoVreme = new Date().getTime();
  Logger.log("=== POČETAK IZVRŠAVANJA SKRIPTE ===");
  
  // Lista glavnih primaoca (To) - izbačen Nebojša Petrović, dodat Nikola Popović
  var primaociTo = "nikola.rakic@kimbo.rs, zoran.jaric@kimbo.rs, srecko.tocakovic@kimbo.rs, vukasin.aleksic@kimbo.rs, servis@kimbo.rs, srdjan.mladenovic@kimbo.rs, nikola.popovic@kimbo.rs";
  
  // Lista primaoca u kopiji (CC)
  var primaociCc = "veran.adamovic@mojipartneri.rs, djordje.despotovic@mojipartneri.rs, suzana.lazarevic@kimbo.rs, filip.gavrilovic@kimbo.rs, igor.potkonjak@gmail.com, natasa@vinteam.rs";
  
  // Naslov mejla
  var naslov = "Sales izvestaj - SKU i kupci";
  
  // ID dokumenta iz koga se preuzimaju tabele i koji se pakuje u Excel
  var spreadsheetId = "1spAv_NmzjSIy_Ga2x3VdQ7wqEAGd-2_7M9y_z3goCl4"; 
  var ss = SpreadsheetApp.openById(spreadsheetId);

  // Optimizovana funkcija za brzo kreiranje HTML tabele (preskače skrivene/filtrirane redove i skrivene kolone)
  function generisiHtmlTabeluBrzo(sheet, startRow, endRow, startCol, endCol) {
    if (!sheet) return "";
    
    var numRows = endRow - startRow + 1;
    var numCols = endCol - startCol + 1;
    var range = sheet.getRange(startRow, startCol, numRows, numCols);
    
    var displayValues = range.getDisplayValues();
    var backgrounds = range.getBackgrounds();
    var fontColors = range.getFontColors();
    var fontWeights = range.getFontWeights();
    var textStyles = range.getTextStyles();
    var alignments = range.getHorizontalAlignments();
    
    var html = '<table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); width: auto;">';
    
    for (var r = 0; r < numRows; r++) {
      var mehanickiRed = startRow + r;
      
      // Provera da li je red ručno sakriven ILI je sakriven filterom
      if (sheet.isRowHiddenByUser(mehanickiRed) || sheet.isRowHiddenByFilter(mehanickiRed)) {
        continue;
      }
      
      html += '<tr style="height: ' + sheet.getRowHeight(mehanickiRed) + 'px;">';
      
      for (var c = 0; c < numCols; c++) {
        var mehanickaKolona = startCol + c;
        if (sheet.isColumnHiddenByUser(mehanickaKolona)) {
          continue; 
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

  // 1. GENERISANJE TABELE 1 ("Vol po SKU" - A5 do AO55, kolone 1-41)
  var sheet1 = ss.getSheetByName("Vol po SKU");
  var htmlTabela1 = generisiHtmlTabeluBrzo(sheet1, 5, 55, 1, 41);
  
  prolaznoVreme = new Date().getTime();
  Logger.log("Generisanje TABELE 1 trajalo: " + (prolaznoVreme - poslednjeVreme) + " ms");
  poslednjeVreme = prolaznoVreme;

  // 2. GENERISANJE TABELE 2 ("Vol po kupcu total" - A5 do AL36, kolone 1-38)
  var sheet2 = ss.getSheetByName("Vol po kupcu total");
  var htmlTabela2 = generisiHtmlTabeluBrzo(sheet2, 5, 36, 1, 38);
  
  prolaznoVreme = new Date().getTime();
  Logger.log("Generisanje TABELE 2 trajalo: " + (prolaznoVreme - poslednjeVreme) + " ms");
  poslednjeVreme = prolaznoVreme;

  // 3. GENERISANJE TABELE 3 ("Vol po kupcu best" - A5 do AL36, kolone 1-38)
  var sheet3 = ss.getSheetByName("Vol po kupcu best");
  var htmlTabela3 = generisiHtmlTabeluBrzo(sheet3, 5, 36, 1, 38);
  
  prolaznoVreme = new Date().getTime();
  Logger.log("Generisanje TABELE 3 trajalo: " + (prolaznoVreme - poslednjeVreme) + " ms");
  poslednjeVreme = prolaznoVreme;

  // 4. GENERISANJE TABELE 4 ("Vol po kupcu worst" - A5 do AL36, kolone 1-38)
  var sheet4 = ss.getSheetByName("Vol po kupcu worst");
  var htmlTabela4 = generisiHtmlTabeluBrzo(sheet4, 5, 36, 1, 38);
  
  prolaznoVreme = new Date().getTime();
  Logger.log("Generisanje TABELE 4 trajalo: " + (prolaznoVreme - poslednjeVreme) + " ms");
  poslednjeVreme = prolaznoVreme;

  // 5. SKLAPANJE TELA MEJLA
  var htmlTelo = '<div style="font-family: Arial, sans-serif; color: #333; line-height: 1.5;">' +
                 "Dobar dan,<br><br>" +
                 
                 "Volume po SKU:<br><br>" +
                 '<div style="overflow-x: auto;">' + htmlTabela1 + '</div><br><br>' + 
                 
                 "Volume top 30 kupaca:<br><br>" + 
                 '<div style="overflow-x: auto;">' + htmlTabela2 + '</div><br><br>' + 
                 
                 "Volume top 30 rastucih kupaca:<br><br>" + 
                 '<div style="overflow-x: auto;">' + htmlTabela3 + '</div><br><br>' + 
                 
                 "Volume top 30 opadajucih kupaca:<br><br>" + 
                 '<div style="overflow-x: auto;">' + htmlTabela4 + '</div><br><br>' + 
                 
                 "U prilogu mejla možete preuzeti ceo Excel dokument (.xlsx) sa svim listovima.<br><br>" +
                 "Srdačan pozdrav" +
                 '</div>';

  // 6. PRIPREMA CELOG DOKUMENTA KAO XLSX ATTACHMENT
  var excelUrl = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?format=xlsx";
  
  var excelOdgovor = UrlFetchApp.fetch(excelUrl, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });
  
  var excelBlob = excelOdgovor.getBlob().setName(ss.getName() + ".xlsx");
  
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
