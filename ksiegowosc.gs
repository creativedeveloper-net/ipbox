/**
 * Tworzy nowy arkusz kalkulacyjny Google Sheets do rozliczeń IP Box.
 * Plik będzie zawierał cztery zakładki: Zmienne, Dochody, Koszty, Podsumowanie.
 */

// Wersja szablonu
const VERSION = "v2.2.0";

function createIpBoxSpreadsheet() {
  try {
    Logger.log(`Rozpoczynam tworzenie arkusza: Szablon Rozliczeń IP Box ${VERSION}...`);
    const spreadsheet = SpreadsheetApp.create(`Szablon Rozliczeń IP Box ${VERSION}`);
    const year = new Date().getFullYear();
    
    // Usunięcie domyślnego arkusza i stworzenie Zmienne
    const defaultSheet = spreadsheet.getSheets()[0];
    const sheetZmienne = spreadsheet.insertSheet('Zmienne', 0);
    spreadsheet.deleteSheet(defaultSheet);
    
    SpreadsheetApp.flush();
    Logger.log("Utworzono arkusz 'Zmienne'.");

    setupZmienneSheet(sheetZmienne);
    Logger.log("Skonfigurowano arkusz 'Zmienne'.");

    // Tworzenie 12 zestawów zakładek miesięcznych (Dochody, Koszty, Podsumowanie)
    for (let month = 1; month <= 12; month++) {
      const monthStr = month.toString().padStart(2, '0');
      
      Logger.log(`Tworzenie zakładek dla miesiąca ${monthStr}...`);
      
      const sheetDochody = spreadsheet.insertSheet(`${monthStr} Dochody`);
      setupDochodySheet(sheetDochody, month);
      
      const sheetKoszty = spreadsheet.insertSheet(`${monthStr} Koszty`);
      setupKosztySheet(sheetKoszty, sheetZmienne, month);
      
      const sheetPodsumowanie = spreadsheet.insertSheet(`${monthStr} Podsumowanie`);
      setupPodsumowanieMiesieczneSheet(sheetPodsumowanie, month, year);
      
      Logger.log(`Skonfigurowano zakładki dla miesiąca ${monthStr}.`);
    }

    // Tworzenie arkusza Podsumowanie Roczne
    Logger.log("Tworzenie arkusza 'Podsumowanie Roczne'...");
    const sheetPodsumowanieRoczne = spreadsheet.insertSheet('Podsumowanie Roczne');
    setupPodsumowanieRoczneSheet(sheetPodsumowanieRoczne, year);
    Logger.log("Skonfigurowano arkusz 'Podsumowanie Roczne'.");

    // Tworzenie arkusza Dokumentacja IP BOX
    Logger.log("Tworzenie arkusza 'Dokumentacja IP BOX'...");
    const sheetDokumentacja = spreadsheet.insertSheet('Dokumentacja IP BOX');
    setupDokumentacjaIPBoxSheet(sheetDokumentacja);
    Logger.log("Skonfigurowano arkusz 'Dokumentacja IP BOX'.");
    
    // Ustawienie aktywnego arkusza na Podsumowanie Roczne
    spreadsheet.setActiveSheet(sheetPodsumowanieRoczne);

    const url = spreadsheet.getUrl();
    Logger.log(`SUKCES! Utworzono nowy szablon rozliczeń IP Box ${year}.`);
    Logger.log(`Link do pliku: ${url}`);
    Logger.log(`ID pliku: ${spreadsheet.getId()}`);

  } catch (e) {
    Logger.log(`WYSTĄPIŁ BŁĄD: ${e.toString()}`);
    Logger.log(`Stos: ${e.stack}`);
  }
}

/**
 * Konfiguruje arkusz "Zmienne".
 */
function setupZmienneSheet(sheet) {
  // Sekcja 1: Zmienne główne (A:B)
  sheet.getRange('A1:B1').setValues([['ZMIENNE GŁÓWNE', 'WARTOŚĆ']]).setFontWeight('bold').setBackground('#e0e0e0');
  sheet.getRange('A2:B5').setValues([
    ['Stawka podatku liniowego', '19%'],
    ['Stawka podatku IP BOX', '5%'],
    ['Składka zdrowotna (miesięcznie)', 381.78],
    ['Składki ZUS społeczne (miesięcznie)', 1485.31]
  ]);
  
  // Sekcja 2: Typy rozliczania kosztów (D:E)
  sheet.getRange('D1:E1').setValues([['TYPY ROZLICZANIA KOSZTÓW', 'WSPÓŁCZYNNIK']]).setFontWeight('bold').setBackground('#e0e0e0');
  sheet.getRange('D2:E4').setValues([
    ['Zwykły', 1.0],
    ['Samochód', 0.5],
    ['Mieszkanie 1', 0.125]
  ]);

  // Sekcja 3: Stawki VAT (G:H)
  sheet.getRange('G1:H1').setValues([['STAWKI VAT', 'WARTOŚĆ']]).setFontWeight('bold').setBackground('#e0e0e0');
  sheet.getRange('G2:H6').setValues([
    ['23%', 0.23],
    ['8%', 0.08],
    ['5%', 0.05],
    ['0%', 0],
    ['zw.', 0]
  ]);
  
  // Formatowanie
  sheet.setColumnWidths(1, 8, 200);
  sheet.getRange('B2:B3').setNumberFormat('0%');
  sheet.getRange('B4:B5').setNumberFormat('#,##0.00" zł"');
  sheet.getRange('E2:E').setNumberFormat('0.00');
  sheet.getRange('H2:H').setNumberFormat('0%');
}

/**
 * Konfiguruje arkusz "Dochody" dla konkretnego miesiąca.
 */
function setupDochodySheet(sheet, month) {
  const headers = ['Data wystawienia', 'Nr faktury', 'Opis usługi', 'Kwota netto', 'Stawka VAT', 'Kwota brutto', 'IP BOX'];
  sheet.getRange('A1:G1').setValues([headers]).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  
  // Formatowanie
  sheet.getRange('A2:A').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('D2:D').setNumberFormat('#,##0.00" zł"');
  sheet.getRange('F2:F').setNumberFormat('#,##0.00" zł"');
  sheet.getRange('G2:G').insertCheckboxes();
  
  // Walidacja dla stawek VAT
  const vatRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getParent().getRange('Zmienne!G2:G'), true)
    .build();
  sheet.getRange('E2:E').setDataValidation(vatRule);
  
  // Formuła na Kwotę brutto
  const formulaBrutto = '=IF(D2<>"", D2 * (1 + IFERROR(VLOOKUP(E2, Zmienne!$G$2:$H, 2, FALSE), 0)), "")';
  sheet.getRange('F2').setFormula(formulaBrutto);
  
  // Ustawienie szerokości kolumn
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidths(4, 3, 120);
  sheet.setColumnWidth(7, 70);
  
  // Zabezpieczenie kolumny z formułą
  const protection = sheet.getRange('F2:F').protect();
  protection.setDescription('Kolumna obliczana automatycznie');
  protection.setWarningOnly(true);
}

/**
 * Konfiguruje arkusz "Koszty" dla konkretnego miesiąca.
 */
function setupKosztySheet(sheet, zmienneSheet, month) {
  const headers = [
    'Data wystawienia', 'Nr faktury', 'Opis', 'Kwota netto', 'Kwota brutto',
    'IP BOX', 'Typ rozliczania', 'Stawka VAT', 'Koszty netto do rozliczenia', 
    'VAT do rozliczenia'
  ];
  sheet.getRange('A1:J1').setValues([headers]).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  
  // Formatowanie
  sheet.getRange('A2:A').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('D2:D').setNumberFormat('#,##0.00" zł"');
  sheet.getRange('E2:E').setNumberFormat('#,##0.00" zł"');
  sheet.getRange('I2:J').setNumberFormat('#,##0.00" zł"');
  sheet.getRange('F2:F').insertCheckboxes();

  // Walidacja danych
  const typRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(zmienneSheet.getRange('D2:D'), true)
    .build();
  sheet.getRange('G2:G').setDataValidation(typRule);

  // Formuła na Stawkę VAT (obliczana z netto i brutto)
  const formulaStawkaVAT = '=IF(AND(D2<>"", E2<>""), TEXT(ROUND((E2/D2-1)*100, 0), "0") & "%", "")';
  sheet.getRange('H2').setFormula(formulaStawkaVAT);

  // Formuła na Koszty netto do rozliczenia
  const formulaNettoRozl = '=IF(AND(D2<>"", G2<>""), D2 * IFERROR(VLOOKUP(G2, Zmienne!$D$2:$E, 2, FALSE), 1), "")';
  sheet.getRange('I2').setFormula(formulaNettoRozl);

  // Formuła na VAT do rozliczenia
  const formulaVatRozl = '=IF(AND(E2<>"", D2<>"", G2<>""), (E2-D2) * IFERROR(VLOOKUP(G2, Zmienne!$D$2:$E, 2, FALSE), 1), "")';
  sheet.getRange('J2').setFormula(formulaVatRozl);

  // Ustawienie szerokości kolumn
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidths(4, 7, 120);

  // Ochrona kolumn z formułami
  const rangesToProtect = ['H2:H', 'I2:I', 'J2:J'];
  rangesToProtect.forEach(range => {
    const protection = sheet.getRange(range).protect();
    protection.setDescription('Kolumna obliczana automatycznie');
    protection.setWarningOnly(true);
  });
}

/**
 * Konfiguruje arkusz "XX Podsumowanie" dla konkretnego miesiąca.
 */
function setupPodsumowanieMiesieczneSheet(sheet, month, year) {
  const monthStr = month.toString().padStart(2, '0');
  const monthNames = [
    'Styczeń', 'Luty', 'Marzec', 'Kwiecień', 'Maj', 'Czerwiec',
    'Lipiec', 'Sierpień', 'Wrzesień', 'Październik', 'Listopad', 'Grudzień'
  ];
  
  // Nagłówek
  sheet.getRange('A1:B1').merge().setValue(`Podsumowanie ${monthNames[month-1]} ${year}`)
    .setFontWeight('bold').setFontSize(14).setBackground('#4285f4').setFontColor('#ffffff');
  
  const categories = [
    ['Kategoria', 'Wartość'],
    ['Przychody netto (IP BOX)', ''],
    ['Przychody netto (Inne)', ''],
    ['ŁĄCZNIE PRZYCHODY NETTO', ''],
    ['', ''],
    ['Koszty netto do rozliczenia (IP BOX)', ''],
    ['Koszty netto do rozliczenia (Inne)', ''],
    ['ŁĄCZNIE KOSZTY NETTO', ''],
    ['', ''],
    ['DOCHÓD (IP BOX)', ''],
    ['DOCHÓD (Inne)', ''],
    ['', ''],
    ['Składka zdrowotna', ''],
    ['Składki ZUS społeczne', ''],
    ['', ''],
    ['Zaliczka na podatek (IP BOX 5%)', ''],
    ['Zaliczka na podatek (Inne 19%)', ''],
    ['', ''],
    ['VAT należny (od sprzedaży)', ''],
    ['VAT naliczony (do odliczenia)', ''],
    ['VAT do zapłaty/zwrotu', '']
  ];
  
  sheet.getRange(2, 1, categories.length, 2).setValues(categories);
  sheet.getRange('A2:B2').setFontWeight('bold').setBackground('#e0e0e0');
  
  // Formuły w kolumnie B
  const dochodySheet = `'${monthStr} Dochody'`;
  const kosztySheet = `'${monthStr} Koszty'`;
  
  // Przychody
  sheet.getRange('B3').setFormula(`=SUMIF(${dochodySheet}!G:G, TRUE, ${dochodySheet}!D:D)`);
  sheet.getRange('B4').setFormula(`=SUMIF(${dochodySheet}!G:G, FALSE, ${dochodySheet}!D:D)`);
  sheet.getRange('B5').setFormula('=B3+B4');
  
  // Koszty
  sheet.getRange('B7').setFormula(`=SUMIF(${kosztySheet}!F:F, TRUE, ${kosztySheet}!I:I)`);
  sheet.getRange('B8').setFormula(`=SUMIF(${kosztySheet}!F:F, FALSE, ${kosztySheet}!I:I)`);
  sheet.getRange('B9').setFormula('=B7+B8');
  
  // Dochód
  sheet.getRange('B11').setFormula('=B3-B7');
  sheet.getRange('B12').setFormula('=B4-B8');
  
  // ZUS
  sheet.getRange('B14').setFormula('=IF(B5>0, Zmienne!$B$4, 0)');
  sheet.getRange('B15').setFormula('=IF(B5>0, Zmienne!$B$5, 0)');
  
  // Podatek
  sheet.getRange('B17').setFormula('=MAX(0, B11 * Zmienne!$B$3)');
  sheet.getRange('B18').setFormula('=MAX(0, (B12 - B15) * Zmienne!$B$2)');
  
  // VAT
  sheet.getRange('B20').setFormula(`=SUM(${dochodySheet}!F:F) - SUM(${dochodySheet}!D:D)`);
  sheet.getRange('B21').setFormula(`=SUM(${kosztySheet}!J:J)`);
  sheet.getRange('B22').setFormula('=B20-B21');
  
  // Formatowanie
  sheet.getRange('A5:B5').setFontWeight('bold');
  sheet.getRange('A9:B9').setFontWeight('bold');
  sheet.getRange('A11:B12').setFontWeight('bold');
  sheet.getRange('A17:B18').setFontWeight('bold');
  
  sheet.getRange('A6:B6').setBackground('#f3f3f3');
  sheet.getRange('A10:B10').setBackground('#f3f3f3');
  sheet.getRange('A13:B13').setBackground('#f3f3f3');
  sheet.getRange('A16:B16').setBackground('#f3f3f3');
  sheet.getRange('A19:B19').setBackground('#f3f3f3');
  
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 150);
  sheet.getRange('B3:B22').setNumberFormat('#,##0.00" zł"');
  
  // Ochrona kolumny z formułami
  const protection = sheet.getRange('B3:B22').protect();
  protection.setDescription('Kolumna obliczana automatycznie');
  protection.setWarningOnly(true);
}

/**
 * Konfiguruje arkusz "Podsumowanie Roczne".
 */
function setupPodsumowanieRoczneSheet(sheet, year) {
/**
 * Konfiguruje arkusz "Podsumowanie Roczne".
 */
function setupPodsumowanieRoczneSheet(sheet, year) {
  const months = [
    'Styczeń', 'Luty', 'Marzec', 'Kwiecień', 'Maj', 'Czerwiec', 
    'Lipiec', 'Sierpień', 'Wrzesień', 'Październik', 'Listopad', 'Grudzień'
  ];
  
  const headers = ['Kategoria', ...months, 'Razem'];
  sheet.getRange(1, 1, 1, 14).setValues([headers]).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');

  const categories = [
    'Przychody netto (IP BOX)',
    'Przychody netto (Inne)',
    'ŁĄCZNIE PRZYCHODY NETTO',
    '',
    'Koszty netto do rozliczenia (IP BOX)',
    'Koszty netto do rozliczenia (Inne)',
    'ŁĄCZNIE KOSZTY NETTO',
    '',
    'DOCHÓD (IP BOX)',
    'DOCHÓD (Inne)',
    '',
    'Składka zdrowotna',
    'Składki ZUS społeczne',
    '',
    'Zaliczka na podatek (IP BOX 5%)',
    'Zaliczka na podatek (Inne 19%)',
    '',
    'VAT należny (od sprzedaży)',
    'VAT naliczony (do odliczenia)',
    'VAT do zapłaty/zwrotu'
  ];

  sheet.getRange(2, 1, categories.length, 1).setValues(categories.map(c => [c]));
  
  // Formatowanie kategorii - puste wiersze
  sheet.getRange('A4:A4').setBackground('#f3f3f3');
  sheet.getRange('A8:A8').setBackground('#f3f3f3');
  sheet.getRange('A11:A11').setBackground('#f3f3f3');
  sheet.getRange('A14:A14').setBackground('#f3f3f3');
  sheet.getRange('A17:A17').setBackground('#f3f3f3');
  
  // Formatowanie kategorii - pogrubienie
  sheet.getRange('A3:N3').setFontWeight('bold');
  sheet.getRange('A7:N7').setFontWeight('bold');
  sheet.getRange('A9:N10').setFontWeight('bold');
  sheet.getRange('A15:N16').setFontWeight('bold');
  
  // Formuły dla każdego miesiąca - pobieranie z zakładek XX Podsumowanie
  for (let i = 0; i < 12; i++) {
    const month = i + 1;
    const monthStr = month.toString().padStart(2, '0');
    const col = String.fromCharCode(66 + i); // B, C, D...
    const podsumowanieSheet = `'${monthStr} Podsumowanie'`;
    
    // Mapowanie wierszy z miesięcznego podsumowania do rocznego
    sheet.getRange(`${col}2`).setFormula(`=${podsumowanieSheet}!B3`);  // Przychody IP BOX
    sheet.getRange(`${col}3`).setFormula(`=${podsumowanieSheet}!B4`);  // Przychody Inne
    sheet.getRange(`${col}4`).setFormula(`=${podsumowanieSheet}!B5`);  // Łącznie przychody
    
    sheet.getRange(`${col}6`).setFormula(`=${podsumowanieSheet}!B7`);  // Koszty IP BOX
    sheet.getRange(`${col}7`).setFormula(`=${podsumowanieSheet}!B8`);  // Koszty Inne
    sheet.getRange(`${col}8`).setFormula(`=${podsumowanieSheet}!B9`);  // Łącznie koszty
    
    sheet.getRange(`${col}10`).setFormula(`=${podsumowanieSheet}!B11`); // Dochód IP BOX
    sheet.getRange(`${col}11`).setFormula(`=${podsumowanieSheet}!B12`); // Dochód Inne
    
    sheet.getRange(`${col}13`).setFormula(`=${podsumowanieSheet}!B14`); // Składka zdrowotna
    sheet.getRange(`${col}14`).setFormula(`=${podsumowanieSheet}!B15`); // ZUS społeczne
    
    sheet.getRange(`${col}16`).setFormula(`=${podsumowanieSheet}!B17`); // Podatek IP BOX
    sheet.getRange(`${col}17`).setFormula(`=${podsumowanieSheet}!B18`); // Podatek Inne
    
    sheet.getRange(`${col}19`).setFormula(`=${podsumowanieSheet}!B20`); // VAT należny
    sheet.getRange(`${col}20`).setFormula(`=${podsumowanieSheet}!B21`); // VAT naliczony
    sheet.getRange(`${col}21`).setFormula(`=${podsumowanieSheet}!B22`); // VAT do zapłaty
  }
  
  // Formuły roczne (kolumna N)
  const rowsToSum = [2, 3, 4, 6, 7, 8, 10, 11, 13, 14, 16, 17, 19, 20, 21];
  rowsToSum.forEach(row => {
    sheet.getRange(`N${row}`).setFormula(`=SUM(B${row}:M${row})`);
  });

  // Formatowanie
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidths(2, 13, 120);
  sheet.getRange('B2:N21').setNumberFormat('#,##0.00" zł"');
  
  // Zamrożenie pierwszego wiersza i kolumny
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

/**
 * Konfiguruje arkusz "Dokumentacja IP BOX".
 */
function setupDokumentacjaIPBoxSheet(sheet) {
  const headers = ['ID Zadania (np. z Jiry)', 'Opis zadania/pracy B+R', 'Data', 'Powiązanie z fakturą (Nr faktury)'];
  sheet.getRange('A1:D1').setValues([headers]).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  
  // Formatowanie
  sheet.getRange('C2:C').setNumberFormat('yyyy-mm-dd');
  
  // Ustawienie szerokości kolumn
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 400);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 180);
  
  // Zamrożenie pierwszego wiersza
  sheet.setFrozenRows(1);
}