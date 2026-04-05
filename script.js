function registerVisaExpenses() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const label = GmailApp.getUserLabelByName('Visa');

  if (!label) {
    return;
  }

  const threads = label.getThreads();

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const message of messages) {
      if (message.isUnread()) {
        const originalBody = message.getPlainBody();
        const lowercaseBody = originalBody.toLowerCase();

        try {
          const isReversal = lowercaseBody.includes('anulación');
          const clpAmountMatch = originalBody.match(/Monto\s+\$([\d.]+)/);
          const usdAmountMatch = originalBody.match(/Monto\s+USD\s+([\d.]+)/);
          let domesticAmount = clpAmountMatch
            ? parseInt(clpAmountMatch[1].replace(/\./g, ''))
            : 0;
          let internationalAmount = usdAmountMatch
            ? parseFloat(usdAmountMatch[1])
            : 0;
          if (isReversal) {
            domesticAmount *= -1;
            internationalAmount *= -1;
          }
          const purchaseDateMatch = originalBody.match(
            /Fecha\s+(\d{2}\/\d{2}\/\d{4})/,
          );
          const purchaseDateStr = purchaseDateMatch ? purchaseDateMatch[1] : '';
          if (!purchaseDateStr) continue;
          const merchantMatch = originalBody.match(/Comercio\s+(.+)/);
          let merchant = merchantMatch
            ? merchantMatch[1].trim()
            : 'Desconocido';
          const cardMatch = originalBody.match(
            /Número tarjeta crédito\s+\*{4}(\d{4})/,
          );
          const visaLast4 = cardMatch ? cardMatch[1] : '';
          const installmentsMatch = originalBody.match(/Cuotas\s+(\d+)/);
          const installments = installmentsMatch ? installmentsMatch[1] : '0';

          // --- SHEET NAME LOGIC ---
          const dateParts = purchaseDateStr.split('/');
          let purchaseDate = new Date(
            dateParts[2],
            dateParts[1] - 1,
            dateParts[0],
          );
          let billingMonth = purchaseDate.getMonth();
          let billingYear = purchaseDate.getFullYear();
          if (purchaseDate.getDate() >= 20) {
            billingMonth++;
            if (billingMonth > 11) {
              billingMonth = 0;
              billingYear++;
            }
          }
          const monthNames = [
            'Enero',
            'Febrero',
            'Marzo',
            'Abril',
            'Mayo',
            'Junio',
            'Julio',
            'Agosto',
            'Septiembre',
            'Octubre',
            'Noviembre',
            'Diciembre',
          ];
          const sheetName =
            'Facturacion ' + monthNames[billingMonth] + ' ' + billingYear;

          let sheet = spreadsheet.getSheetByName(sheetName);

          if (!sheet) {
            sheet = spreadsheet.insertSheet(sheetName);
            SpreadsheetApp.flush();

            Utilities.sleep(2000);

            const headers = [
              'Fecha',
              'Comercio',
              'Nacional',
              'Internacional',
              'Cuotas',
              'Visa',
              'Total Nacional',
              'Total Internacional',
            ];
            sheet
              .getRange(1, 1, 1, 8)
              .setValues([headers])
              .setFontWeight('bold')
              .setBackground('#f3f3f3');
            sheet.setFrozenRows(1);
            SpreadsheetApp.flush();

            Utilities.sleep(2000);

            sheet.getRange('G2').setFormula('=SUM(C2:C1000)');
            SpreadsheetApp.flush();

            sheet.getRange('H2').setFormula('=SUM(D2:D1000)');
            SpreadsheetApp.flush();
          }

          sheet.getRange('G2').setNumberFormat('$#,##0');
          sheet.getRange('H2').setNumberFormat('USD #,##0.00');
          SpreadsheetApp.flush();

          let columnAValues = sheet.getRange('A:A').getValues();
          let targetRow = 2;
          for (let i = 0; i < columnAValues.length; i++) {
            if (columnAValues[i][0] === '' && i > 0) {
              targetRow = i + 1;
              break;
            }
          }

          sheet
            .getRange(targetRow, 1, 1, 6)
            .setValues([
              [
                purchaseDateStr,
                merchant,
                domesticAmount,
                internationalAmount,
                installments,
                visaLast4,
              ],
            ]);

          sheet.getRange(targetRow, 3).setNumberFormat('$#,##0');
          sheet.getRange(targetRow, 4).setNumberFormat('USD #,##0.00');

          if (isReversal) {
            sheet.getRange(targetRow, 1, 1, 6).setBackground('#fff0f0');
          }

          message.markRead();
          SpreadsheetApp.flush();
        } catch (error) {
          console.error('❌ ERROR DETALLADO: ' + error.message);
          console.error('Línea del error: ' + error.stack);
        }
      }
    }
  }
}
