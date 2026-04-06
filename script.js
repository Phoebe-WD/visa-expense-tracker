/**
 * Reads unread Visa notification emails from Gmail, extracts purchase details
 * (amount, date, merchant, card number, installments), and logs each expense
 * into the corresponding monthly billing sheet in Google Sheets.
 *
 * - Emails must be labeled "Visa" in Gmail.
 * - Purchases on or after the 20th are billed to the next month's sheet.
 * - Cancellations ("anulación") are recorded as negative amounts with a red background.
 * - A new sheet is automatically created if the billing month doesn't exist yet.
 */
function registerVisaExpenses() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // ⚙️ CUSTOMIZE: Change 'Visa' to match your Gmail label name (e.g., 'Bank Alerts', 'Mastercard')
  const label = GmailApp.getUserLabelByName('Visa');

  if (!label) {
    console.warn('No se encontró la etiqueta Visa');
    return;
  }

  // Retrieve all email threads under the "Visa" label
  const threads = label.getThreads();
  console.log('Hilos encontrados: ' + threads.length);

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const message of messages) {
      // Only process unread messages to avoid duplicate entries
      if (message.isUnread()) {
        const originalBody = message.getPlainBody();
        const lowerBody = originalBody.toLowerCase();

        try {
          // =============================================
          // EXTRACTION: Parse transaction data from email
          // =============================================

          // ⚙️ CUSTOMIZE: Change 'anulación' to the keyword your bank uses for cancellations/reversals
          // (e.g., 'reversal', 'refund', 'chargeback')
          const isCancellation = lowerBody.includes('anulación');

          // ⚙️ CUSTOMIZE: Adjust these regex patterns to match your bank's email format.
          // The current patterns expect Chilean bank Visa notification format.
          // Example: if your bank says "Amount: $123.45", change to /Amount:\s+\$([\d.]+)/
          const clpAmountMatch = originalBody.match(/Monto\s+\$([\d.]+)/);
          const usdAmountMatch = originalBody.match(/Monto\s+USD\s+([\d.]+)/);

          // Parse amounts, removing dot thousand separators for CLP
          let domesticAmount = clpAmountMatch
            ? parseInt(clpAmountMatch[1].replace(/\./g, ''))
            : 0;
          let internationalAmount = usdAmountMatch
            ? parseFloat(usdAmountMatch[1])
            : 0;

          // Negate amounts if transaction is a cancellation
          if (isCancellation) {
            domesticAmount *= -1;
            internationalAmount *= -1;
          }

          // ⚙️ CUSTOMIZE: Adjust regex if your bank uses a different date label or format.
          // Current format: "Fecha DD/MM/YYYY". For "Date: MM/DD/YYYY", change to /Date:\s+(\d{2}\/\d{2}\/\d{4})/
          // and swap parts[0]/parts[1] in the Date constructor below accordingly.
          const dateMatch = originalBody.match(/Fecha\s+(\d{2}\/\d{2}\/\d{4})/);
          const purchaseDateStr = dateMatch ? dateMatch[1] : '';
          if (!purchaseDateStr) continue; // Skip if no date found

          // ⚙️ CUSTOMIZE: Change regex if your bank uses a different label for the merchant.
          // e.g., "Merchant: <name>" → /Merchant:\s+(.+)/
          const merchantMatch = originalBody.match(/Comercio\s+(.+)/);
          let merchant = merchantMatch
            ? merchantMatch[1].trim()
            : 'Desconocido';

          // ⚙️ CUSTOMIZE: Change regex to match your bank's card number format.
          // e.g., "Card ending in 1234" → /Card ending in\s+(\d{4})/
          const visaMatch = originalBody.match(
            /Número tarjeta crédito\s+\*{4}(\d{4})/,
          );
          const visa = visaMatch ? visaMatch[1] : '';

          // ⚙️ CUSTOMIZE: Change regex if your bank uses a different label for installments.
          // e.g., "Installments: 3" → /Installments:\s+(\d+)/
          // If your bank doesn't support installments, you can remove this.
          const installmentsMatch = originalBody.match(/Cuotas\s+(\d+)/);
          const installments = installmentsMatch ? installmentsMatch[1] : '0';

          // =============================================
          // SHEET NAME: Determine which billing month sheet to use
          // =============================================

          // Parse date string (DD/MM/YYYY) into a Date object
          const parts = purchaseDateStr.split('/');
          let dateObject = new Date(parts[2], parts[1] - 1, parts[0]);
          let billingMonth = dateObject.getMonth();
          let billingYear = dateObject.getFullYear();

          // ⚙️ CUSTOMIZE: Change 20 to your billing cycle cutoff day.
          // Purchases on or after this day are billed in the NEXT month.
          // e.g., if your cycle closes on the 25th, change 20 to 25.
          if (dateObject.getDate() >= 20) {
            billingMonth++;
            if (billingMonth > 11) {
              billingMonth = 0;
              billingYear++;
            }
          }

          // ⚙️ CUSTOMIZE: Change month names to your preferred language.
          // English: ['January','February','March','April','May','June','July','August','September','October','November','December']
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

          // ⚙️ CUSTOMIZE: Change the prefix 'Facturacion' to your preferred sheet name format.
          // e.g., 'Billing ' for English → "Billing March 2026"
          const sheetName =
            'Facturacion ' + monthNames[billingMonth] + ' ' + billingYear;

          let sheet = spreadsheet.getSheetByName(sheetName);

          // =============================================
          // SHEET CREATION: Create a new sheet if it doesn't exist
          // =============================================
          if (!sheet) {
            console.log('--- CREANDO HOJA NUEVA: ' + sheetName + ' ---');
            sheet = spreadsheet.insertSheet(sheetName);
            SpreadsheetApp.flush();
            console.log('Paso 1: Hoja insertada correctamente.');

            // Wait to ensure the sheet is fully created before writing
            Utilities.sleep(2000);

            // ⚙️ CUSTOMIZE: Change header names to your preferred language or column labels.
            // English example: ['Date','Merchant','Domestic','International','Installments','Card','Total Domestic','Total International']
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
            sheet.setFrozenRows(1); // Freeze header row for scrolling
            SpreadsheetApp.flush();
            console.log('Paso 2: Encabezados puestos.');

            Utilities.sleep(2000);

            // Add SUM formulas for totals — set one at a time to avoid batch failures
            console.log('Paso 3: Intentando poner fórmulas...');
            sheet.getRange('G2').setFormula('=SUM(C2:C1000)'); // Total domestic (CLP)
            SpreadsheetApp.flush();
            console.log('Subpaso 3.1: Fórmula Nacional lista.');

            sheet.getRange('H2').setFormula('=SUM(D2:D1000)'); // Total international (USD)
            SpreadsheetApp.flush();
            console.log('Subpaso 3.2: Fórmula Internacional lista.');

            // ⚙️ CUSTOMIZE: Change currency formats to match your local currency.
            // e.g., EUR: '€#,##0.00', GBP: '£#,##0.00', MXN: '$#,##0.00'
            sheet.getRange('G2').setNumberFormat('$#,##0'); // CLP format
            sheet.getRange('H2').setNumberFormat('USD #,##0.00'); // USD format
            console.log('Paso 4: Formatos aplicados.');
          }

          // =============================================
          // DATA WRITING: Append the expense row to the sheet
          // =============================================
          console.log('Paso 5: Calculando fila destino...');

          // Find the next empty row after existing data
          let lastRow = sheet.getLastRow();
          let targetRow = lastRow + 1;

          // Ensure we never overwrite the header row (row 1)
          if (targetRow < 2) targetRow = 2;

          console.log('Fila destino calculada: ' + targetRow);

          // Write the 6 data columns: Date, Merchant, Domestic, International, Installments, Visa
          sheet
            .getRange(targetRow, 1, 1, 6)
            .setValues([
              [
                purchaseDateStr,
                merchant,
                domesticAmount,
                internationalAmount,
                installments,
                visa,
              ],
            ]);

          // Highlight cancellation rows with a light red background
          if (isCancellation) {
            sheet.getRange(targetRow, 1, 1, 6).setBackground('#fff0f0');
          }

          // Mark the email as read so it won't be processed again
          message.markRead();
          SpreadsheetApp.flush();
          console.log('✅ Proceso completado para: ' + merchant);
        } catch (e) {
          console.error('❌ ERROR DETALLADO: ' + e.message);
          console.error('Línea del error: ' + e.stack);
        }
      }
    }
  }
}
