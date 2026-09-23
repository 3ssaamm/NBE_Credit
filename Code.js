/**
 * This is the FINAL refined script. 
 * - Handles the 29th Cutoff for BOTH One-Time and Installment payments.
 * - One-time payments: Due next month (unless bought on 30/31).
 * - Installments: First payment 55 days after purchase (respecting 29th cutoff).
 */
function updateLiveDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transactionsSheet = ss.getSheetByName("Transactions");
  const installmentsSheet = ss.getSheetByName("Installments");
  const debtsSheet = ss.getSheetByName("Monthly Debts");
  const historySheet = ss.getSheetByName("Payment History");

  const paidMonthsRange = historySheet.getRange("A2:A" + historySheet.getLastRow());
  const paidMonthsSet = new Set(paidMonthsRange.getDisplayValues().flat().filter(String));

  const allDebts = {};

  // --- 1. Process One-Time Payments ---
  const transactionsData = transactionsSheet.getDataRange().getValues().slice(1);
  transactionsData.forEach(row => {
    const dateValue = row[1], personRaw = row[3], amount = parseFloat(row[4]);
    if (dateValue && personRaw && typeof personRaw === 'string' && personRaw.trim() !== '' && !isNaN(amount)) {
      const person = personRaw.trim().toLowerCase();
      const purchaseDate = new Date(dateValue);
      const purchaseDay = purchaseDate.getDate();

      let dueDate;
      if (purchaseDay <= 29) {
        dueDate = new Date(purchaseDate.getFullYear(), purchaseDate.getMonth() + 1, 1);
      } else {
        dueDate = new Date(purchaseDate.getFullYear(), purchaseDate.getMonth() + 2, 1);
      }

      const monthKey = Utilities.formatDate(dueDate, ss.getSpreadsheetTimeZone(), "MMMM yyyy");
      if (!allDebts[monthKey]) allDebts[monthKey] = { total: 0, people: {} };
      allDebts[monthKey].total += amount;
      allDebts[monthKey].people[person] = (allDebts[monthKey].people[person] || 0) + amount;
    }
  });

  // --- 2. Process Installments ---
  const installmentsData = installmentsSheet.getRange("A2:K" + installmentsSheet.getLastRow()).getValues();
  const newInstallmentStatuses = [];

  installmentsData.forEach(row => {
    const purchaseDate = new Date(row[1]), durationMonths = parseInt(row[4]), emi = parseFloat(row[7]), payerRaw = row[8];

    if (isNaN(durationMonths) || isNaN(emi) || emi === 0 || !payerRaw) {
      newInstallmentStatuses.push(["", ""]);
      return;
    }
    
    const person = payerRaw.toString().trim().toLowerCase();
    
    // Calculate 55th day
    let firstPayDate = new Date(purchaseDate);
    firstPayDate.setDate(purchaseDate.getDate() + 55);

    // Apply 29th cutoff rule to the 55th day result
    let startMonthOffset = 0;
    if (firstPayDate.getDate() > 29) {
      startMonthOffset = 1; // Move to next month if it lands on 30th/31st
    }

    let paymentsMade = 0;
    for (let i = 0; i < durationMonths; i++) {
      // Calculate the month this specific installment is due
      const dueMonthDate = new Date(firstPayDate.getFullYear(), firstPayDate.getMonth() + i + startMonthOffset, 1);
      const monthKey = Utilities.formatDate(dueMonthDate, ss.getSpreadsheetTimeZone(), "MMMM yyyy");
      
      if (paidMonthsSet.has(monthKey)) paymentsMade++;

      if (!allDebts[monthKey]) allDebts[monthKey] = { total: 0, people: {} };
      allDebts[monthKey].total += emi;
      allDebts[monthKey].people[person] = (allDebts[monthKey].people[person] || 0) + emi;
    }

    newInstallmentStatuses.push([paymentsMade, (paymentsMade >= durationMonths ? "Completed" : "Ongoing")]);
  });

  if (newInstallmentStatuses.length > 0) {
    installmentsSheet.getRange("J2:K" + (newInstallmentStatuses.length + 1)).setValues(newInstallmentStatuses);
  }

  // --- 3. Display Report ---
  debtsSheet.clear();
  let currentRow = 1;
  const months = Object.keys(allDebts).sort((a, b) => new Date(a) - new Date(b));

  const today = new Date(), dayOfMonth = today.getDate();
  const currentMonthKey = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), "MMMM yyyy");
  const colors = { paid: "#d9ead3", due: "#fff2cc", header: "#f3f3f3" };

  months.forEach(month => {
    const data = allDebts[month];
    const isPaid = paidMonthsSet.has(month);
    const isDue = (month === currentMonthKey && dayOfMonth <= 25 && !isPaid);
    let blockColor = isPaid ? colors.paid : (isDue ? colors.due : null);

    const titleRange = debtsSheet.getRange(currentRow, 1, 1, 2).merge();
    titleRange.setValue(month).setFontWeight("bold").setFontSize(14).setBackground(isPaid ? colors.paid : colors.header);
    if (isPaid) titleRange.setFontLine("line-through");
    currentRow++;
    
    const rows = [["Total Bill:", data.total.toFixed(2)], ["Person", "Owes"]];
    Object.keys(data.people).sort().forEach(p => {
      rows.push([p.charAt(0).toUpperCase() + p.slice(1), data.people[p].toFixed(2)]);
    });

    const range = debtsSheet.getRange(currentRow, 1, rows.length, 2);
    range.setValues(rows);
    debtsSheet.getRange(currentRow, 1, 1, 2).setFontWeight("bold");
    debtsSheet.getRange(currentRow + 1, 1, 1, 2).setFontStyle("italic");
    if (blockColor) range.setBackground(blockColor);
    if (isPaid) range.setFontLine("line-through");

    currentRow += rows.length + 1;
  });
}