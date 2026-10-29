const testInstData = [
  // 0: Kettle (Mido)
  [1, new Date("2025-11-07"), "Kettle", 738.00, 12, 0, 0, 61.50, "Mido", 9, "Ongoing"],
  // 1: Kettle (Mai)
  [2, new Date("2025-11-07"), "Kettle", 738.00, 12, 0, 0, 61.50, "Mai", 9, "Ongoing"],
  // 2: Noon 2 of 6 (Completed)
  [17, new Date("2025-11-30"), "Noon 2 of 6", 264.00, 2, 0, 0, 132.00, "Mido", 2, "Completed"],
  // 3: Dad's Phone
  [18, new Date("2025-12-27"), "Dad's Phone", 7323.96, 12, 0, 0, 610.33, "Dad", 8, "Ongoing"],
  // 4: Creatine
  [29, new Date("2026-06-26"), "Creatine", 1514.70, 6, 0, 0, 252.45, "Mido", 2, "Ongoing"],
  // 5: Power bank (Jul 18, 2026)
  [30, new Date("2026-07-18"), "Power bank", 1044.44, 6, 0, 0, 174.07, "Mido", 1, "Ongoing"],
  // 6: Fan (Dad, Aug 28, 2026)
  [33, new Date("2026-08-28"), "Fan", 480.00, 6, 0, 0, 80.00, "Dad", 0, "Ongoing"],
  // 7: Test almost completed plan (5 of 6)
  [99, new Date("2026-03-15"), "Test Ending Plan", 600.00, 6, 0, 0, 100.00, "Mai", 5, "Ongoing"],
  // 8: Future installment (Nov 2026)
  [100, new Date("2026-11-15"), "Future Plan", 1000.00, 6, 0, 0, 166.67, "Mido", 0, "Ongoing"]
];

const paymentDate = new Date("2026-10-25");
const numRows = testInstData.length;
const paymentsCol = testInstData.map(r => [r[9]]);
const statusCol = testInstData.map(r => [r[10]]);
const audit = [];

for (let r = 0; r < numRows; r++) {
  const row = testInstData[r];
  const pDate = row[1];
  const desc = row[2];
  const months = parseInt(row[4], 10);
  const emi = parseFloat(row[7]) || 0;
  const payer = String(row[8] || "").trim();
  const paymentsMade = parseInt(row[9], 10) || 0;
  const currentStatus = String(row[10] || "").trim();

  if (months <= 0 || emi <= 0 || !payer) continue;
  if (paymentsMade >= months || currentStatus.toLowerCase() === "completed") continue;

  // Don't increment future installments that start after payment date
  if (pDate && paymentDate && pDate.getTime() > paymentDate.getTime() + 86400000) {
    continue;
  }

  const newPayments = paymentsMade + 1;
  const newStatus = (newPayments >= months) ? "Completed" : (currentStatus || "Ongoing");

  audit.push({
    rowIndex: r,
    rowNum: r + 2,
    desc: desc,
    payer: payer,
    prevPayments: paymentsMade,
    newPayments: newPayments,
    prevStatus: currentStatus,
    newStatus: newStatus
  });

  paymentsCol[r][0] = newPayments;
  statusCol[r][0] = newStatus;
}

console.log(`Updated ${audit.length} installments:`);
console.table(audit);

console.assert(audit.length === 7, "Exactly 7 installments must be updated (Noon 2 of 6 and Future Plan excluded)!");
console.assert(paymentsCol[0][0] === 10, "Kettle 1 must be 10!");
console.assert(paymentsCol[1][0] === 10, "Kettle 2 must be 10!");
console.assert(paymentsCol[2][0] === 2 && statusCol[2][0] === "Completed", "Noon 2 of 6 must remain 2 Completed!");
console.assert(paymentsCol[6][0] === 1, "Fan must be 1!");
console.assert(paymentsCol[7][0] === 6 && statusCol[7][0] === "Completed", "Ending plan must be 6 Completed!");
console.assert(paymentsCol[8][0] === 0, "Future plan must remain 0!");
console.log("All assertions passed!");

