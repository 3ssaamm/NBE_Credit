// Mock utilities from Code.js
function parseDateValue(rawDate, tz) {
  if (!rawDate) return null;
  if (rawDate instanceof Date) return rawDate;
  if (typeof rawDate === "string") {
    const d = new Date(rawDate);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function getBillingCycleForInstallment(purchaseDate, installmentIndex, tz) {
  if (!purchaseDate) return null;
  const month = purchaseDate.getMonth();
  const year = purchaseDate.getFullYear();

  let cycleMonth = month + 1 + installmentIndex;
  let cycleYear = year + Math.floor(cycleMonth / 12);
  cycleMonth = ((cycleMonth % 12) + 12) % 12;

  let dueMonth = cycleMonth + 1;
  let dueYear = cycleYear;
  if (dueMonth > 11) {
    dueMonth = 0;
    dueYear++;
  }

  const cycleDate = new Date(dueYear, dueMonth, 1, 12, 0, 0);
  const pad = n => String(n).padStart(2, "0");
  const cycleSortKey = `${dueYear}-${pad(dueMonth + 1)}`;
  return { cycleSortKey };
}

// Simulated Installment Sheet Rows based on user's screenshot
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
  [99, new Date("2026-03-15"), "Test Ending Plan", 600.00, 6, 0, 0, 100.00, "Mai", 5, "Ongoing"]
];

const activeCycleKey = "2026-10"; // October 2026 due cycle (closed end of Sept 2026)

console.log("=== Testing Installment Increment Logic ===");
const audit = [];
const paymentsCol = testInstData.map(r => [r[9]]);
const statusCol = testInstData.map(r => [r[10]]);

for (let r = 0; r < testInstData.length; r++) {
  const row = testInstData[r];
  const pDate = row[1];
  const desc = row[2];
  const months = row[4];
  const emi = row[7];
  const payer = row[8];
  const paymentsMade = row[9];
  const currentStatus = row[10];

  if (months <= 0 || emi <= 0 || !payer) continue;
  if (paymentsMade >= months || currentStatus.toLowerCase() === "completed") continue;

  let isDue = false;
  if (pDate) {
    const cyc = getBillingCycleForInstallment(pDate, paymentsMade, "Africa/Cairo");
    if (cyc && cyc.cycleSortKey <= activeCycleKey) {
      isDue = true;
    }
  }

  if (isDue) {
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
}

console.log(`Updated ${audit.length} installments:`);
console.table(audit);

// Verification tests
console.assert(testInstData[2][9] === 2 && testInstData[2][10] === "Completed", "Noon 2 of 6 was Completed and must not change!");
console.assert(paymentsCol[2][0] === 2, "Noon 2 of 6 paymentsCol must stay 2!");
console.assert(paymentsCol[0][0] === 10, "Kettle must increment from 9 to 10!");
console.assert(paymentsCol[6][0] === 1, "Fan must increment from 0 to 1!");
console.assert(paymentsCol[7][0] === 6 && statusCol[7][0] === "Completed", "Ending plan must become Completed!");

console.log("=== Testing Reversibility (Clear Date) ===");
audit.forEach(item => {
  const r = item.rowIndex;
  paymentsCol[r][0] = item.prevPayments;
  statusCol[r][0] = item.prevStatus;
});

console.assert(paymentsCol[0][0] === 9, "Kettle rolled back to 9!");
console.assert(paymentsCol[6][0] === 0, "Fan rolled back to 0!");
console.assert(paymentsCol[7][0] === 5 && statusCol[7][0] === "Ongoing", "Ending plan rolled back to 5 Ongoing!");
console.log("All unit tests passed successfully!");

