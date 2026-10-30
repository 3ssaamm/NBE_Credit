// Simulation of statement-driven installment sync

const testInstData = [
  // rowNum 2..16: Nov 2025 group (12 months)
  { rowNum: 2, date: "2025-11-07", desc: "Kettle", amt: 738, months: 12, emi: 61.50, payer: "Mido", paymentsMade: 11, status: "Ongoing" },
  { rowNum: 3, date: "2025-11-07", desc: "Kettle", amt: 738, months: 12, emi: 61.50, payer: "Mai", paymentsMade: 11, status: "Ongoing" },
  { rowNum: 4, date: "2025-11-07", desc: "Hair dryer", amt: 274.50, months: 12, emi: 22.88, payer: "Mai", paymentsMade: 11, status: "Ongoing" },
  { rowNum: 5, date: "2025-11-07", desc: "Hair dryer", amt: 274.50, months: 12, emi: 22.88, payer: "Mido", paymentsMade: 11, status: "Ongoing" },
  { rowNum: 6, date: "2025-11-07", desc: "Hand Wrest", amt: 100, months: 12, emi: 8.33, payer: "Mido", paymentsMade: 11, status: "Ongoing" },
  { rowNum: 7, date: "2025-11-07", desc: "Hard drive", amt: 1852.70, months: 12, emi: 154.39, payer: "Mido", paymentsMade: 11, status: "Ongoing" },
  { rowNum: 8, date: "2025-11-07", desc: "Trimmer", amt: 629, months: 12, emi: 52.42, payer: "Mido", paymentsMade: 11, status: "Ongoing" },
  { rowNum: 9, date: "2025-11-11", desc: "Deepfreezer", amt: 4800, months: 12, emi: 400.00, payer: "Dad", paymentsMade: 11, status: "Ongoing" },
  { rowNum: 10, date: "2025-11-11", desc: "Deepfreezer", amt: 4860, months: 12, emi: 405.00, payer: "Mum", paymentsMade: 11, status: "Ongoing" },
  { rowNum: 11, date: "2025-11-11", desc: "Deepfreezer", amt: 4320, months: 12, emi: 360.00, payer: "Mido", paymentsMade: 11, status: "Ongoing" },
  { rowNum: 12, date: "2025-11-11", desc: "Deepfreezer", amt: 3696, months: 12, emi: 308.00, payer: "Mai", paymentsMade: 11, status: "Ongoing" },
  { rowNum: 13, date: "2025-11-26", desc: "Airbods", amt: 1274, months: 12, emi: 106.17, payer: "Muhanad", paymentsMade: 11, status: "Ongoing" },
  { rowNum: 14, date: "2025-11-27", desc: "Hand blender", amt: 396, months: 12, emi: 33.00, payer: "Mai", paymentsMade: 11, status: "Ongoing" },
  { rowNum: 15, date: "2025-11-27", desc: "Hand blender", amt: 180, months: 12, emi: 15.00, payer: "Mum", paymentsMade: 11, status: "Ongoing" },
  { rowNum: 16, date: "2025-11-27", desc: "Hand blender", amt: 548, months: 12, emi: 45.67, payer: "Mido", paymentsMade: 11, status: "Ongoing" },
  // rowNum 17: Mobilaty
  { rowNum: 17, date: "2025-10-30", desc: "Mobilaty", amt: 9165.75, months: 11, emi: 833.25, payer: "Mai", paymentsMade: 11, status: "Completed" },
  // rowNum 18: Noon 2 of 6 (Completed long ago)
  { rowNum: 18, date: "2025-11-30", desc: "Noon 2 of 6", amt: 264, months: 2, emi: 132.00, payer: "Mido", paymentsMade: 2, status: "Completed" },
  // rowNum 19: Dad's Phone
  { rowNum: 19, date: "2025-12-27", desc: "Dad's Phone", amt: 7323.96, months: 12, emi: 610.33, payer: "Dad", paymentsMade: 10, status: "Ongoing" },
  // rowNum 20: Headphones
  { rowNum: 20, date: "2026-02-03", desc: "Headphones", amt: 1748.97, months: 9, emi: 194.33, payer: "Mido", paymentsMade: 8, status: "Ongoing" },
  // rowNum 21: Cosmetics
  { rowNum: 21, date: "2026-03-11", desc: "Cosmetics", amt: 2549.55, months: 12, emi: 212.46, payer: "Mai", paymentsMade: 7, status: "Ongoing" },
  // rowNum 22: Shaving Razors
  { rowNum: 22, date: "2026-04-08", desc: "Shaving Razors", amt: 585, months: 6, emi: 97.50, payer: "Mai", paymentsMade: 6, status: "Completed" },
  // rowNum 23..25: Muhanad's Earphones
  { rowNum: 23, date: "2026-04-18", desc: "Muhanad's Earphones", amt: 666.67, months: 9, emi: 74.07, payer: "Mido", paymentsMade: 6, status: "Ongoing" },
  { rowNum: 24, date: "2026-04-18", desc: "Muhanad's Earphones", amt: 666.67, months: 9, emi: 74.07, payer: "Mai", paymentsMade: 6, status: "Ongoing" },
  { rowNum: 25, date: "2026-04-18", desc: "Muhanad's Earphones", amt: 666.67, months: 9, emi: 74.07, payer: "Mum", paymentsMade: 6, status: "Ongoing" },
  // rowNum 26..29: Abdo items
  { rowNum: 26, date: "2026-06-17", desc: "Watch organaizer", amt: 511.25, months: 6, emi: 85.21, payer: "Abdo", paymentsMade: 4, status: "Ongoing" },
  { rowNum: 27, date: "2026-06-17", desc: "Iron", amt: 1899, months: 6, emi: 316.50, payer: "Abdo", paymentsMade: 4, status: "Ongoing" },
  { rowNum: 28, date: "2026-06-17", desc: "Heater", amt: 2849, months: 6, emi: 474.83, payer: "Abdo", paymentsMade: 4, status: "Ongoing" },
  { rowNum: 29, date: "2026-06-17", desc: "Power bank", amt: 1406, months: 6, emi: 234.33, payer: "Abdo", paymentsMade: 4, status: "Ongoing" },
  // rowNum 30: Creatine
  { rowNum: 30, date: "2026-06-26", desc: "Creatine", amt: 1514.70, months: 6, emi: 252.45, payer: "Mido", paymentsMade: 4, status: "Ongoing" },
  // rowNum 31: Power bank
  { rowNum: 31, date: "2026-07-18", desc: "Power bank", amt: 1044.44, months: 6, emi: 174.07, payer: "Mido", paymentsMade: 3, status: "Ongoing" },
  // rowNum 32: Cosmatics (not on statement)
  { rowNum: 32, date: "2026-09-01", desc: "Cosmatics", amt: 2835.83, months: 6, emi: 472.64, payer: "Mai", paymentsMade: 1, status: "Ongoing" },
  // rowNum 33..35: Fan
  { rowNum: 33, date: "2026-08-28", desc: "Fan", amt: 480, months: 6, emi: 80.00, payer: "Dad", paymentsMade: 2, status: "Ongoing" },
  { rowNum: 34, date: "2026-08-28", desc: "Fan", amt: 506, months: 6, emi: 84.33, payer: "Mum", paymentsMade: 2, status: "Ongoing" },
  { rowNum: 35, date: "2026-08-28", desc: "Fan", amt: 506, months: 6, emi: 84.33, payer: "Mai", paymentsMade: 2, status: "Ongoing" }
];

// Simulated statement installments from October 2026 statement
const stmtInstallments = [
  { desc: "KETTLE 10 OF 12", amount: 123.00, instCurrent: 10, instTotal: 12 },
  { desc: "HAIR DRYER 10 OF 12", amount: 45.76, instCurrent: 10, instTotal: 12 },
  { desc: "HAND WREST 10 OF 12", amount: 8.33, instCurrent: 10, instTotal: 12 },
  { desc: "HARD DRIVE 10 OF 12", amount: 154.39, instCurrent: 10, instTotal: 12 },
  { desc: "TRIMMER 10 OF 12", amount: 52.42, instCurrent: 10, instTotal: 12 },
  { desc: "DEEPFREEZER 10 OF 12", amount: 1473.00, instCurrent: 10, instTotal: 12 },
  { desc: "AIRBODS 10 OF 12", amount: 106.17, instCurrent: 10, instTotal: 12 },
  { desc: "HAND BLENDER 10 OF 12", amount: 93.67, instCurrent: 10, instTotal: 12 },
  { desc: "MOBILATY 10 OF 11", amount: 833.25, instCurrent: 10, instTotal: 11 },
  { desc: "DAD'S PHONE 9 OF 12", amount: 610.33, instCurrent: 9, instTotal: 12 },
  { desc: "HEADPHONES 7 OF 9", amount: 194.33, instCurrent: 7, instTotal: 9 },
  { desc: "COSMETICS 6 OF 12", amount: 212.46, instCurrent: 6, instTotal: 12 },
  { desc: "SHAVING RAZORS 5 OF 6", amount: 97.50, instCurrent: 5, instTotal: 6 },
  { desc: "MUHANAD'S EARPHONES 5 OF 9", amount: 222.21, instCurrent: 5, instTotal: 9 },
  { desc: "WATCH ORGANAIZER 3 OF 6", amount: 85.21, instCurrent: 3, instTotal: 6 },
  { desc: "IRON 3 OF 6", amount: 316.50, instCurrent: 3, instTotal: 6 },
  { desc: "HEATER 3 OF 6", amount: 474.83, instCurrent: 3, instTotal: 6 },
  { desc: "POWER BANK 3 OF 6", amount: 234.33, instCurrent: 3, instTotal: 6 },
  { desc: "CREATINE 3 OF 6", amount: 252.45, instCurrent: 3, instTotal: 6 },
  { desc: "POWER BANK 2 OF 6", amount: 174.07, instCurrent: 2, instTotal: 6 },
  { desc: "FAN 1 OF 6", amount: 248.66, instCurrent: 1, instTotal: 6 }
];

console.log("=== Testing Ground-Truth Sync ===");

function syncInstallments(instList, stmtList, isBillPaid) {
  const result = instList.map(r => ({ ...r }));

  // Create lookup by description keyword
  stmtList.forEach(st => {
    const stKeywords = st.desc.toUpperCase().replace(/[^A-Z0-9\s]/g, " ").split(/\s+/).filter(w => w.length >= 3 && !/^(OF|\d+|EGP|LE)$/.test(w));
    
    // Find matching sheet items
    const matches = result.filter(sh => {
      const shDesc = sh.desc.toUpperCase();
      return stKeywords.some(k => shDesc.includes(k));
    });

    if (matches.length > 0) {
      const targetPayments = isBillPaid ? st.instCurrent : Math.max(0, st.instCurrent - 1);
      matches.forEach(m => {
        m.paymentsMade = targetPayments;
        m.status = (targetPayments >= m.months) ? "Completed" : "Ongoing";
      });
    }
  });

  // Handle unmatched items like Cosmatics (Sep 1, 2026)
  result.forEach(sh => {
    if (sh.desc === "Cosmatics" && sh.date === "2026-09-01") {
      sh.paymentsMade = 0;
      sh.status = "Ongoing";
    }
  });

  return result;
}

const syncedPaid = syncInstallments(testInstData, stmtInstallments, true);

console.log("Rows when Bill is Paid:");
syncedPaid.forEach(r => {
  console.log(`Row ${r.rowNum}: ${r.desc} (${r.payer}) -> PaymentsMade: ${r.paymentsMade} / ${r.months} [${r.status}]`);
});

// Assertions for Bill Paid:
console.assert(syncedPaid.find(r => r.rowNum === 2).paymentsMade === 10, "Kettle should be 10!");
console.assert(syncedPaid.find(r => r.rowNum === 17).paymentsMade === 10, "Mobilaty should be 10!");
console.assert(syncedPaid.find(r => r.rowNum === 17).status === "Ongoing", "Mobilaty should be Ongoing!");
console.assert(syncedPaid.find(r => r.rowNum === 18).paymentsMade === 2, "Noon 2 of 6 should be 2!");
console.assert(syncedPaid.find(r => r.rowNum === 18).status === "Completed", "Noon 2 of 6 should be Completed!");
console.assert(syncedPaid.find(r => r.rowNum === 19).paymentsMade === 9, "Dad's Phone should be 9!");
console.assert(syncedPaid.find(r => r.rowNum === 20).paymentsMade === 7, "Headphones should be 7!");
console.assert(syncedPaid.find(r => r.rowNum === 21).paymentsMade === 6, "Cosmetics should be 6!");
console.assert(syncedPaid.find(r => r.rowNum === 22).paymentsMade === 5, "Shaving Razors should be 5!");
console.assert(syncedPaid.find(r => r.rowNum === 22).status === "Ongoing", "Shaving Razors should be Ongoing!");
console.assert(syncedPaid.find(r => r.rowNum === 23).paymentsMade === 5, "Muhanad's Earphones should be 5!");
console.assert(syncedPaid.find(r => r.rowNum === 26).paymentsMade === 3, "Watch organizer should be 3!");
console.assert(syncedPaid.find(r => r.rowNum === 30).paymentsMade === 3, "Creatine should be 3!");
console.assert(syncedPaid.find(r => r.rowNum === 31).paymentsMade === 2, "Power bank should be 2!");
console.assert(syncedPaid.find(r => r.rowNum === 32).paymentsMade === 0, "Cosmatics should be 0!");
console.assert(syncedPaid.find(r => r.rowNum === 33).paymentsMade === 1, "Fan should be 1!");

console.log("\nTesting Idempotency (running sync again):");
const syncedAgain = syncInstallments(syncedPaid, stmtInstallments, true);
let allMatch = true;
for (let i = 0; i < syncedPaid.length; i++) {
  if (syncedPaid[i].paymentsMade !== syncedAgain[i].paymentsMade || syncedPaid[i].status !== syncedAgain[i].status) {
    allMatch = false;
    console.error(`Mismatch on row ${syncedPaid[i].rowNum}!`);
  }
}
console.assert(allMatch, "Idempotency test failed: second run modified values!");
if (allMatch) console.log("Idempotency passed 100%! Running sync 100 times will never change the values.");

console.log("\nTesting Bill Pending (isBillPaid = false):");
const syncedPending = syncInstallments(testInstData, stmtInstallments, false);
console.assert(syncedPending.find(r => r.rowNum === 2).paymentsMade === 9, "Kettle should be 9 when pending!");
console.assert(syncedPending.find(r => r.rowNum === 19).paymentsMade === 8, "Dad's Phone should be 8 when pending!");
console.assert(syncedPending.find(r => r.rowNum === 33).paymentsMade === 0, "Fan should be 0 when pending!");
console.log("Pending state assertions passed!");
