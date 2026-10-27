// Image 2 transactions (rows 302 to 328):
const txs = [
  { row: 302, date: "2026-09-30", payer: "Mido", amount: 1320.00 },
  { row: 303, date: "2026-10-01", payer: "Mai", amount: 89.96 },
  { row: 304, date: "2026-10-02", payer: "Mido", amount: 1447.80 },
  { row: 305, date: "2026-10-03", payer: "Mido", amount: 570.00 },
  { row: 306, date: "2026-10-03", payer: "Zoza", amount: 523.60 },
  { row: 307, date: "2026-10-03", payer: "Mido", amount: 2316.00 },
  { row: 308, date: "2026-10-04", payer: "Muhanad", amount: 135.00 },
  { row: 309, date: "2026-10-06", payer: "Mido", amount: 80.00 },
  { row: 310, date: "2026-10-08", payer: "Mido", amount: 333.10 },
  { row: 311, date: "2026-10-08", payer: "Dad", amount: 200.00 },
  { row: 312, date: "2026-10-08", payer: "Mum", amount: 200.00 },
  { row: 313, date: "2026-10-08", payer: "Muhanad", amount: 100.00 },
  { row: 314, date: "2026-10-04", payer: "Muhanad", amount: 80.00 },
  { row: 315, date: "2026-10-06", payer: "Muhanad", amount: 15.00 },
  { row: 316, date: "2026-10-10", payer: "Muhanad", amount: 30.00 },
  { row: 317, date: "2026-10-11", payer: "Muhanad", amount: 30.00 },
  { row: 318, date: "2026-10-14", payer: "Mai", amount: 100.00 },
  { row: 319, date: "2026-10-14", payer: "Dad", amount: 200.00 },
  { row: 320, date: "2026-10-14", payer: "Mido", amount: 367.00 },
  { row: 321, date: "2026-10-15", payer: "Nourween", amount: 5100.00 },
  { row: 322, date: "2026-10-15", payer: "Nourween", amount: 5100.00 },
  { row: 323, date: "2026-10-15", payer: "Hager", amount: 2222.00 },
  { row: 324, date: "2026-10-14", payer: "Muhanad", amount: 99.00 },
  { row: 325, date: "2026-10-15", payer: "Muhanad", amount: 125.00 },
  { row: 326, date: "2026-10-15", payer: "Muhanad", amount: 25.00 },
  { row: 327, date: "2026-10-15", payer: "Mai", amount: 65.00 },
  { row: 328, date: "2026-10-23", payer: "Mido", amount: 164.99 }
];

// November 2026 base installments from Image 1:
const installments = {
  "Abdo": 1110.88,
  "Dad": 1090.33,
  "Hager": 0.00,
  "Hoda": 0.00,
  "Mai": 1366.38,
  "Mido": 1484.61,
  "Muhanad": 106.17,
  "Mum": 578.41,
  "Nevin": 0.00,
  "Nourween": 0.00,
  "Zoza": 0.00
};

const purchByPayer = {};
txs.forEach(t => {
  purchByPayer[t.payer] = (purchByPayer[t.payer] || 0) + t.amount;
});

console.log("Purchases in Nov 2026 cycle:");
console.table(purchByPayer);

const totalNovForecast = {};
let grandTotal = 0;
Object.keys(installments).forEach(p => {
  const inst = installments[p];
  const purch = purchByPayer[p] || 0;
  const tot = inst + purch;
  totalNovForecast[p] = {
    Installments: inst.toFixed(2),
    Purchases: purch.toFixed(2),
    ForecastNovTotal: tot.toFixed(2)
  };
  grandTotal += tot;
});

console.table(totalNovForecast);
console.log("Grand Total November Forecast:", grandTotal.toFixed(2), "EGP");
