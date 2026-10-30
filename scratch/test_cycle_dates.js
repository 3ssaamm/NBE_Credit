function getBillingCycleForPurchase(purchaseDate) {
  if (!purchaseDate) return null;
  let day = purchaseDate.getDate();
  let month = purchaseDate.getMonth();
  let year = purchaseDate.getFullYear();

  let cycleYear = year;
  let cycleMonth = month;

  const nextDay = new Date(year, month, day + 1, 12, 0, 0);
  const isLastDay = (nextDay.getMonth() !== month);

  if (isLastDay) {
    cycleMonth = month + 1;
    if (cycleMonth > 11) {
      cycleMonth = 0;
      cycleYear++;
    }
  }

  let dueYear = cycleYear;
  let dueMonth = cycleMonth + 1;
  if (dueMonth > 11) {
    dueMonth = 0;
    dueYear++;
  }
  const pad = n => String(n).padStart(2, '0');
  return `${dueYear}-${pad(dueMonth + 1)}`;
}

// Dates to test:
console.log("2026-08-30 ->", getBillingCycleForPurchase(new Date(2026, 7, 30)));
console.log("2026-08-31 (last day of Aug) ->", getBillingCycleForPurchase(new Date(2026, 7, 31)));
console.log("2026-09-01 ->", getBillingCycleForPurchase(new Date(2026, 8, 1)));
console.log("2026-09-29 ->", getBillingCycleForPurchase(new Date(2026, 8, 29)));
console.log("2026-09-30 (last day of Sep) ->", getBillingCycleForPurchase(new Date(2026, 8, 30)));
console.log("2026-10-01 ->", getBillingCycleForPurchase(new Date(2026, 9, 1)));

