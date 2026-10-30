const items = [
  { desc: "Kettle", emi: 61.50 },
  { desc: "Hair dryer", emi: 22.88 },
  { desc: "Hand Wrest", emi: 8.33 },
  { desc: "Hard drive", emi: 154.39 },
  { desc: "Trimmer", emi: 52.42 },
  { desc: "Deepfreezer", emi: 360.00 },
  { desc: "Hand blender", emi: 45.67 },
  { desc: "Headphones", emi: 194.33 },
  { desc: "Muhanad Earphones", emi: 74.07 },
  { desc: "Creatine", emi: 252.45 },
  { desc: "Power bank", emi: 174.07 },
  { desc: "Fan", emi: 84.50 }
];

console.log("Looking for combination equaling 115.09:");
for (let i = 0; i < items.length; i++) {
  if (Math.abs(items[i].emi - 115.09) < 0.5) console.log("Single:", items[i]);
  for (let j = i + 1; j < items.length; j++) {
    if (Math.abs(items[i].emi + items[j].emi - 115.09) < 0.5) console.log("Double:", items[i].desc, items[j].desc, items[i].emi + items[j].emi);
    for (let k = j + 1; k < items.length; k++) {
      if (Math.abs(items[i].emi + items[j].emi + items[k].emi - 115.09) < 0.5) {
        console.log("Triple:", items[i].desc, items[j].desc, items[k].desc, items[i].emi + items[j].emi + items[k].emi);
      }
    }
  }
}

