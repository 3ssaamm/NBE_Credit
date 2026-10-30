const fs = require('fs');

// November 2026 base installments:
const installments = {
  "Mido": [
    { desc: "Kettle", emi: 61.50 },
    { desc: "Hair dryer", emi: 22.88 },
    { desc: "Hand Wrest", emi: 8.33 },
    { desc: "Hard drive", emi: 154.39 },
    { desc: "Trimmer", emi: 52.42 },
    { desc: "Deepfreezer", emi: 360.00 },
    { desc: "Hand blender", emi: 45.67 },
    { desc: "Headphones", emi: 194.33 },
    { desc: "Muhanad's Earphones", emi: 74.07 },
    { desc: "Creatine", emi: 252.45 },
    { desc: "Power bank", emi: 174.07 },
    { desc: "Fan", emi: 84.50 }
  ]
};

const total = installments.Mido.reduce((a, b) => a + b.emi, 0);
console.log("Total Mido Installments in Installments sheet:", total.toFixed(2));
console.log("Difference from 1369.52:", (total - 1369.52).toFixed(2));

// Which combination of EMIs equals 115.09?
const emis = installments.Mido.map(x => x.emi);
for (let i = 0; i < emis.length; i++) {
  if (Math.abs(emis[i] - 115.09) < 0.1) console.log("Single match:", installments.Mido[i]);
  for (let j = i + 1; j < emis.length; j++) {
    if (Math.abs(emis[i] + emis[j] - 115.09) < 0.1) {
      console.log("Double match:", installments.Mido[i], installments.Mido[j]);
    }
  }
}

