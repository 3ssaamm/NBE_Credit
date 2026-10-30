const fs = require('fs');

// Let's inspect test_nov_calc.js or any store files
console.log("Searching for 558.23 in scratch directory...");
const files = fs.readdirSync('scratch');
files.forEach(f => {
  const c = fs.readFileSync('scratch/' + f, 'utf8');
  if (c.includes('558') || c.includes('558.23')) {
    console.log(`Found in scratch/${f}`);
  }
});

