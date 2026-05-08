const { processMpsFile } = require('../extractor');
const fs = require('fs');

// Load site master if possible, but for simple audit we can pass empty
const result = processMpsFile('MPS2604-1.xlsx');

console.log('Total Results:', result.finalResults.length);
console.log('Unmatched Count (Rows):', result.unusedData.length);

const unmatchedByModel = {};
result.unusedData.forEach(item => {
    const key = `${item.Model} (${item.Month})`;
    unmatchedByModel[key] = (unmatchedByModel[key] || 0) + item.Qty;
});

console.log('\n--- Top 10 Unmatched Models ---');
Object.entries(unmatchedByModel)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .forEach(([model, qty]) => {
        console.log(`${model}: ${qty}`);
    });

// Detailed check for the 34 items the user mentioned
if (result.unusedData.length === 34) {
    console.log('\nBingo! Found exactly 34 unmatched items.');
} else {
    console.log(`\nFound ${result.unusedData.length} unmatched items. (User said 34)`);
}

// Check month distribution in finalResults
const monthDist = {};
result.finalResults.forEach(r => {
    monthDist[r.Month] = (monthDist[r.Month] || 0) + r.Qty;
});
result.unusedData.forEach(r => {
    monthDist[r.Month] = (monthDist[r.Month] || 0) + r.Qty;
});
console.log('\n--- Month Distribution (All) ---');
console.log(monthDist);
