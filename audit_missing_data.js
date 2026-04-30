const XLSX = require('xlsx');
const path = require('path');
const { processMpsFile } = require('./extractor');

const targetFile = 'c:/Users/i0215099/Desktop/MPS_UPDATE/MPS2603-1.xlsx';

async function runAudit() {
    console.log('--- MPS Data Full Inspection (전수검사) ---');
    console.log('Target File:', targetFile);

    try {
        const result = await processMpsFile(targetFile);
        const { finalResults, unusedData } = result;

        console.log('\n[Summary]');
        console.log('Total Matched Rows:', finalResults.length);
        console.log('Total Unused/Unmapped Rows:', unusedData.length);

        if (unusedData.length > 0) {
            console.log('\n[Unused Data Pool - Problematic Rows]');
            console.log('Month | Site | Group | ModelCode | ProductName | Category');
            console.log('---------------------------------------------------------');
            
            // Group by category for better reporting
            const grouped = {};
            unusedData.forEach(row => {
                const cat = row.Category || 'Unknown';
                if (!grouped[cat]) grouped[cat] = [];
                grouped[cat].push(row);
            });

            for (const cat in grouped) {
                console.log(`\nCategory: ${cat} (${grouped[cat].length} rows)`);
                // Show first 20 examples
                grouped[cat].slice(0, 20).forEach(row => {
                    console.log(`${row.Month} | ${row.Site} | ${row.Group} | ${row.ModelCode || '(Empty)'} | ${row.ProductName} | ${row.Category}`);
                });
                if (grouped[cat].length > 20) {
                    console.log(`... and ${grouped[cat].length - 20} more rows.`);
                }
            }
        } else {
            console.log('\n✨ No problematic rows found. All data matched successfully!');
        }

    } catch (err) {
        console.error('Audit failed:', err);
    }
}

runAudit();
