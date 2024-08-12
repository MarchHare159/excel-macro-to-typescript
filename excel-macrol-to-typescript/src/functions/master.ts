import * as Excel from 'exceljs';
import { rl, paths } from "../utils/utils";
import * as fs from 'fs';
import * as path from 'path';
import { parse } from 'csv-parse/sync';
import * as readline from 'readline';
const inputPath = paths.inputPath;

export async function checkScript() {
    const wb = new Excel.Workbook();
    const wbFanaticsPath = await selectFile("Choose Fanatics File");
    const wbTargetPath = await selectFile("Choose Target File");

    const wbFanatics = await wb.xlsx.readFile(wbFanaticsPath);
    const wbTarget = await wb.xlsx.readFile(wbTargetPath);

    const shtFanatics = wbFanatics.worksheets[0];
    const shtTarget = wbTarget.worksheets[0];

    const shtMaster = wb.addWorksheet('master');
    const shtFanNew = wb.addWorksheet('Fanatics New');
    const shtTNew = wb.addWorksheet('Target New');

    const fanaticsData = shtFanatics.getColumn(7).values.slice(1); // G column
    const targetData = shtTarget.getColumn(7).values.slice(1); // G column

    // Clear previous data in master sheet
    for (let row = 2; row <= 100000; row++) {
        shtMaster.getCell(`S${row}`).value = null;
        shtMaster.getCell(`T${row}`).value = null;
    }

    const arrDataFanatics = fanaticsData.filter(Boolean);
    const arrDataTarget = targetData.filter(Boolean);

    // Process Fanatics data
    let insertRowFan = 2;
    for (const strsku of arrDataFanatics) {
        const foundRow = shtFanatics.getColumn(2).values.findIndex(value => value === strsku);
        if (foundRow > -1 && shtFanatics.getCell(foundRow + 1, 4)?.value || 0 > 0) { // D column
            shtFanNew.addRow(shtFanatics.getRow(foundRow + 1).values);
        }
    }

    // Save Fanatics data
    const fanaticsFilePath = path.join(__dirname, "Fanatics checkscript export.csv");
    await wbFanatics.xlsx.writeFile(fanaticsFilePath);

    // Process Target data
    let insertRowTarget = 2;
    for (const strsku of arrDataTarget) {
        const foundRow = shtTarget.getColumn(2).values.findIndex(value => value === strsku);
        if (foundRow > -1 && shtTarget.getCell(foundRow + 1, 3)?.value || 0 > 0) { // C column
            shtTNew.addRow(shtTarget.getRow(foundRow + 1).values);
        }
    }

    // Save Target data
    const targetFilePath = path.join(__dirname, "Target checkscript export.csv");
    await wbTarget.xlsx.writeFile(targetFilePath);

    console.log("Done!");
}

export async function preAndAllocate() {
    await updateOnHand();
    await allocate();
    rl.close();
}

export async function allocateCSVFanAndHib() {
    const wbS = new Excel.Workbook();
    await wbS.xlsx.readFile(inputPath); // Change this to your workbook path
    const wbD = new Excel.Workbook();
    
    const shtS = wbS.getWorksheet('Fanatics');
    if (!shtS) {
        console.error("Sheet 'Fanatics' not found.");
        return;
    }

    // Copy the Fanatics sheet to the new workbook
    const copiedSheet = wbD.addWorksheet('Fanatics');
    shtS.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            copiedSheet.getCell(rowNumber, colNumber).value = cell.value;
        });
    });

    // Call UpdateFanSkus function (you'll need to implement this)
    await updateFanSkus(wbD, copiedSheet);

    const strsaveaspath = wbS.getWorksheet('Setup')?.getCell('B14').value as string;
    const strmacpath = wbS.getWorksheet('Setup')?.getCell('B8').value as string;

    const formattedDate = `${new Date().getFullYear()}${String(new Date().getMonth() + 1).padStart(2, '0')}${String(new Date().getDate()).padStart(2, '0')}_` +
                          `${String(new Date().getHours()).padStart(2, '0')}${String(new Date().getMinutes()).padStart(2, '0')}${String(new Date().getSeconds()).padStart(2, '0')}`;
    
    const strfile = `IB_LDM_${formattedDate}.csv`;
    const fullPath = path.join(strmacpath, strfile);
    
    // Save the new workbook as CSV
    await wbD.csv.writeFile(fullPath);

    // Rename the file (if necessary)
    const finalPath = path.join(strsaveaspath, strfile);
    await fs.rename(fullPath, finalPath, (err) => {
        if (err) {
            console.error("Error renaming file:", err);
            return;
        }
        console.log("File renamed successfully.");
    });

    console.log("Fanatics Exported to CSV!");
    rl.close();
}

async function updateFanSkus(wbD: Excel.Workbook, shtFan: Excel.Worksheet) {
    const shtFanskus = wbD.getWorksheet('Fanatics Item Update');
    if (!shtFanskus) {
        console.error("Sheet 'Fanatics Item Update' not found.");
        return;
    }

    const lastRowFan = shtFan.lastRow?.number || 0;

    for (let x = 2; x <= lastRowFan; x++) {
        const oldSku = shtFan.getCell(`B${x}`).value as string;

        const newSku = await vLookup(oldSku, shtFanskus);

        if(newSku){
            shtFan.getCell(`B${x}`).value = newSku;
        }
    }
}

async function vLookup(value: string, shtFanskus: Excel.Worksheet): Promise<string | undefined> {
    const range = shtFanskus.getColumn(1).values as (string | null)[];
    for (let i = 1; i < range.length; i++) {
        if (range[i] === value) {
            return shtFanskus.getCell(i + 1, 2).value as string; // Return value from column B
        }
    }
    return undefined; // Return undefined if not found
}

async function updateOnHand() {
    const workbook = new Excel.Workbook();
    const setupSheetName = 'Setup';
    const masterSheetName = 'Master';
    const pivotSheetName = 'Pivot';
    const allocationThresholdsSheetName = 'Allocation Thresholds';

    // Load the workbook and sheets
    await workbook.xlsx.readFile(inputPath);
    const setupSheet = workbook.getWorksheet(setupSheetName);
    const masterSheet = workbook.getWorksheet(masterSheetName);
    const pivotSheet = workbook.getWorksheet(pivotSheetName);
    const allocationThresholdsSheet = workbook.getWorksheet(allocationThresholdsSheetName);

    if (!setupSheet) {
        console.error(`Sheet "${setupSheetName}" not found.`);
        return;
    }

    if (!masterSheet) {
        console.error(`Sheet "${masterSheetName}" not found.`);
        return;
    }

    if (!pivotSheet) {
        console.error(`Sheet "${pivotSheetName}" not found.`);
        return;
    }

    if (!allocationThresholdsSheet) {
        console.error(`Sheet "${allocationThresholdsSheetName}" not found.`);
        return;
    }

    const inventoryFilePath = setupSheet.getCell('B2').value;
    if (!inventoryFilePath) {
        console.error("Please select/enter a valid path to the master inventory file.");
        return;
    }

    // Read the CSV file
    const csvData = fs.readFileSync(inventoryFilePath as string, 'utf8');
    const records = parse(csvData, {
        columns: true,
        skip_empty_lines: true
    }) as Array<{ SKU: string; OnHand: string }>;

    // Clear contents in Setup sheet column I
    for (let row = 1; row <= setupSheet.rowCount; row++) {
        setupSheet.getCell(`I${row}`).value = null; // Clear cell value
    }
    setupSheet.getCell('I1').value = "SKUS Not Found";

    // Clear contents in Allocation Thresholds sheet column D
    for (let row = 1; row <= allocationThresholdsSheet.rowCount; row++) {
        allocationThresholdsSheet.getCell(`D${row}`).value = null; // Clear cell value
    }
    allocationThresholdsSheet.getCell('D1').value = "SKUS Not Found on Allocation Thresholds Tab Column A";

    // Clear previous data in Master sheet
    for (let col of ['A', 'C', 'D', 'E', 'F', 'G']) {
        for (let row = 1; row <= masterSheet.rowCount; row++) {
            masterSheet.getCell(`${col}${row}`).value = null; // Clear cell value
        }
    }

    // Copy data from inventory to master
    records.forEach((record, index) => {
        const sku = record.SKU;
        const onHand = record.OnHand;
        masterSheet.getCell(`A${index + 4}`).value = sku; // Offset by 4 rows (1 header + 3 offset)
        masterSheet.getCell(`B${index + 4}`).value = onHand; // Adjust as necessary
    });

    // Process each store sheet
    const storeSheets = ['Fanatics', 'Target', 'Groupon', 'Belk'];
    for (const store of storeSheets) {
        const storeSheet = workbook.getWorksheet(store);
        if (!storeSheet) {
            console.error(`Sheet "${store}" not found.`);
            continue; // Skip to the next store if the sheet is not found
        }
        const lastRowStore = storeSheet.lastRow?.number || 0;

        for (let row = 2; row <= lastRowStore; row++) {
            const sku = storeSheet.getCell(`B${row}`).value; // Adjust column as necessary
            // Logic to check against inventory and thresholds
            const foundInInventory = records.some(record => record.SKU === sku);
            if (!foundInInventory) {
                setupSheet.getCell(`I${setupSheet.lastRow?.number || 0 + 1 || 2}`).value = sku; // Append to SKUs Not Found
            }

            const foundInThresholds = allocationThresholdsSheet.getColumn('A').values.includes(sku);
            if (!foundInThresholds) {
                allocationThresholdsSheet.getCell(`D${allocationThresholdsSheet.lastRow?.number || 0 + 1 || 2}`).value = sku; // Append to SKUs Not Found
            }
        }
    }

    // Save the workbook
    await workbook.xlsx.writeFile(inputPath);
    console.log("Update completed!");
}

async function allocate() {
    const wb = new Excel.Workbook();
    const setupSheetName = 'Setup';
    const masterSheetName = 'Master';
    const storeSheets = ['Fanatics', 'Target', 'Groupon', 'Belk'];

    await wb.xlsx.readFile(inputPath); // Change this to your workbook path
    const setupSheet = wb.getWorksheet(setupSheetName);
    if (!setupSheet) {
        console.error(`Sheet ${setupSheetName} not found.`);
        return;
    }

    const ws = wb.getWorksheet(masterSheetName);
    if(!ws){
        console.error(`Sheet ${masterSheetName} not found`);
        return;
    }

    if (!setupSheet.getCell('B2').value) {
        console.error("Please select/enter a valid path to the master inventory file.");
        return;
    }

    const lastCol = ws.getRow(4).cellCount || 0;
    const lastRow = ws.lastRow?.number || 0;

    // Trim spaces in SKU columns for each store sheet
    for (let y = 4; y <= lastCol; y++) {
        const workSheetName = ws.getCell(4, y).value;
        const storeSheet = wb.getWorksheet(`${workSheetName}`);
        if (storeSheet) {
            const lastRowStore = storeSheet.lastRow?.number || 0;
            for (let z = 2; z <= lastRowStore; z++) {
                storeSheet.getCell(`B${z}`).value = storeSheet.getCell(`B${z}`).value?.toString().trim();
            }
        }
    }

    for (let x = 5; x <= lastRow; x++) {
        const strSku = ws.getCell(`A${x}`).value?.toString().trim();
        let intCountStores = 0;

        for (let y = 4; y <= lastCol; y++) {
            const workSheetName = ws.getCell(4, y).value;
            const storeSheet = wb.getWorksheet(`${workSheetName}`);
            if (storeSheet) {
                const foundRow = storeSheet.getColumn('B').values.indexOf(strSku) + 1; // +1 for 1-based index
                if (foundRow > 0) {
                    intCountStores++;
                }
            }
        }

        // Handle allocation logic here
        // Example: If foundRow > 0, update the quantity in the respective store sheet
        for (const store of storeSheets) {
            const storeSheet = wb.getWorksheet(store);
            if (storeSheet) {
                const foundRow = storeSheet.getColumn('B').values.indexOf(strSku) + 1; // Adjust column as necessary
                if (foundRow > 0) {
                    storeSheet.getCell(`D${foundRow}`).value = ws.getCell(`D${x}`).value; // Adjust columns as necessary
                }
            }
        }
    }

    // Export logic for each store sheet
    for (const store of storeSheets) {
        const storeSheet = wb.getWorksheet(store);
        if (storeSheet) {
            const date = new Date();
            const formattedDate = store === 'Fanatics' ? 
                `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}_` +
                `${String(date.getHours()).padStart(2, '0')}${String(date.getMinutes()).padStart(2, '0')}${String(date.getSeconds()).padStart(2, '0')}`
                :
                `${String(date.getMonth() + 1).padStart(2, '0')}_${String(date.getDate()).padStart(2, '0')}_${date.getFullYear()}`;
            const prefix = store === 'Fanatics' ? 'IB_LDM_' : store;
            const filePath = `data/${prefix}_${formattedDate}.xlsx`;
            await wb.xlsx.writeFile(filePath);
            // const filePath = `data/${store}_${new Date().toISOString().split('T')[0]}.xlsx`;
            // await wb.xlsx.writeFile(filePath);
        }
    }

    console.log("Allocated!");
}

async function selectFile(title: string): Promise<string> {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    return new Promise((resolve) => {
        rl.question(`${title}: `, (filePath) => {
            rl.close();
            resolve(filePath);
        });
    });
}