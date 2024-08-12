import * as Excel from 'exceljs';
import { rl, confirmOverwrite, paths } from '../utils/utils';
const inputPath = paths.inputPath;

export async function getProposedThresholds(filePath: string, addSkus: boolean) {
    const workbook = new Excel.Workbook();
    const tempWorkbook = await workbook.xlsx.readFile(filePath);
    const dataSheet = tempWorkbook.worksheets[0];

    const rowData = dataSheet.getSheetValues();
    const data: any[][] = rowData
        .filter(row => row !== undefined)
        .map(row => row as any[]);

    const tempAt = await getUsedRange("Allocation Thresholds");
    const missingSkus = missingItemsWithoutDups(data, 1, tempAt, 1);

    if (!Array.isArray(missingSkus)) {
        console.error('No missing SKUs found.');
        return;
    }

    // Write missingSkus to Allocation Thresholds
    if (addSkus) {
        let index: number | null = null;

        for (let i = 0; i < tempAt.length; i++) {
            if (tempAt[i][0] === '') {
                index = i;
                break;
            }
        }

        if (index !== null) {
            for (let i = 0; i < missingSkus.length; i++) {
                await writeToSheet("Allocation Thresholds", index + 1, 1, missingSkus[i]);
                index++;
            }
        }
    }

    // Initialize at
    let at = await getUsedRange("Allocation Thresholds");

    // Remove headers from data
    for (let j = 1; j <= data[0].length; j++) {
        data[1][j] = null; // Clear header
    }

    // Reformat dates in data
    for (let i = 2; i < data.length; i++) {
        if (i == 2) {
            console.log(data[2][4]);
        }
        data[i][4] = dateFormat(data[i][4]);
    }

    // Set column C in at to 0
    for (let i = 1; i < at.length; i++) {
        at[i][2] = 0; // Assuming column C is index 2
    }

    // Fill column C in at
    let tempSku = data[2][1]; // Assuming column A is index 1
    let tempDate = data[2][4]; // Assuming column D is index 4
    let mySum = 0;
    let mostSales = [0, 0, 0];

    // Loop through data and update column 3 in at
    for (let i = 2; i < data.length; i++) {
        if (data[i][1] === null) { // Check for null at the end of file
            const updatedAt = vLookUpReplace(tempSku, at, 3, totalSum(mostSales));
            if (Array.isArray(updatedAt)) {
                at = updatedAt;
            } else {
                console.error('vLookUpReplace did not return a valid array.');
            }
            break;
        }

        if (data[i][1] === tempSku) {
            if (data[i][4] === tempDate) {
                mySum += data[i][3]; // Assuming column C is index 3
            } else {
                mostSales = largestInts(mySum, mostSales);
                mySum = data[i][3];
                tempDate = data[i][4];
            }
        } else {
            const updatedAt = vLookUpReplace(tempSku, at, 3, totalSum(mostSales));
            if (Array.isArray(updatedAt)) {
                at = updatedAt;
            } else {
                console.error('vLookUpReplace did not return a valid array.');
            }
            tempSku = data[i][1];
            tempDate = data[i][4];
            mySum = data[i][3];
            mostSales = [mySum, 0, 0];
        }
    }

    // Write at column 3 to sheet
    for (let i = 2; i < at.length; i++) {
        if (at[i][1] !== null) {
            await writeToSheet("Allocation Thresholds", i + 1, 3, at[i][3]); // Column C is index 3
        }
    }

    // Sort Allocations Thresholds tab
    await sortSheet("Allocation Thresholds");

    console.log("Done!");
}

export async function updateProposedThresholds(filePath: string, isGreen: boolean = false) {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sht = workbook.getWorksheet('Allocation Thresholds');

    if (!sht) {
        console.error("Sheet 'Allocation Thresholds' not found.");
        rl.close();
        return;
    }

    const confirmed = await confirmOverwrite("Are you sure you want to overwrite thresholds with computer calculated proposed thresholds? (yes/no): ");
    if (!confirmed) {
        console.error('Operation cancelled.');
        rl.close();
        return;
    }

    let lastRow: number | null = null;

    // Find the last row in column A
    for (let row = sht.rowCount; row >= 1; row--) {
        const cellValue = sht.getCell(`A${row}`).value;
        if (cellValue !== null && cellValue !== '') {
            lastRow = row;
            break;
        }
    }

    if (lastRow === null) {
        console.error("Last row on Allocation Thresholds sheet not found.");
        rl.close();
        return;
    }

    // Update values in column B based on the condition in column C
    for (let x = 2; x <= lastRow; x++) {
        const valueB = sht.getCell(`B${x}`).value;
        const valueC = sht.getCell(`C${x}`).value;

        if (isGreen) {
            // Condition for Green
            if (typeof valueC === 'number' && typeof valueB === 'number' && valueC > valueB) {
                sht.getCell(`B${x}`).value = valueC; // Copy value from C to B
            }
        } else {
            // Condition for standard update
            if (typeof valueC === 'number' && valueC >= 10) {
                sht.getCell(`B${x}`).value = valueC; // Copy value from C to B
            }
        }
    }

    // Save the updated workbook
    await workbook.xlsx.writeFile(inputPath);
    console.log("Done!");
    rl.close();
}

export async function updateProposedThreshodsGreen(filePath: string) {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sht = workbook.getWorksheet('Allocation Thresholds');

    if (!sht) {
        console.error("Sheet 'Allocation Thresholds' not found.");
        rl.close();
        return;
    }

    const confirmed = await confirmOverwrite("Are you sure you want to overwrite thresholds with computer calculated proposed thresholds? (yes/no): ");
    if (!confirmed) {
        console.error('Operation cancelled.');
        rl.close();
        return;
    }

    const lastRow = findLastRow(sht, 'A');

    console.warn(lastRow);


    if (lastRow == -1) {
        console.error("Last Row on Allocations Threshold sheet not found.");
        return;
    }

    for (let x = 2; x <= lastRow; x++) {
        const valueC = sht.getCell(`C${x}`).value as number;
        const valueB = sht.getCell(`B${x}`).value as number;

        if (valueC > valueB) {
            sht.getCell(`B${x}`).value = valueC;
        }
    }

    await workbook.xlsx.writeFile(inputPath);
    console.log("Done!");
}

export async function updateProposedThresholds10(filePath: string) {
    const confirmed = await confirmOverwrite("Are you sure you want to overwrite thresholds with 10s? (yes/no): ");
    if (!confirmed) {
        console.log('Operation cancelled.');
        rl.close();
        return;
    }

    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sht = workbook.getWorksheet('Allocation Thresholds');

    if (!sht) {
        console.error("Sheet 'Allocation Thresholds' not found.");
        rl.close();
        return;
    }

    let lastRow: number | null = null;

    // Find the last row in column A
    for (let row = sht.rowCount; row >= 1; row--) {
        const cellValue = sht.getCell(`A${row}`).value;
        if (cellValue !== null && cellValue !== '') {
            lastRow = row;
            break;
        }
    }

    if (lastRow === null) {
        console.error("Last row on Allocation Thresholds sheet not found.");
        rl.close();
        return;
    }

    // Set all values in column B from row 2 to lastRow to 10
    for (let x = 2; x <= lastRow; x++) {
        sht.getCell(`B${x}`).value = 10;
    }

    // Save the updated workbook
    await workbook.xlsx.writeFile(inputPath);
    console.log("Done!");
    rl.close();
}

async function getUsedRange(sheetName: string): Promise<any[]> {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(inputPath); // Replace with your file path
    const sheet = workbook.getWorksheet(sheetName);

    if (!sheet) {
        throw new Error(`Sheet '${sheetName}' dose not exist.`);
    }

    const data: any[] = [];
    const lastRow = sheet.lastRow?.number || 0;

    for (let i = 1; i <= lastRow; i++) { // Start from row 1 to include headers
        const row = sheet.getRow(i);
        data.push(row.values); // Push the entire row values
    }

    return data;
}

function findLastRow(sheet: Excel.Worksheet, column: string): number {
    const columnCells = sheet.getColumn(column).values;
    for (let i = columnCells.length - 1; i >= 1; i--) { // Start from 1 to skip the header
        if (columnCells[i] !== null) {
            return i; // Return the last row number
        }
    }
    return -1; // Not found
}

function missingItemsWithoutDups(
    inArr: any[][],
    inArrCol: number,
    searchArr: any[][],
    searchArrCol: number
): any[] | boolean {
    // Initialize answer array
    const answer: any[] = [];
    let found: boolean;

    // Start searching for missing elements
    for (let i = 1; i < inArr.length; i++) { // Skip header
        found = false;
        for (let j = 1; j < searchArr.length; j++) { // Skip header
            if (searchArr[j][searchArrCol] === inArr[i][inArrCol]) {
                found = true;
                break;
            }
        }
        // Add to answer if not found in searchArr
        if (!found) {
            answer.push(inArr[i][inArrCol]);
        }
    }

    // Check if there are no missing items
    if (answer.length === 0) {
        return false; // No missing SKUs
    } else {
        return removeColumnDuplicates(answer); // Return unique missing items
    }
}

function removeColumnDuplicates(myArr: any[]): any[] {
    // Remove duplicates
    const preAnswer = deleteColumnDuplicates(myArr);
    // Remove null or empty spaces
    return deleteNullSpaces(preAnswer);
}

function deleteColumnDuplicates(myArr: any[]): (string | null)[] {
    const answer: (string | null)[] = [...myArr];

    for (let i = 0; i < answer.length; i++) {
        const myVal = answer[i];
        if (myVal !== '') {
            for (let j = i + 1; j < answer.length; j++) {
                if (myVal === answer[j]) {
                    answer[j] = null;
                }
            }
        }
    }

    return answer;
}

function deleteNullSpaces(myArr: (string | null)[]): string[] {
    // Count non-null and non-empty strings
    const nonNulls: string[] = myArr.filter((item): item is string => item !== null && item !== '');
    return nonNulls;
}

function dateFormat(myDate: string): string {
    // This function changes the date to mm/dd/yyyy 0:00 format.
    // e.g. "5/17/2022 10:47" --> "5/17/2022 0:00"

    // Check if the input is a valid date string
    if (typeof myDate !== 'string') {
        console.error("Invalid input: myDate must be a string.");
        return "";
    }
    const dateParts = myDate.split(" ");
    if (dateParts.length === 0) return myDate; // Return original if empty

    const dateString = dateParts[0]; // Get the date part

    // Split the date string by slashes
    const parts = dateString.split("/");
    if (parts.length !== 3) return myDate; // Return original if not in mm/dd/yyyy format

    const month = parts[0];
    const day = parts[1];
    const year = parts[2];

    // Validate the date parts
    if (!month || !day || !year) return myDate; // Return original if any part is missing

    // Format the new date string
    return `${month}/${day}/${year} 0:00`;
}

function vLookUpReplace(searchVal: any, myArr: any[][], colNum: number, writeVal: any): any[] | boolean {
    // This function searches the first column of myArr for searchVal.
    // When it finds it, writeVal is written to colNum column in the same row,
    // and the array is returned. It only does this for the first instance of the searchVal.

    const myStart = 0; // Starting index for JavaScript arrays
    const myEnd = myArr.length - 1; // End index for JavaScript arrays
    let found = false;
    const answer = myArr.map(row => [...row]); // Create a copy of the array

    for (let i = myStart; i <= myEnd; i++) {
        if (answer[i][0] === searchVal) {
            found = true;
            answer[i][colNum] = writeVal; // Replace the value in the specified column
            break; // Exit after the first match
        }
    }

    return found ? answer : false; // Return the modified array or false if not found
}

function totalSum(theInts: number[]): number {
    // This function returns the total sum of the first 3 ints in an int array

    let answer = 0;

    // Calculate the upper bound for summation
    const upperBound = Math.min(2, theInts.length - 1); // Ensure we don't exceed the array length

    for (let i = 0; i <= upperBound; i++) {
        answer += theInts[i];
    }

    return answer;
}

function largestInts(myInt: number, theInts: number[]): number[] {
    // This function takes in an integer and an array containing three integers.
    // It returns an array containing the largest of the four input integers.

    let minIndex: number;

    // Find the index of the minimum value in theInts
    if (theInts[0] < theInts[1]) {
        if (theInts[0] < theInts[2]) {
            minIndex = 0;
        } else {
            minIndex = 2;
        }
    } else {
        if (theInts[1] < theInts[2]) {
            minIndex = 1;
        } else {
            minIndex = 2;
        }
    }

    // Replace min value in theInts with myInt if it's larger
    if (myInt > theInts[minIndex]) {
        theInts[minIndex] = myInt;
    }

    // Return solution
    return theInts;
}

async function writeToSheet(sheetName: string, row: number, col: number, value: any) {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(inputPath);
    const worksheet = workbook.getWorksheet(sheetName);
    if (worksheet) {
        worksheet.getCell(row, col).value = value; // Set the value in the specified cell
        await workbook.xlsx.writeFile(inputPath); // Save the workbook
    } else {
        console.error(`Sheet ${sheetName} not found.`);

    }
}

async function sortSheet(sheetName: string) {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(inputPath);
    const worksheet = workbook.getWorksheet(sheetName);
    if (worksheet) {
        // Get the last row with data
        const lastRow = worksheet.lastRow?.number || 0;

        // Create an array to hold the rows for sorting
        const rows: Excel.Row[] = [];

        // Collect rows from the worksheet
        for (let i = 2; i <= lastRow; i++) { // Start from row 2 to skip headers
            const row = worksheet.getRow(i);
            rows.push(row);
        }

        // Sort rows based on column C (index 2, since it's 0-based)
        rows.sort((a, b) => {
            const valueA = a.getCell(3).value; // Column C
            const valueB = b.getCell(3).value;

            // Handle undefined values
            if (valueA === undefined) return 1; // Treat undefined as greater
            if (valueB === undefined) return -1; // Treat undefined as lesser

            return (valueB as number) - (valueA as number); // Descending sort
        });

        // Write sorted rows back to the worksheet
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (row && row.values) {
                const values = row.values as Excel.CellValue[];
                values.forEach((value, index) => {
                    worksheet.getCell(i + 2, index + 1).value = value; // Write back to the sheet
                });
            }
        }

        // Save the workbook
        await workbook.xlsx.writeFile(inputPath); // Save the workbook
        console.log("Sorting done!");
    } else {
        console.error(`Sheet ${sheetName} not found.`);
    }
}