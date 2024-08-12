import * as Excel from 'exceljs';
import * as fs from 'fs';
import * as readline from 'readline';
import { paths } from '../utils/utils';
const filePath = paths.inputPath;

export async function browseXlsFile() {
    const filename = await selectFile();
    if (filename) {
        // Assuming you want to set the filename to a specific cell in a workbook
        const wb = new Excel.Workbook();
        await wb.xlsx.readFile(filePath);
        const shtSetup = wb.getWorksheet('Setup');
        if(!shtSetup){
            console.error("Sheet 'Setup' not found.}");
            return;
        }
        shtSetup.getCell('B2').value = filename;        
        
        // Save the workbook if needed
        await wb.xlsx.writeFile(filePath); // Change this to your desired path
        console.log(`File selected: ${filename}`);
    }
    else{
        console.error("File not found");
    }
}

export async function browseCSVFolder() {
    const folderPath = await selectFolder();
    if (folderPath) {
        // Assuming you want to set the folder path to a specific cell in a workbook
        const wb = new Excel.Workbook();
        await wb.xlsx.readFile(filePath);
        const shtSetup = wb.getWorksheet('Setup');
        if(!shtSetup){
            console.error("Sheet 'Setup' not found.");
            return;
        }
        shtSetup.getCell('B8').value = folderPath;

        // Save the workbook if needed
        await wb.xlsx.writeFile(filePath); // Change this to your desired path
        console.log(`Folder selected: ${folderPath}`);
    }
    else{
        console.error("Folder not found.");
    }
}

export async function browseDesktopFolder() {
    const folderPath = await selectFolder();
    if (folderPath) {
        // Create or modify an Excel workbook
        const wb = new Excel.Workbook();
        await wb.xlsx.readFile(filePath);
        await wb.xlsx.readFile(filePath);
        const shtSetup = wb.getWorksheet('Setup');
        if(!shtSetup){
            console.error("Sheet 'Setup' not found");
            return;
        }
        shtSetup.getCell('B14').value = folderPath;

        // Save the workbook if needed
        await wb.xlsx.writeFile(filePath); // Change this to your desired path
        console.log(`Folder selected: ${folderPath}`);
    } else {
        console.error("Folder not found");
    }
}

async function selectFile(): Promise<string | undefined> {
    const filePath = await promptUserForFile();
    if (filePath && fs.existsSync(filePath)) {
        return filePath;
    }
    return undefined;
}

async function selectFolder(): Promise<string | undefined> {
    const folderPath = await promptUserForFolder();
    if (folderPath && fs.existsSync(folderPath)) {
        return folderPath;
    }
    return undefined;
}

async function promptUserForFile(): Promise<string | undefined> {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    return new Promise((resolve) => {
        rl.question('Please enter the path to the Excel or CSV file: ', (answer) => {
            rl.close();
            resolve(answer);
        });
    });
}

async function promptUserForFolder(): Promise<string | undefined> {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    return new Promise((resolve) => {
        rl.question('Please enter the path to the folder: ', (answer) => {
            rl.close();
            resolve(answer);
        });
    });
}
