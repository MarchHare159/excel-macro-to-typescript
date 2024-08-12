import * as readline from 'readline';

export const paths = {
    inputPath: 'data/Allocation Master New.xlsm',
    outputPath: 'data/1.xlsm',
}

export const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

export async function confirmOverwrite(message: string): Promise<boolean> {
    return new Promise((resolve) => {
        rl.question(message, (answer) => {
            resolve(answer.toLowerCase() === 'yes');
        });
    });
}