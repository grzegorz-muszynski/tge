const readline = require('readline');
const xlsx = require('xlsx');
const xlsxStyle = require('xlsx-style');
const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const chromium = require('puppeteer');
const { generateDates, parseDate, getDate59DaysBefore, getPrevDay } = require('./helpers');

puppeteer.use(StealthPlugin());

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

async function makeExcel(startDate, endDate) {
    console.log('Rozpoczynam pobieranie danych z wybranych dat:');
    const dates = await generateDates(startDate, endDate);
    
    const wb = xlsx.utils.book_new();
    const browser = await puppeteer.launch({ headless: true, executablePath: chromium.executablePath() });
    const page = await browser.newPage();
    
    let allDataTable = [['Godzina', ...dates]];

    for (let i = 0; i < 24; i++) {
        allDataTable.push([i + ':00']);
    }

    for (const date of dates) {
        console.log(date);
        const prevDate = getPrevDay(date);
        await page.goto('https://tge.pl/energia-elektryczna-rdn?dateShow=' + prevDate + '&dateAction=prev', { waitUntil: 'networkidle2' });
        await page.waitForSelector('.footable.table.table-hover.table-padding');
        
        const tableData = await page.evaluate(() => {
            const table = document.querySelector('.footable.table.table-hover.table-padding'); 
            const rows = Array.from(table.querySelectorAll('tr'));
            return rows.slice(2, 26).map(row => row.querySelectorAll('td')[1].innerText.trim());
        });

        for (let i = 0; i < 24; i++) {
            allDataTable[i + 1].push(tableData[i]);
        }
    }
    
    const ws = xlsx.utils.aoa_to_sheet(allDataTable);
    xlsx.utils.book_append_sheet(wb, ws, 'Dane');
    await browser.close();
    xlsxStyle.writeFile(wb, 'tabela.xlsx');
    console.log('Plik Excel został zapisany jako tabela.xlsx');
}

const dateRegex = /^\d{2}-\d{2}-\d{4}$/;
const earliestDate = getDate59DaysBefore();

function askDate(question, callback) {
    rl.question(question, (date) => {
        if (dateRegex.test(date)) {
            callback(date);
        } else {
            console.log('Niepoprawny format daty. Spróbuj ponownie.');
            askDate(question, callback);
        }
    });
}

askDate(`Wprowadź datę początkową (DD-MM-RRRR), nie wcześniej niż ${earliestDate}: `, (startDate) => {
    const parsedStartDate = parseDate(startDate);
    const parsedEarliestDate = parseDate(earliestDate);

    if (parsedStartDate < parsedEarliestDate) {
        console.log(`Podana data jest wcześniejsza niż ${earliestDate}.`);
        rl.close();
        return;
    }

    askDate('Wprowadź datę końcową (DD-MM-YYYY): ', (endDate) => {
        const parsedEndDate = parseDate(endDate);

        if (parsedStartDate > parsedEndDate) {
            console.log('Podana data końcowa jest wcześniejsza niż data początkowa.');
            rl.close();
            return;
        }
        
        makeExcel(startDate, endDate);
        rl.close();
    });
});
