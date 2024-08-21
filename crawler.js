const xlsx = require('xlsx');
const puppeteer = require('puppeteer-extra');
const moment = require('moment');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
puppeteer.use(StealthPlugin());

async function generateDates(startDate, endDate) {
    let dates = [];
    let currentDate = moment(startDate, 'DD-MM-YYYY');
    let end = moment(endDate, 'DD-MM-YYYY');

    while (currentDate <= end) {
        dates.push(currentDate.format('DD-MM-YYYY'));
        currentDate = currentDate.add(1, 'days');
    }

    return dates;
}

(async () => {
    // Poniżej, w cudzysłowie należy podać daty dla których chcemy pobrać dane. Ważne aby były w formacie DD-MM-YYYY jak poniżej
    const startDate = '28-06-2024';
    const endDate = '21-08-2024';

    const dates = await generateDates(startDate, endDate);
    // Tworzenie nowego excela
    const wb = xlsx.utils.book_new();

    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();

    // const workbook = new ExcelJS.Workbook();


    // Tworzenie arkusza dla każdej daty
    for (const date of dates) {
        // otwieranie strony
        await page.goto('https://tge.pl/energia-elektryczna-rdn?dateShow=' + date + '&dateAction=prev', { waitUntil: 'networkidle2' });

        // czekanie na załadowanie tabeli
        await page.waitForSelector('.footable.table.table-hover.table-padding');

        // Pobieranie danych z tabeli
        const tableData = await page.evaluate(() => {
            const table = document.querySelector('.footable.table.table-hover.table-padding'); 
            const rows = Array.from(table.querySelectorAll('tr'));
            return rows.map(row => {
                const cells = Array.from(row.querySelectorAll('td, th'));
                return cells.map(cell => cell.innerText.trim());
            });
        });

        const ws = xlsx.utils.aoa_to_sheet(tableData);
        
        // Dodawanie arkuszu do excela
        xlsx.utils.book_append_sheet(wb, ws, date);
    }

    // wyłączenie przeglądarki
    await browser.close();

    // Zapisywanie excela
    xlsx.writeFile(wb, 'tabela.xlsx');
    console.log('Plik Excel został zapisany jako tabela.xlsx');
})();
