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
    // Poniżej, w cudzysłowie należy wpisać daty dla których chcemy pobrać dane. Ważne aby były w formacie DD-MM-YYYY jak poniżej, np.:
    // const startDate = '27-06-2024';
    // const endDate = '21-08-2024';
    const startDate = '29-07-2024';
    const endDate = '02-08-2024';

    const dates = await generateDates(startDate, endDate);
    // Tworzenie nowego excela
    const wb = xlsx.utils.book_new();

    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();

    // do zbierania danych
    let allDataTable = [];
    let ws = null;

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
            
            
            // Pominięcie dwóch pierwszych wierszy (nagłówki)
            const dataRows = rows.slice(2);
            
            return rows.map(row => {
                const cells = Array.from(row.querySelectorAll('td, th'));

                return cells.map(cell => cell.innerText.trim());
            });
        });
        
        // dodawanie daty do każdego wiersza
        const tableDataWithDate = tableData.map(row => [date, ...row]);

        allDataTable = allDataTable.concat(tableDataWithDate);
        ws = xlsx.utils.aoa_to_sheet(allDataTable);
    }
    // Dodawanie arkuszu do excela
    xlsx.utils.book_append_sheet(wb, ws, 'Całość');

    // wyłączenie przeglądarki
    await browser.close();

    // Zapisywanie excela
    xlsx.writeFile(wb, 'tabela.xlsx');
    console.log('Plik Excel został zapisany jako tabela.xlsx');
})();
