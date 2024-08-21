const xlsx = require('xlsx');
const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
puppeteer.use(StealthPlugin());

(async () => {
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();
    await page.goto('https://tge.pl/energia-elektryczna-rdn?dateShow=18-07-2024&dateAction=prev', { waitUntil: 'networkidle2' });

    // waiting for the table to load
    // await page.waitForSelector('.table');
    await page.waitForSelector('.footable.table.table-hover.table-padding');

    // Pobierz dane z tabeli
    const tableData = await page.evaluate(() => {
        // Użyj rzeczywistego selektora tabeli
        const table = document.querySelector('.footable.table.table-hover.table-padding'); 
        const rows = Array.from(table.querySelectorAll('tr'));
        return rows.map(row => {
            const cells = Array.from(row.querySelectorAll('td, th'));
            return cells.map(cell => cell.innerText.trim());
        });
    });

    // Zakończenie działania przeglądarki
    await browser.close();

    // Tworzenie nowego skoroszytu Excel
    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.aoa_to_sheet(tableData);

    // Dodaj arkusz do skoroszytu
    xlsx.utils.book_append_sheet(wb, ws, 'Tabela');

    // Zapisz plik Excel
    xlsx.writeFile(wb, 'tabela.xlsx');

    console.log('Plik Excel został zapisany jako tabela.xlsx');
})();
