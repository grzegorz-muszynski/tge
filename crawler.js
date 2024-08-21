const puppeteer = require('puppeteer');
const xlsx = require('xlsx');

(async () => {
    // Uruchomienie przeglądarki Puppeteer
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();

    // Przejdź na stronę docelową
    await page.goto('https://tge.pl/energia-elektryczna-rdn');

    // Zaczekaj, aż tabela załaduje się na stronie
    await page.waitForSelector('.table'); // Użyj rzeczywistego selektora tabeli

    // Pobierz dane z tabeli
    const tableData = await page.evaluate(() => {
        const table = document.querySelector('.table'); // Użyj rzeczywistego selektora tabeli
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
