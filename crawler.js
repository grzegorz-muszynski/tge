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

// Funkcja wyliczająca szerokość kolumn w excelu
function calculateColumnWidths(data) {
    const colWidths = data[0].map((_, colIndex) => {
        return Math.max(...data.map(row => (row[colIndex] ? row[colIndex].toString().length : 0)));
    });
    return colWidths.map(width => ({ wch: width }));
}

(async () => {
    // Poniżej, w cudzysłowie należy wpisać daty dla których chcemy pobrać dane. Ważne aby były w formacie DD-MM-YYYY jak poniżej, np.:
    // const startDate = '27-06-2024';
    // const endDate = '21-08-2024';
    const startDate = '01-08-2024';
    const endDate = '02-08-2024';

    const dates = await generateDates(startDate, endDate);
    // Tworzenie nowego excela
    const wb = xlsx.utils.book_new();

    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();

    // do zbierania danych
    let allDataTable = [];
    let ws = null;
    let isFirstIteration = true;

    // Tworzenie arkusza dla każdej daty
    for (const date of dates) {
        // otwieranie strony
        await page.goto('https://tge.pl/energia-elektryczna-rdn?dateShow=' + date + '&dateAction=prev', { waitUntil: 'networkidle2' });

        // czekanie na załadowanie tabeli
        await page.waitForSelector('.footable.table.table-hover.table-padding');

        // Pobieranie danych z tabeli
        const tableData = await page.evaluate(() => {
            const table = document.querySelector('.footable.table.table-hover.table-padding'); 
            const allRows = Array.from(table.querySelectorAll('tr'));

            return allRows.map(row => {
                const cells = Array.from(row.querySelectorAll('td, th'));

                return cells.map(cell => cell.innerText.trim());
            });
        });

        // Pominięcie dwóch pierwszych wierszy (z nagłówkami) przy kolejnych iteracjach
        const dataRows = isFirstIteration ? tableData.slice(0, -3) : tableData.slice(2, -3);
        isFirstIteration = false;
        
        // dodawanie daty do każdego wiersza
        const tableDataWithDate = dataRows.map(row => [date, ...row]);

        allDataTable = allDataTable.concat(tableDataWithDate);
    }

    // Dodawanie pustych komórek w pierwszym wierszu
    allDataTable[0].splice(3, 0, ''); 
    allDataTable[0].splice(5, 0, ''); 
    allDataTable[0].splice(8, 0, '');

    // Usunięcie zbędnych dat
    allDataTable[0][0] = '';
    allDataTable[1][0] = '';
    
    ws = xlsx.utils.aoa_to_sheet(allDataTable);
    
    // Wyliczanie szerokości kolumn
    // const colWidths = calculateColumnWidths(allDataTable);

    // Ręczne ustawienie szerokości kolumn
    const colWidths = [
        { wch: 12},
        { wch: 7},
        { wch: 16},
        { wch: 16},
        { wch: 16},
        { wch: 16},
        { wch: 16},
        { wch: 16}
    ]

    ws['!cols'] = colWidths;

    // Dodawanie arkuszu do excela
    xlsx.utils.book_append_sheet(wb, ws, 'Całość');

    // wyłączenie przeglądarki
    await browser.close();

    // Zapisywanie excela
    xlsx.writeFile(wb, 'tabela.xlsx');
    console.log('Plik Excel został zapisany jako tabela.xlsx');
})();
