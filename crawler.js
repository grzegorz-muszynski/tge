const xlsx = require('xlsx');
const xlsxStyle = require('xlsx-style');

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

function mergeTablesSideBySide(table1, table2) {
    const maxRows = Math.max(table1.length, table2.length);
    const mergedTable = [];

    for (let i = 0; i < maxRows; i++) {
        const row1 = table1[i] || [];
        const row2 = table2[i] || [];
        const mergedRow = row1.concat(row2);
        mergedTable.push(mergedRow);
    }

    return mergedTable;
}

(async () => {
    // Poniżej, w cudzysłowie należy wpisać daty dla których chcemy pobrać dane. Ważne aby były w formacie DD-MM-YYYY jak poniżej, np.:
    // const startDate = '27-06-2024';
    // const endDate = '21-08-2024';
    const startDate = '28-07-2024';
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
    let dateNumber = 0;

    // Pokazywanie komunikatów z przeglądarki w konsoli
    page.on('console', msg => console.log('PAGE LOG:', msg.text()));

    
    // Tworzenie arkusza dla każdej daty
    for (const date of dates) {
        // otwieranie strony
        await page.goto('https://tge.pl/energia-elektryczna-rdn?dateShow=' + date + '&dateAction=prev', { waitUntil: 'networkidle2' });
        // czekanie na załadowanie tabeli
        await page.waitForSelector('.footable.table.table-hover.table-padding');
        
        
        // Pobieranie danych z tabeli
        const tableData = await page.evaluate((dateNumber) => {
            const table = document.querySelector('.footable.table.table-hover.table-padding'); 
            const allRows = Array.from(table.querySelectorAll('tr'));
            
            return allRows.map(row => {
                if (dateNumber < 1) {
                    cells = Array.from(row.querySelectorAll('td, th')).slice(0, 2);
                } else {
                    cells = Array.from(row.querySelectorAll('td, th')).slice(1, 2);
                }
                
                return cells.map(cell => cell.innerText.trim());
            });
        }, dateNumber);
        
        
        // Pominięcie dwóch pierwszych wierszy (z nagłówkami) przy kolejnych iteracjach
        // const dataRows = isFirstIteration ? tableData.slice(0, -3) : tableData.slice(2, -3);
        const dataRows = tableData.slice(2, -3);
        
        // dodawanie daty do każdej kolumny z wyjątkiem pierwszej
        if (isFirstIteration) {
            dataRows.unshift(['', date]);
        } else {   
            dataRows.unshift([date]);
        }

        // dodawanie nagłówków w dwóch pierwszych wierszach
        if (dateNumber === 1) {
            dataRows.splice(0, 0, ['FIXING I']);
            dataRows.splice(0, 0, ['Kurs (PLN/MWh)']);
        } else {
            dataRows.splice(0, 0, ['']);
            dataRows.splice(0, 0, ['']);
        }
        
        const tableDataWithDate = dataRows; 
        
        // allDataTable = allDataTable.concat(tableDataWithDate);
        allDataTable = mergeTablesSideBySide(allDataTable, tableDataWithDate);
        isFirstIteration = false;
        dateNumber = dateNumber + 1;
    }
    
    ws = xlsx.utils.aoa_to_sheet(allDataTable);

    // Step 4: Apply bold formatting to specific cells
    const boldStyle = { font: { bold: true } };
    ws['B2'].s = boldStyle;
    ws['B1'].s = boldStyle;

    for (let i = 4; i < 28; i++) {
        ws['A' + i].s = boldStyle;
    }

    
    // Wyliczanie szerokości kolumn
    // const colWidths = calculateColumnWidths(allDataTable);

    // Ręczne ustawienie szerokości kolumn
    const colWidths = [
        { wch: 12},
        { wch: 16},
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
    // xlsx.writeFile(wb, 'tabela.xlsx');
    xlsxStyle.writeFile(wb, 'tabela.xlsx');
    console.log('Plik Excel został zapisany jako tabela.xlsx');
})();
