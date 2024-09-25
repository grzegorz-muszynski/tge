const readline = require('readline');

const xlsx = require('xlsx');
const xlsxStyle = require('xlsx-style');

const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');

const { generateDates, calculateColumnWidths, mergeTablesSideBySide, parseDate, getDate59DaysBefore, getPrevDay, getNextDay } = require('./helpers');

puppeteer.use(StealthPlugin());

// Prompts
const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

// Funkcja główna
async function makeExcel (startDate, endDate) {
    console.log('Rozpoczynam pobieranie danych z wybranych dat:');

    // Poniżej, w cudzysłowie należy wpisać daty dla których chcemy pobrać dane. Ważne aby były w formacie DD-MM-YYYY jak poniżej, np.:
    // const startDate = '27-06-2024';
    // const endDate = '21-08-2024';

    const dates = await generateDates(startDate, endDate);
    
    // console.log('Pobieranie danych dla dat: ', dates);
    
    // Tworzenie nowego excela
    const wb = xlsx.utils.book_new();

    // Otwieranie przeglądarki
    const browser = await puppeteer.launch({headless: true});
    const page = await browser.newPage();
    
    // do zbierania danych
    let allDataTable = [];
    let ws = null;
    let isFirstIteration = true;
    let dateNumber = 0;
    
    // Pokazywanie komunikatów z przeglądarki w konsoli
    // page.on('console', msg => console.log('PAGE LOG:', msg.text()));
    
    // Tworzenie arkusza dla każdej daty
    for (const date of dates) {
        console.log(date);
        // for (let i = 0; i < dates.length - 1; i++) {
    //     const date = dates[i];
        const prevDate = getPrevDay(date);
        
        // otwieranie strony
        await page.goto('https://tge.pl/energia-elektryczna-rdn?dateShow=' + prevDate + '&dateAction=prev', { waitUntil: 'networkidle2' });
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
        
        // przerabianie daty o jeden dzień na przód i dodawanie w tej formie do każdej kolumny z wyjątkiem pierwszej (tej z godzinami)
        // const nextDayDate = getNextDay(date);
        if (isFirstIteration) {
            dataRows.unshift(['', date]);
        } else {   
            dataRows.unshift([date]);
        }

        // dodawanie nagłówków w dwóch pierwszych wierszach
        if (dateNumber === 0) {
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

    // Stylowanie
    const boldStyle = { font: { bold: true } };
    
    // Pogrubianie tekstu w pierwszej kolumnie
    for (let i = 1; i < 28; i++) {
        ws['A' + i].s = boldStyle;
    }
    
    // Pogrubianie dat
    function numberToLetters(num) {
        let letters = '';
        while (num >= 0) {
            letters = String.fromCharCode((num % 26) + 65) + letters;
            num = Math.floor(num / 26) - 1;
        }
        return letters;
    }
    
    const iterations = dates.length; // Set the number of iterations you need

    for (let i = 1; i < iterations + 1; i++) {
        ws[numberToLetters(i) + '3'].s = boldStyle;
    }

    
    // Wyliczanie szerokości kolumn
    const colWidths = calculateColumnWidths(allDataTable);

    // Ręczne ustawienie szerokości kolumn
    // const colWidths = [
    //     { wch: 12},
    //     { wch: 16},
    //     { wch: 16},
    //     { wch: 16},
    //     { wch: 16},
    //     { wch: 16},
    //     { wch: 16},
    //     { wch: 16}
    // ]

    ws['!cols'] = colWidths;

    // Dorównanie do prawej strony komórek
    const rightAlignStyle = { alignment: { horizontal: "right" } };
    Object.keys(ws).forEach(cell => {
        if (ws[cell].s) {
            ws[cell].s = { ...ws[cell].s, ...rightAlignStyle };
        } else {
            ws[cell].s = rightAlignStyle;
        }
    });

    // Dodawanie arkuszu do excela
    xlsx.utils.book_append_sheet(wb, ws, 'Całość');

    // wyłączenie przeglądarki
    await browser.close();

    // Zapisywanie excela
    // xlsx.writeFile(wb, 'tabela.xlsx');
    xlsxStyle.writeFile(wb, 'tabela.xlsx');
    console.log('Plik Excel został zapisany jako tabela.xlsx');
};

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

askDate(`Wprowadź do terminala datę początkową w formacie DD-MM-RRRR, np.: "01-01-2024". Nie może być wcześniejsza niż ${earliestDate}: `, (startDate) => {
    const parsedStartDate = parseDate(startDate);
    const parsedEarliestDate = parseDate(earliestDate);

    if (parsedStartDate < parsedEarliestDate) {
        console.log(`Podana data jest wcześniejsza niż ${earliestDate}: Uruchom program jeszcze raz.`);
        rl.close();
        return;
    }

    askDate('Wprowadź datę końcową w formacie DD-MM-YYYY: ', (endDate) => {
        const parsedEndDate = parseDate(endDate);

        if (parsedStartDate > parsedEndDate) {
            console.log('Podana data końcowa jest w kalendarzu przed datą początkową. Uruchom program jeszcze raz.');
            rl.close();
            return;
        }
        
        makeExcel(startDate, endDate);
        rl.close();
    });
});