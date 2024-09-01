const moment = require('moment');

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

// Funkcja przerabiająca datę zapisaną w stringu w obiekt
function parseDate(dateStr) {
    const [day, month, year] = dateStr.split('-').map(Number);
    return new Date(year, month - 1, day);
}

// Funkcja sprawdzająca dzisiejszy dzień i zwracająca datę 60 dni wcześniej - pierwszy dzień z danymi na stronie
function getDate59DaysBefore() {
    const currentDate = new Date();
    currentDate.setDate(currentDate.getDate() - 60);

    const day = String(currentDate.getDate()).padStart(2, '0');
    const month = String(currentDate.getMonth() + 1).padStart(2, '0');
    const year = currentDate.getFullYear();

    return `${day}-${month}-${year}`;
}
// console.log(getDate59DaysBefore());

// Funkcja przerabiają datę zapisaną w stringu na datę z poprzedniego dnia
function getPrevDay(date) {
    const [day, month, year] = date.split('-').map(Number);
    const nextDay = new Date(year, month - 1, day);
    nextDay.setDate(nextDay.getDate() - 1);
    
    // Formatowanie z powrotem do stringów
    const newDay = String(nextDay.getDate()).padStart(2, '0');
    const newMonth = String(nextDay.getMonth() + 1).padStart(2, '0');
    const newYear = nextDay.getFullYear();
    
    return `${newDay}-${newMonth}-${newYear}`;
}

// Funkcja przerabiają datę zapisaną w stringu na datę z kolejnego dnia
function getNextDay(date) {
    const [day, month, year] = date.split('-').map(Number);
    const nextDay = new Date(year, month - 1, day);
    nextDay.setDate(nextDay.getDate() + 1);
    
    // Formatowanie z powrotem do stringów
    const newDay = String(nextDay.getDate()).padStart(2, '0');
    const newMonth = String(nextDay.getMonth() + 1).padStart(2, '0');
    const newYear = nextDay.getFullYear();
    
    return `${newDay}-${newMonth}-${newYear}`;
}

module.exports = {
    generateDates,
    calculateColumnWidths,
    mergeTablesSideBySide,
    parseDate,
    getDate59DaysBefore,
    getPrevDay,
    getNextDay
};