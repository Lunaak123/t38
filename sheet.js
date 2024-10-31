let excelData = []; // Placeholder for Excel data
let filteredData = []; // Placeholder for filtered data

// Load the Google Sheets file when the page loads
document.addEventListener('DOMContentLoaded', async () => {
    const urlParams = new URLSearchParams(window.location.search);
    const fileUrl = urlParams.get('fileUrl');

    if (fileUrl) {
        await loadExcelData(fileUrl);
    }
});

// Function to load Excel data
async function loadExcelData(url) {
    const response = await fetch(url);
    const data = await response.arrayBuffer();
    const workbook = XLSX.read(data);
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
    displayData(excelData);
}

// Function to display data in the table
function displayData(data) {
    const sheetContent = document.getElementById('sheet-content');
    sheetContent.innerHTML = '';

    const table = document.createElement('table');
    data.forEach((row) => {
        const tr = document.createElement('tr');
        row.forEach((cell) => {
            const td = document.createElement('td');
            td.textContent = cell === null ? 'NULL' : cell;
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });
    sheetContent.appendChild(table);
}

// Apply operation when button is clicked
document.getElementById('apply-operation').addEventListener('click', applyOperation);

// Highlight data based on selections
function applyOperation() {
    const primaryColumn = document.getElementById('primary-column').value.trim().toUpperCase();
    const operationColumnsInput = document.getElementById('operation-columns').value.trim();
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    const rowRangeFrom = parseInt(document.getElementById('row-range-from').value, 10);
    const rowRangeTo = parseInt(document.getElementById('row-range-to').value, 10);

    if (!primaryColumn || !operationColumnsInput) {
        alert('Please enter the primary column and columns to operate on.');
        return;
    }

    const operationColumns = operationColumnsInput.split(',').map(col => col.trim().toUpperCase());
    filteredData = excelData.filter((row, index) => {
        if (index < rowRangeFrom - 1 || index > rowRangeTo - 1) return false;

        const isPrimaryNull = row[primaryColumn.charCodeAt(0) - 65] === null || row[primaryColumn.charCodeAt(0) - 65] === "";
        const columnChecks = operationColumns.map(col => operation === 'null' ? row[col.charCodeAt(0) - 65] === null || row[col.charCodeAt(0) - 65] === "" : row[col.charCodeAt(0) - 65] !== null && row[col.charCodeAt(0) - 65] !== "");

        return operationType === 'and' ? !isPrimaryNull && columnChecks.every(Boolean) : !isPrimaryNull && columnChecks.some(Boolean);
    });

    displayData(filteredData);
    highlightData(rowRangeFrom, rowRangeTo, primaryColumn, operationColumns);
}

// Highlight selected rows and columns
function highlightData(rowRangeFrom, rowRangeTo, primaryColumn, operationColumns) {
    const table = document.querySelector('table');
    if (!table) return;

    const rows = table.querySelectorAll('tr');

    rows.forEach((row, index) => {
        if (index >= rowRangeFrom - 1 && index <= rowRangeTo - 1) {
            const primaryCell = row.cells[primaryColumn.charCodeAt(0) - 65]; // Convert column letter to index
            const shouldHighlight = operationColumns.some(col => {
                const colCell = row.cells[col.charCodeAt(0) - 65]; // Get cell for operation
                return colCell && (colCell.textContent.trim() === '' || colCell.textContent.trim() === 'NULL');
            });

            if (shouldHighlight) {
                row.style.backgroundColor = '#d1e7dd'; // Highlight color
            } else {
                row.style.backgroundColor = ''; // Reset color
            }
        } else {
            row.style.backgroundColor = ''; // Reset color for rows outside the range
        }
    });
}

// Download functionality
document.getElementById('download-button').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'flex';
});

// Confirm download button
document.getElementById('confirm-download').addEventListener('click', () => {
    const filename = document.getElementById('filename').value || 'downloaded_file';
    const format = document.getElementById('file-format').value;

    // Download logic based on format
    if (format === 'xlsx') {
        const ws = XLSX.utils.aoa_to_sheet(filteredData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        XLSX.writeFile(wb, `${filename}.xlsx`);
    } else if (format === 'csv') {
        const csvContent = XLSX.utils.sheet_to_csv(XLSX.utils.aoa_to_sheet(filteredData));
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.setAttribute('href', url);
        link.setAttribute('download', `${filename}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    document.getElementById('download-modal').style.display = 'none';
});

// Close modal
document.getElementById('close-modal').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'none';
});
