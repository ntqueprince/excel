// Global variables
let originalData = [];
let currentData = [];
let columnTypes = {};
let sortState = { column: null, direction: 'asc' };

// Tab switching
function switchTab(tabName) {
    document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    
    document.querySelector(`[onclick="switchTab('${tabName}')"]`).classList.add('active');
    document.getElementById(`${tabName}-tab`).classList.add('active');
}

// File upload handlers
function handleDragOver(e) {
    e.preventDefault();
    e.currentTarget.classList.add('dragover');
}

function handleDragLeave(e) {
    e.currentTarget.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    e.currentTarget.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function handleFileUpload(e) {
    const file = e.target.files[0];
    if (file) {
        processFile(file);
    }
}

// File processing
function processFile(file) {
    showLoading();
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (jsonData.length > 0) {
                parseAndDisplayData(jsonData);
            } else {
                alert('The file appears to be empty or contains no readable data.');
            }
        } catch (error) {
            alert('Error reading file: ' + error.message);
        } finally {
            hideLoading();
        }
    };
    reader.readAsArrayBuffer(file);
}

// Data processing from textarea
function processData() {
    const textarea = document.getElementById('pasteArea');
    const text = textarea.value.trim();
    
    if (!text) {
        alert('Please paste some data first.');
        return;
    }

    showLoading();
    
    try {
        const lines = text.split('\n').filter(line => line.trim());
        const data = lines.map(line => {
            // Try tab-separated first, then comma-separated
            return line.includes('\t') ? line.split('\t') : line.split(',');
        });
        
        parseAndDisplayData(data);
    } catch (error) {
        alert('Error parsing data: ' + error.message);
    } finally {
        hideLoading();
    }
}

// Parse and display data
function parseAndDisplayData(rawData) {
    if (rawData.length === 0) return;

    // First row is headers
    const headers = rawData[0];
    const rows = rawData.slice(1);

    // Determine column types
    columnTypes = {};
    headers.forEach((header, index) => {
        const sampleValues = rows.slice(0, 10).map(row => row[index]).filter(val => val != null && val !== '');
        columnTypes[header] = determineColumnType(sampleValues);
    });

    // Convert data to objects
    originalData = rows.map(row => {
        const obj = {};
        headers.forEach((header, index) => {
            obj[header] = row[index] || '';
        });
        return obj;
    });

    currentData = [...originalData];
    renderTable();
    updateStats();
    
    document.getElementById('controlsSection').style.display = 'block';
    document.getElementById('tableContainer').style.display = 'block';
}

// Determine column type (numeric, percentage, or text)
function determineColumnType(values) {
    if (values.length === 0) return 'text';
    
    let numericCount = 0;
    let percentageCount = 0;
    
    values.forEach(value => {
        const strValue = String(value).trim();
        if (strValue.endsWith('%')) {
            percentageCount++;
        } else if (!isNaN(parseFloat(strValue)) && isFinite(strValue)) {
            numericCount++;
        }
    });
    
    const total = values.length;
    if (percentageCount / total > 0.5) return 'percentage';
    if (numericCount / total > 0.5) return 'numeric';
    return 'text';
}

// Get numeric value from cell
function getNumericValue(value) {
    if (value == null || value === '') return null;
    
    const strValue = String(value).trim();
    if (strValue.endsWith('%')) {
        return parseFloat(strValue.slice(0, -1));
    }
    
    const num = parseFloat(strValue);
    return isNaN(num) ? null : num;
}

// Function to capture the table and copy it to the clipboard
function captureTable() {
    const tableContainer = document.getElementById('tableContainer');
    const tableWrapper = document.querySelector('.table-wrapper');
    const originalMaxHeight = tableWrapper.style.maxHeight;

    showLoading();

    // Temporarily increase the max-height to ensure all 80 rows are visible
    const rowHeight = 35; // approximate height of one table row
    const desiredHeight = 80 * rowHeight;
    tableWrapper.style.maxHeight = `${desiredHeight}px`;
    tableWrapper.style.overflow = 'hidden'; // Hide the scrollbar

    setTimeout(() => {
        html2canvas(tableContainer, {
            useCORS: true,
            logging: false,
            scale: 2 // Improved scale for better resolution
        }).then(canvas => {
            // Restore the original height and overflow
            tableWrapper.style.maxHeight = originalMaxHeight;
            tableWrapper.style.overflow = '';

            // Convert canvas to a blob
            canvas.toBlob(blob => {
                // Create a new ClipboardItem
                const item = new ClipboardItem({'image/png': blob});
                
                // Write the item to the clipboard
                navigator.clipboard.write([item]).then(() => {
                    alert('Table image copied to clipboard! You can now paste it.');
                    hideLoading();
                }).catch(err => {
                    console.error('Could not copy image to clipboard: ', err);
                    alert('Failed to copy image to clipboard. Please try again.');
                    hideLoading();
                });
            });
        }).catch(err => {
            // Restore on error
            tableWrapper.style.maxHeight = originalMaxHeight;
            tableWrapper.style.overflow = '';

            console.error('Failed to capture table: ', err);
            alert('Failed to capture table image. Please try again.');
            hideLoading();
        });
    }, 200); // Increased delay for better rendering
}


// Render table
function renderTable() {
    if (currentData.length === 0) return;

    const headers = Object.keys(currentData[0]);
    const thead = document.getElementById('tableHead');
    const tbody = document.getElementById('tableBody');

    // Clear existing content
    thead.innerHTML = '';
    tbody.innerHTML = '';

    // Create header row
    const headerRow = document.createElement('tr');
    
    headers.forEach(header => {
        const th = document.createElement('th');
        const headerDiv = document.createElement('div');
        headerDiv.className = 'column-header';
        
        const columnNameGroup = document.createElement('div');
        columnNameGroup.className = 'column-name-group';

        const headerText = document.createElement('span');
        headerText.textContent = header;
        columnNameGroup.appendChild(headerText);

        // Sort button (only for numeric/percentage columns)
        if (columnTypes[header] === 'numeric' || columnTypes[header] === 'percentage') {
            const sortBtn = document.createElement('button');
            sortBtn.className = 'sort-button';
            sortBtn.innerHTML = '<i class="fas fa-sort"></i>';
            sortBtn.onclick = () => sortColumn(header);
            columnNameGroup.appendChild(sortBtn);
        }
        
        headerDiv.appendChild(columnNameGroup);
        
        const actions = document.createElement('div');
        actions.className = 'column-actions';
        
        // Delete button
        const deleteBtn = document.createElement('button');
        deleteBtn.className = 'column-delete';
        deleteBtn.innerHTML = '<i class="fas fa-times"></i>';
        deleteBtn.onclick = () => deleteColumn(header);
        actions.appendChild(deleteBtn);

        headerDiv.appendChild(actions);
        th.appendChild(headerDiv);
        headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);

    // Create data rows
    currentData.forEach(row => {
        const tr = document.createElement('tr');
        
        headers.forEach(header => {
            const td = document.createElement('td');
            const value = row[header];
            
            // Format numeric values to one decimal place
            if (columnTypes[header] === 'numeric' || columnTypes[header] === 'percentage') {
                const numericValue = getNumericValue(value);
                if (numericValue !== null) {
                    td.textContent = numericValue.toFixed(1) + (columnTypes[header] === 'percentage' ? '%' : '');
                } else {
                    td.textContent = value;
                }
            } else {
                td.textContent = value;
            }
            
            tr.appendChild(td);
        });
        
        tbody.appendChild(tr);
    });

    updateStats();
}

// Column deletion
function deleteColumn(columnName) {
    if (confirm(`Are you sure you want to delete the "${columnName}" column?`)) {
        currentData = currentData.map(row => {
            const newRow = { ...row };
            delete newRow[columnName];
            return newRow;
        });
        
        delete columnTypes[columnName];
        renderTable();
    }
}

// Sorting
function sortColumn(columnName) {
    if (sortState.column === columnName) {
        sortState.direction = sortState.direction === 'asc' ? 'desc' : 'asc';
    } else {
        sortState.column = columnName;
        sortState.direction = 'asc';
    }

    currentData.sort((a, b) => {
        const aValue = getNumericValue(a[columnName]);
        const bValue = getNumericValue(b[columnName]);
        
        if (aValue === null && bValue === null) return 0;
        if (aValue === null) return 1;
        if (bValue === null) return -1;
        
        const comparison = aValue - bValue;
        return sortState.direction === 'asc' ? comparison : -comparison;
    });

    renderTable();
    updateSortIndicators();
}

function updateSortIndicators() {
    document.querySelectorAll('.sort-button').forEach(btn => {
        btn.classList.remove('active');
        btn.innerHTML = '<i class="fas fa-sort"></i>';
    });

    if (sortState.column) {
        const activeButton = document.querySelector(`[onclick*="${sortState.column}"]`);
        if (activeButton) {
            activeButton.classList.add('active');
            const icon = sortState.direction === 'asc' ? 'fa-sort-up' : 'fa-sort-down';
            activeButton.innerHTML = `<i class="fas ${icon}"></i>`;
        }
    }
}

// Export functionality
function exportToExcel() {
    if (currentData.length === 0) {
        alert('No data to export.');
        return;
    }

    const headers = Object.keys(currentData[0]);
    const wsData = [headers, ...currentData.map(row => headers.map(header => row[header]))];
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    
    XLSX.writeFile(wb, `excel_data_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

// Utility functions
function showLoading() {
    document.getElementById('loading').style.display = 'block';
}

function hideLoading() {
    document.getElementById('loading').style.display = 'none';
}

function updateStats() {
    document.getElementById('rowCount').textContent = currentData.length;
    document.getElementById('columnCount').textContent = Object.keys(currentData[0] || {}).length;
}