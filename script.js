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

// Apply auto-coloring
function applyCellColoring(value, type) {
    const numValue = getNumericValue(value);
    if (numValue === null) return '';

    if (type === 'percentage') {
        if (numValue >= 95) return 'cell-dark-green';
        if (numValue >= 90) return 'cell-light-green';
        if (numValue >= 80) return 'cell-yellow';
        if (numValue >= 70) return 'cell-orange';
        if (numValue >= 50) return 'cell-light-red';
        if (numValue >= 30) return 'cell-red';
        return 'cell-dark-red';
    } else if (type === 'numeric') {
        // Find the original column name based on a sample value from the current row
        const rowKeys = Object.keys(currentData[0]);
        let originalColumn;
        for (const key of rowKeys) {
            const sampleValue = currentData[0][key];
            if (getNumericValue(sampleValue) === numValue) {
                originalColumn = key;
                break;
            }
        }
        
        if (!originalColumn) return '';

        const allValues = originalData.map(row => getNumericValue(row[originalColumn])).filter(val => val !== null);
        if (allValues.length === 0) return '';
        
        const min = Math.min(...allValues);
        const max = Math.max(...allValues);
        const range = max - min;
        
        if (range === 0) return 'cell-dark-green';
        
        const percentile = (numValue - min) / range;
        
        if (percentile >= 0.9) return 'cell-dark-green';
        if (percentile >= 0.75) return 'cell-light-green';
        if (percentile >= 0.5) return 'cell-yellow';
        if (percentile >= 0.25) return 'cell-orange';
        return 'cell-light-red';
    }
    
    return '';
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
    if (currentData.length === 0) {
        document.getElementById('tableHead').innerHTML = '';
        document.getElementById('tableBody').innerHTML = '';
        updateStats();
        return;
    }

    const headers = Object.keys(currentData[0]);
    const thead = document.getElementById('tableHead');
    const tbody = document.getElementById('tableBody');

    // Clear existing content
    thead.innerHTML = '';
    tbody.innerHTML = '';

    // Create header row
    const headerRow = document.createElement('tr');
    
    // Add master checkbox for selecting all columns
    const selectAllColumnsTh = document.createElement('th');
    selectAllColumnsTh.className = 'checkbox-column';
    const selectAllColumnsCheckbox = document.createElement('input');
    selectAllColumnsCheckbox.type = 'checkbox';
    selectAllColumnsCheckbox.id = 'selectAllColumns';
    selectAllColumnsCheckbox.title = 'Select All Columns';
    selectAllColumnsCheckbox.onchange = toggleSelectAllColumns;
    selectAllColumnsTh.appendChild(selectAllColumnsCheckbox);
    headerRow.appendChild(selectAllColumnsTh);

    headers.forEach(header => {
        const th = document.createElement('th');
        const headerDiv = document.createElement('div');
        headerDiv.className = 'column-header';

        const columnCheckbox = document.createElement('input');
        columnCheckbox.type = 'checkbox';
        columnCheckbox.className = 'column-checkbox';
        columnCheckbox.setAttribute('data-column-name', header);
        columnCheckbox.onchange = updateSelectedColumnCount;
        
        const headerText = document.createElement('span');
        headerText.textContent = header;
        
        const columnNameGroup = document.createElement('div');
        columnNameGroup.className = 'column-name-group';
        columnNameGroup.appendChild(columnCheckbox);
        columnNameGroup.appendChild(headerText);

        // Sort button for all columns
        const sortBtn = document.createElement('button');
        sortBtn.className = 'sort-button';
        sortBtn.innerHTML = '<i class="fas fa-sort"></i>';
        sortBtn.onclick = () => sortColumn(header);
        columnNameGroup.appendChild(sortBtn);
        
        headerDiv.appendChild(columnNameGroup);
        th.appendChild(headerDiv);
        headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);

    // Create data rows
    currentData.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        
        // Add a checkbox cell for each row
        const tdCheckbox = document.createElement('td');
        tdCheckbox.className = 'row-checkbox-cell';
        const rowCheckbox = document.createElement('input');
        rowCheckbox.type = 'checkbox';
        rowCheckbox.className = 'row-checkbox';
        rowCheckbox.setAttribute('data-row-index', rowIndex);
        rowCheckbox.onchange = updateSelectedRowCount;
        tdCheckbox.appendChild(rowCheckbox);
        tr.appendChild(tdCheckbox);
        
        headers.forEach(header => {
            const td = document.createElement('td');
            const value = row[header];
            
            // Format numeric values
            if (columnTypes[header] === 'numeric' || columnTypes[header] === 'percentage') {
                const numericValue = getNumericValue(value);
                if (numericValue !== null) {
                    const originalStringValue = String(value).trim();
                    if (originalStringValue.includes('.') || originalStringValue.endsWith('%')) {
                        td.textContent = numericValue.toFixed(1) + (columnTypes[header] === 'percentage' ? '%' : '');
                    } else {
                        // Display integer value without .0
                        td.textContent = numericValue.toString() + (columnTypes[header] === 'percentage' ? '%' : '');
                    }
                } else {
                    td.textContent = value;
                }
            } else {
                td.textContent = value;
            }

            // Apply auto-coloring
            const colorClass = applyCellColoring(value, columnTypes[header]);
            if (colorClass) {
                td.className = colorClass;
            }
            
            tr.appendChild(td);
        });
        
        tbody.appendChild(tr);
    });

    updateStats();
    updateSelectedColumnCount(); // Initial update for columns
    updateSelectedRowCount(); // Initial update for rows
}

function toggleSelectAllColumns() {
    const isChecked = document.getElementById('selectAllColumns').checked;
    document.querySelectorAll('.column-checkbox').forEach(checkbox => {
        checkbox.checked = isChecked;
    });
    updateSelectedColumnCount();
}

function updateSelectedColumnCount() {
    const checkedCount = document.querySelectorAll('.column-checkbox:checked').length;
    document.getElementById('selectedColumnCount').textContent = checkedCount;
}

function deleteSelectedColumns() {
    const selectedCheckboxes = document.querySelectorAll('.column-checkbox:checked');
    if (selectedCheckboxes.length === 0) {
        alert('Please select at least one column to delete.');
        return;
    }

    if (confirm(`Are you sure you want to delete the selected ${selectedCheckboxes.length} columns?`)) {
        const columnsToDelete = Array.from(selectedCheckboxes).map(checkbox => checkbox.getAttribute('data-column-name'));
        
        currentData = currentData.map(row => {
            const newRow = { ...row };
            columnsToDelete.forEach(colName => {
                delete newRow[colName];
            });
            return newRow;
        });

        columnsToDelete.forEach(colName => {
            delete columnTypes[colName];
        });
        
        renderTable();
    }
}

// NEW FUNCTION: Delete selected rows
function deleteSelectedRows() {
    const selectedCheckboxes = document.querySelectorAll('.row-checkbox:checked');
    if (selectedCheckboxes.length === 0) {
        alert('Please select at least one row to delete.');
        return;
    }

    if (confirm(`Are you sure you want to delete the selected ${selectedCheckboxes.length} rows?`)) {
        const rowsToDeleteIndices = Array.from(selectedCheckboxes).map(checkbox => parseInt(checkbox.getAttribute('data-row-index')));
        
        // Filter out the rows to be deleted
        currentData = currentData.filter((row, index) => !rowsToDeleteIndices.includes(index));
        
        renderTable();
    }
}

// NEW FUNCTION: Update selected row count
function updateSelectedRowCount() {
    const checkedCount = document.querySelectorAll('.row-checkbox:checked').length;
    document.getElementById('selectedRowCount').textContent = checkedCount;
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
        const aValue = a[columnName];
        const bValue = b[columnName];

        const type = columnTypes[columnName];

        if (type === 'numeric' || type === 'percentage') {
            const aNum = getNumericValue(aValue);
            const bNum = getNumericValue(bValue);

            if (aNum === null && bNum === null) return 0;
            if (aNum === null) return 1;
            if (bNum === null) return -1;
            
            const comparison = aNum - bNum;
            return sortState.direction === 'asc' ? comparison : -comparison;
        } else {
            // Text sorting logic
            const comparison = String(aValue).localeCompare(String(bValue), undefined, { sensitivity: 'base' });
            return sortState.direction === 'asc' ? comparison : -comparison;
        }
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
