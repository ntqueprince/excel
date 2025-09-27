// Global variables
let originalData = [];
let currentData = [];
let columnTypes = {};
let sortState = { column: null, direction: 'asc' };
let activeFilters = {};

// Pagination variables
const rowsPerPage = 100;
let currentPage = 1;

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
    headers.forEach((header, index) => {
    const sampleValues = rows.slice(0, 10).map(row => row[index]).filter(val => val != null && val !== '');

    // ✅ अगर header में % या 'csat' लिखा है तो उसे forcefully percentage column मान लो
    if (header.toLowerCase().includes('%') || header.toLowerCase().includes('csat')) {
        columnTypes[header] = 'percentage';
    } else {
        columnTypes[header] = determineColumnType(sampleValues);
    }
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
    currentPage = 1;
    renderTable();
    updateStats();
    renderPaginationControls();

    document.getElementById('controlsSection').style.display = 'block';
    document.getElementById('tableContainer').style.display = 'block';
    document.getElementById('paginationControls').style.display = 'flex';
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

    // ✅ अब सिर्फ percentage columns पर ही color लागू होगा
    if (type === 'percentage') {
        if (numValue >= 95) return 'cell-dark-green';
        if (numValue >= 90) return 'cell-light-green';
        if (numValue >= 80) return 'cell-yellow';
        if (numValue >= 70) return 'cell-orange';
        if (numValue >= 50) return 'cell-light-red';
        if (numValue >= 30) return 'cell-red';
        return 'cell-dark-red';
    }

    // ✅ बाकी सब पर कोई color नहीं
    return '';
}

// Function to capture the table and copy it to the clipboard
function captureTable() {
    const tableContainer = document.getElementById('tableContainer');
    const tableWrapper = document.querySelector('.table-wrapper');
    const originalMaxHeight = tableWrapper.style.maxHeight;

    showLoading();

    const rowHeight = 35;
    const desiredHeight = 80 * rowHeight;
    tableWrapper.style.maxHeight = `${desiredHeight}px`;
    tableWrapper.style.overflow = 'hidden';

    setTimeout(() => {
        html2canvas(tableContainer, {
            useCORS: true,
            logging: false,
            scale: 2
        }).then(canvas => {
            tableWrapper.style.maxHeight = originalMaxHeight;
            tableWrapper.style.overflow = '';

            canvas.toBlob(blob => {
                const item = new ClipboardItem({'image/png': blob});
                
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
            tableWrapper.style.maxHeight = originalMaxHeight;
            tableWrapper.style.overflow = '';

            console.error('Failed to capture table: ', err);
            alert('Failed to capture table image. Please try again.');
            hideLoading();
        });
    }, 200);
}

// Render table (with pagination)
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

    thead.innerHTML = '';
    tbody.innerHTML = '';

    const headerRow = document.createElement('tr');
    
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

        const sortBtn = document.createElement('button');
        sortBtn.className = 'sort-button';
        sortBtn.innerHTML = '<i class="fas fa-sort"></i>';
        sortBtn.onclick = () => sortColumn(header);
        columnNameGroup.appendChild(sortBtn);
        
        headerDiv.appendChild(columnNameGroup);
        
        if (columnTypes[header] === 'text') {
            const filterButton = document.createElement('button');
            filterButton.className = 'filter-button';
            filterButton.innerHTML = '<i class="fas fa-filter"></i>';
            filterButton.onclick = (e) => toggleFilterMenu(e, header);
            headerDiv.appendChild(filterButton);
        }

        th.appendChild(headerDiv);
        headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);

    const startIndex = (currentPage - 1) * rowsPerPage;
    const endIndex = startIndex + rowsPerPage;
    const paginatedData = currentData.slice(startIndex, endIndex);

    paginatedData.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        
        const tdCheckbox = document.createElement('td');
        tdCheckbox.className = 'row-checkbox-cell';
        const rowCheckbox = document.createElement('input');
        rowCheckbox.type = 'checkbox';
        rowCheckbox.className = 'row-checkbox';
        rowCheckbox.setAttribute('data-row-index', startIndex + rowIndex); // Use global index
        rowCheckbox.onchange = updateSelectedRowCount;
        tdCheckbox.appendChild(rowCheckbox);
        tr.appendChild(tdCheckbox);
        
        headers.forEach(header => {
            const td = document.createElement('td');
            const value = row[header];
            
            if (columnTypes[header] === 'numeric' || columnTypes[header] === 'percentage') {
                const numericValue = getNumericValue(value);
                if (numericValue !== null) {
                    const originalStringValue = String(value).trim();
                    td.textContent = numericValue.toFixed(2) + (columnTypes[header] === 'percentage' ? '%' : '');
                } else {
                    td.textContent = value;
                }
            } else {
                td.textContent = value;
            }

            const colorClass = applyCellColoring(value, columnTypes[header]);
            if (colorClass) {
                td.className = colorClass;
            }
            
            tr.appendChild(td);
        });
        
        tbody.appendChild(tr);
    });

    updateStats();
    updateSelectedColumnCount();
    updateSelectedRowCount();
    updateSortIndicators();
}

// Toggle filter menu
function toggleFilterMenu(e, columnName) {
    e.stopPropagation();

    document.querySelectorAll('.filter-menu.show').forEach(menu => {
        if (menu.id !== `filter-menu-${columnName}`) {
            menu.classList.remove('show');
        }
    });

    let filterMenu = document.getElementById(`filter-menu-${columnName}`);
    if (!filterMenu) {
        filterMenu = document.createElement('div');
        filterMenu.className = 'filter-menu';
        filterMenu.id = `filter-menu-${columnName}`;
        e.target.parentNode.appendChild(filterMenu);
        populateFilterMenu(columnName, filterMenu);
    }
    
    filterMenu.onclick = (event) => event.stopPropagation();
    
    filterMenu.classList.toggle('show');

    document.addEventListener('click', function closeMenu(event) {
        if (!filterMenu.contains(event.target) && !e.target.contains(event.target)) {
            filterMenu.classList.remove('show');
            document.removeEventListener('click', closeMenu);
        }
    });
}

// Populate filter menu with unique values
// Populate filter menu with unique values
// UPDATED: Populate filter menu with unique values
function populateFilterMenu(columnName, menuElement) {
    menuElement.innerHTML = '';
    const uniqueValues = [...new Set(originalData.map(row => row[columnName]))].filter(v => v !== null && v !== '').sort();

    // Add Select/Unselect All buttons
    const controlsDiv = document.createElement('div');
    controlsDiv.className = 'filter-controls';
    
    const selectAllBtn = document.createElement('button');
    selectAllBtn.textContent = 'Select All';
    selectAllBtn.className = 'btn btn-filter-control';
    selectAllBtn.onclick = () => {
        menuElement.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
            checkbox.checked = true;
        });
    };

    const unselectAllBtn = document.createElement('button');
    unselectAllBtn.textContent = 'Unselect All';
    unselectAllBtn.className = 'btn btn-filter-control';
    unselectAllBtn.onclick = () => {
        menuElement.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
            checkbox.checked = false;
        });
    };

    controlsDiv.appendChild(selectAllBtn);
    controlsDiv.appendChild(unselectAllBtn);
    menuElement.appendChild(controlsDiv);

    // 🔥 Search bar
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = 'Search...';
    searchInput.className = 'filter-search';
    menuElement.appendChild(searchInput);

    // 🔥 Apply Filter button ko search ke niche rakha
    const applyBtn = document.createElement('button');
    applyBtn.textContent = 'Apply Filter';
    applyBtn.className = 'filter-apply-btn';
    applyBtn.onclick = () => {
        applyFilters(columnName, menuElement);
        menuElement.classList.remove('show');
    };
    menuElement.appendChild(applyBtn);

    // Search logic
    searchInput.onkeyup = () => {
        const filter = searchInput.value.toLowerCase();
        menuElement.querySelectorAll('.filter-option').forEach(option => {
            const value = option.textContent.toLowerCase();
            option.style.display = value.includes(filter) ? '' : 'none';
        });
    };

    // Options container
    const optionsContainer = document.createElement('div');
    optionsContainer.className = 'filter-options-container';

    const currentActiveValues = activeFilters[columnName] || null;
    uniqueValues.forEach(value => {
        const option = document.createElement('label');
        option.className = 'filter-option';
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = value;
        checkbox.checked = currentActiveValues ? currentActiveValues.includes(value) : true;
        
        option.appendChild(checkbox);
        option.appendChild(document.createTextNode(value));
        optionsContainer.appendChild(option);
    });
    menuElement.appendChild(optionsContainer);
}

    // Always show Apply Filter button
    const applyBtn = document.createElement('button');
    applyBtn.textContent = 'Apply Filter';
    applyBtn.className = 'filter-apply-btn';
    applyBtn.style.marginTop = '8px';
    applyBtn.onclick = () => {
        applyFilters(columnName, menuElement);
        menuElement.classList.remove('show');
    };
    menuElement.appendChild(applyBtn);
// Apply filters
function applyFilters(columnName, menuElement) {
    const selectedValues = [];
    menuElement.querySelectorAll('.filter-options-container input[type="checkbox"]:checked').forEach(checkbox => {
        selectedValues.push(checkbox.value);
    });

    if (selectedValues.length === 0) {
        delete activeFilters[columnName];
    } else {
        activeFilters[columnName] = selectedValues;
    }

    currentData = originalData.filter(row => {
        for (const col in activeFilters) {
            if (!activeFilters[col].includes(row[col])) {
                return false;
            }
        }
        return true;
    });

    currentPage = 1;
    renderTable();
    renderPaginationControls();
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
        
        originalData = originalData.map(row => {
            const newRow = { ...row };
            columnsToDelete.forEach(colName => {
                delete newRow[colName];
            });
            return newRow;
        });

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
        renderPaginationControls();
    }
}

function deleteSelectedRows() {
    const selectedCheckboxes = document.querySelectorAll('.row-checkbox:checked');
    if (selectedCheckboxes.length === 0) {
        alert('Please select at least one row to delete.');
        return;
    }

    if (confirm(`Are you sure you want to delete the selected ${selectedCheckboxes.length} rows?`)) {
        const rowsToDeleteIndices = Array.from(selectedCheckboxes).map(checkbox => parseInt(checkbox.getAttribute('data-row-index')));
        
        originalData = originalData.filter((row, index) => !rowsToDeleteIndices.includes(index));

        currentData = originalData.filter(row => {
            for (const col in activeFilters) {
                if (!activeFilters[col].includes(row[col])) {
                    return false;
                }
            }
            return true;
        });

        const totalPages = Math.ceil(currentData.length / rowsPerPage);
        if (currentPage > totalPages && totalPages > 0) {
            currentPage = totalPages;
        } else if (totalPages === 0) {
            currentPage = 1;
        }

        renderTable();
        renderPaginationControls();
    }
}

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

// Pagination Logic
function renderPaginationControls() {
    const totalPages = Math.ceil(currentData.length / rowsPerPage);
    const pageNumbersDiv = document.getElementById('pageNumbers');
    pageNumbersDiv.innerHTML = '';
    
    const maxPagesToShow = 5;
    let startPage = Math.max(1, currentPage - Math.floor(maxPagesToShow / 2));
    let endPage = Math.min(totalPages, startPage + maxPagesToShow - 1);

    if (endPage - startPage + 1 < maxPagesToShow) {
        startPage = Math.max(1, endPage - maxPagesToShow + 1);
    }
    
    for (let i = startPage; i <= endPage; i++) {
        const pageBtn = document.createElement('button');
        pageBtn.textContent = i;
        pageBtn.className = `btn btn-secondary page-number ${i === currentPage ? 'active' : ''}`;
        pageBtn.onclick = () => changePage(i);
        pageNumbersDiv.appendChild(pageBtn);
    }

    document.getElementById('prevPageBtn').disabled = currentPage === 1;
    document.getElementById('nextPageBtn').disabled = currentPage === totalPages;
}

function changePage(pageNumber) {
    if (pageNumber >= 1 && pageNumber <= Math.ceil(currentData.length / rowsPerPage)) {
        currentPage = pageNumber;
        renderTable();
        renderPaginationControls();
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
