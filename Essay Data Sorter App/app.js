// Global data storage
let allData = [];
let filteredData = [];
let selectedRowIndices = new Set(); // Track selected rows by originalIndex
let checkboxStates = {}; // Store checkbox states: {index: {copypaste: bool, reviewed: bool, export: bool}}
let lastSelectedIndex = null; // Track last selected index for shift-click range selection

// DOM elements
const pasteArea = document.getElementById('pasteArea');
const parsePasteBtn = document.getElementById('parsePasteBtn');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const parseStatus = document.getElementById('parseStatus');
const filterSection = document.getElementById('filterSection');
const resultsSection = document.getElementById('resultsSection');
const resultsCount = document.getElementById('resultsCount');
const resultsTableBody = document.getElementById('resultsTableBody');
const clearFiltersBtn = document.getElementById('clearFiltersBtn');
const clearDataBtn = document.getElementById('clearDataBtn');
const clearSelectionBtn = document.getElementById('clearSelectionBtn');
const selectAllBtn = document.getElementById('selectAllBtn');
const copyExportBtn = document.getElementById('copyExportBtn');
const fromDateInput = document.getElementById('fromDate');
const toDateInput = document.getElementById('toDate');

// Filter checkbox containers
const classCheckboxes = document.getElementById('classCheckboxes');
const essayTypeCheckboxes = document.getElementById('essayTypeCheckboxes');
const teacherCheckboxes = document.getElementById('teacherCheckboxes');
const evaluatorCheckboxes = document.getElementById('evaluatorCheckboxes');
const evaluationCheckboxes = document.getElementById('evaluationCheckboxes');

// Event listeners
parsePasteBtn.addEventListener('click', handlePasteParse);
fileInput.addEventListener('change', handleFileUpload);
clearFiltersBtn.addEventListener('click', clearAllFilters);
clearDataBtn.addEventListener('click', clearAllData);
clearSelectionBtn.addEventListener('click', clearSelection);
selectAllBtn.addEventListener('click', selectAllRows);
copyExportBtn.addEventListener('click', copyExportNames);
fromDateInput.addEventListener('change', applyFilters);
toDateInput.addEventListener('change', applyFilters);

// Parse pasted text
function handlePasteParse() {
    const text = pasteArea.value.trim();
    if (!text) {
        showStatus('Please paste some data first.', 'error');
        return;
    }
    
    const parsed = parseTextData(text);
    if (parsed.length > 0) {
        allData = parsed;
        showStatus(`Successfully parsed ${parsed.length} entries.`, 'success');
        generateFilterCheckboxes();
        applyFilters();
    } else {
        showStatus('No valid data found. Please check your format.', 'error');
    }
}

// Handle Excel file upload
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    fileInfo.textContent = `Selected: ${file.name}`;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
            
            // Convert Excel rows to our format
            const textRows = jsonData.map(row => row.join('\t'));
            const text = textRows.join('\n');
            
            const parsed = parseTextData(text);
            if (parsed.length > 0) {
                allData = parsed;
                showStatus(`Successfully parsed ${parsed.length} entries from Excel file.`, 'success');
                generateFilterCheckboxes();
                applyFilters();
            } else {
                showStatus('No valid data found in Excel file.', 'error');
            }
        } catch (error) {
            showStatus(`Error reading Excel file: ${error.message}`, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Parse tab-separated text data
function parseTextData(text) {
    const lines = text.split('\n');
    const parsed = [];
    
    for (let line of lines) {
        line = line.trim();
        if (!line) continue; // Skip empty lines
        
        // Split by tab - this preserves empty strings for consecutive tabs
        const rawColumns = line.split('\t');
        
        // Process columns: trim whitespace but keep empty strings for position tracking
        const columns = rawColumns.map(col => col.trim());
        
        // We expect at least 7 columns (studentName, classAndCode, teacher, essayType, evaluation, evaluator, date)
        // But evaluator can be empty, which may result in consecutive tabs
        // Handle cases where we have exactly 7 columns or more
        if (columns.length < 6) {
            // Not enough columns even after trimming
            continue;
        }
        
        try {
            // Extract data from columns by position (allowing for empty evaluator)
            // Handle cases where evaluator column is empty (creates empty string in array)
            const studentName = (columns[0] || '').trim();
            const classAndCode = (columns[1] || '').trim();
            const teacher = (columns[2] || '').trim();
            const essayType = (columns[3] || '').trim();
            const evaluation = (columns[4] || '').trim();
            // Evaluator might be at index 5, but if it's empty and there are consecutive tabs,
            // we need to check if index 6 exists before index 7
            let evaluator = '';
            let date = '';
            
            // Find evaluator and date - they could be at different positions if evaluator is empty
            // If columns[5] is empty and columns[6] looks like a date, then evaluator is empty
            // If columns[5] is not empty, check if columns[6] is a date
            if (columns.length >= 7) {
                // Standard case: columns[5] is evaluator, columns[6] is date
                evaluator = (columns[5] || '').trim();
                date = (columns[6] || '').trim();
            } else if (columns.length === 6) {
                // Only 6 columns - likely evaluator is missing
                // columns[5] should be the date
                evaluator = '';
                date = (columns[5] || '').trim();
            }
            
            // Skip if essential fields are missing
            if (!studentName || !classAndCode || !date) {
                continue;
            }
            
            // Validate date format (should be YYYY-MM-DD)
            if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
                continue;
            }
            
            // Extract class - handle both formats:
            // Old format: "중1 유마 M^2511" (space-separated, first two parts are class, rest is code)
            // New format: "PurpleG3M_26WI" (single string, entire string is class, no code)
            const classParts = classAndCode.split(/\s+/).filter(part => part.length > 0);
            let classValue, code;
            
            if (classParts.length >= 2) {
                // Old format: space-separated parts
                classValue = classParts.slice(0, 2).join(' ');
                code = classParts.slice(2).join(' ');
            } else if (classParts.length === 1) {
                // New format: single string (e.g., "PurpleG3M_26WI")
                classValue = classParts[0];
                code = '';
            } else {
                // Empty class field - skip this row
                continue;
            }
            
            parsed.push({
                studentName,
                class: classValue,
                code,
                teacher,
                essayType,
                evaluation,
                evaluator,
                date,
                originalIndex: parsed.length // Preserve original order
            });
        } catch (error) {
            console.warn('Error parsing line:', line, error);
            continue;
        }
    }
    
    return parsed;
}

// Generate filter checkboxes from parsed data
function generateFilterCheckboxes() {
    if (allData.length === 0) return;
    
    // Extract unique values
    const classes = [...new Set(allData.map(d => d.class).filter(c => c))].sort();
    const essayTypes = [...new Set(allData.map(d => d.essayType).filter(e => e))].sort();
    const teachers = [...new Set(allData.map(d => d.teacher).filter(t => t))].sort();
    const evaluators = [...new Set(allData.map(d => d.evaluator || ''))].sort();
    const evaluations = [...new Set(allData.map(d => d.evaluation).filter(e => e))].sort();
    
    // Generate Class checkboxes
    classCheckboxes.innerHTML = '';
    classes.forEach(cls => {
        const checkbox = createCheckbox('class', cls, cls);
        classCheckboxes.appendChild(checkbox);
    });
    
    // Generate Essay Type checkboxes
    essayTypeCheckboxes.innerHTML = '';
    essayTypes.forEach(type => {
        const checkbox = createCheckbox('essayType', type, type);
        essayTypeCheckboxes.appendChild(checkbox);
    });
    
    // Generate Teacher checkboxes
    teacherCheckboxes.innerHTML = '';
    teachers.forEach(teacher => {
        const checkbox = createCheckbox('teacher', teacher, teacher);
        teacherCheckboxes.appendChild(checkbox);
    });
    
    // Generate Evaluator checkboxes
    evaluatorCheckboxes.innerHTML = '';
    evaluators.forEach(eval => {
        const label = eval || '[Empty]';
        const checkbox = createCheckbox('evaluator', eval, label);
        evaluatorCheckboxes.appendChild(checkbox);
    });
    
    // Generate Evaluation checkboxes
    evaluationCheckboxes.innerHTML = '';
    evaluations.forEach(eval => {
        const checkbox = createCheckbox('evaluation', eval, eval);
        evaluationCheckboxes.appendChild(checkbox);
    });
    
    // Show filter section
    filterSection.style.display = 'block';
    resultsSection.style.display = 'block';
    
    // Show counts
    showStatus(`${allData.length} entries loaded. Found ${classes.length} classes, ${essayTypes.length} essay types, ${teachers.length} teachers.`, 'success');
}

// Create a checkbox element
function createCheckbox(filterType, value, label) {
    const div = document.createElement('div');
    div.className = 'checkbox-item';
    
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    // Use special identifier for empty values to ensure valid HTML ID
    const safeValue = value === '' ? 'empty' : value;
    checkbox.id = `${filterType}-${safeValue}`.replace(/\s+/g, '-').replace(/[^a-zA-Z0-9-]/g, '');
    checkbox.value = value;
    checkbox.dataset.filterType = filterType;
    checkbox.addEventListener('change', applyFilters);
    
    const labelEl = document.createElement('label');
    labelEl.htmlFor = checkbox.id;
    labelEl.textContent = label;
    
    div.appendChild(checkbox);
    div.appendChild(labelEl);
    
    return div;
}

// Get selected filter values
function getSelectedFilters() {
    const checkboxes = document.querySelectorAll('input[type="checkbox"][data-filter-type]');
    const filters = {
        class: [],
        essayType: [],
        teacher: [],
        evaluator: [],
        evaluation: [],
        dateRange: {
            from: fromDateInput.value || null,
            to: toDateInput.value || null
        }
    };
    
    checkboxes.forEach(cb => {
        if (cb.checked) {
            const filterType = cb.dataset.filterType;
            filters[filterType].push(cb.value);
        }
    });
    
    return filters;
}

// Apply filters to data
function applyFilters() {
    const filters = getSelectedFilters();
    
    filteredData = allData.filter(item => {
        // Check Class filter
        if (filters.class.length > 0 && !filters.class.includes(item.class)) {
            return false;
        }
        
        // Check Essay Type filter
        if (filters.essayType.length > 0 && !filters.essayType.includes(item.essayType)) {
            return false;
        }
        
        // Check Teacher filter
        if (filters.teacher.length > 0 && !filters.teacher.includes(item.teacher)) {
            return false;
        }
        
        // Check Evaluator filter
        if (filters.evaluator.length > 0) {
            const itemEvaluator = item.evaluator || '';
            if (!filters.evaluator.includes(itemEvaluator)) {
                return false;
            }
        }
        
        // Check Evaluation filter
        if (filters.evaluation.length > 0 && !filters.evaluation.includes(item.evaluation)) {
            return false;
        }
        
        // Check Date Range filter
        if (filters.dateRange.from || filters.dateRange.to) {
            const itemDate = item.date;
            if (!itemDate) return false; // Skip items without dates
            
            // Parse dates for comparison (format: YYYY-MM-DD)
            const itemDateObj = new Date(itemDate);
            if (isNaN(itemDateObj.getTime())) return false; // Invalid date
            
            if (filters.dateRange.from) {
                const fromDate = new Date(filters.dateRange.from);
                if (itemDateObj < fromDate) return false;
            }
            
            if (filters.dateRange.to) {
                const toDate = new Date(filters.dateRange.to);
                // Set to end of day for inclusive comparison
                toDate.setHours(23, 59, 59, 999);
                if (itemDateObj > toDate) return false;
            }
        }
        
        return true;
    });
    
    displayResults();
}

// Display filtered results
function displayResults() {
    resultsTableBody.innerHTML = '';
    
    if (filteredData.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = '<td colspan="11" style="text-align: center; padding: 20px;">No results match your filters.</td>';
        resultsTableBody.appendChild(row);
        resultsCount.textContent = 'Showing 0 of ' + allData.length + ' entries';
        return;
    }
    
    // Sort by originalIndex to maintain the original order
    const sortedData = [...filteredData].sort((a, b) => {
        return (a.originalIndex || 0) - (b.originalIndex || 0);
    });
    
    sortedData.forEach((item, index) => {
        const row = document.createElement('tr');
        const originalIdx = item.originalIndex || index;
        const rowId = `row-${originalIdx}`;
        const isSelected = selectedRowIndices.has(originalIdx);
        
        // Get saved checkbox states
        const savedStates = checkboxStates[originalIdx] || { copypaste: false, reviewed: false, export: false };
        
        row.className = isSelected ? 'selected-row' : '';
        row.dataset.originalIndex = originalIdx;
        
        // Make row clickable for selection (but prevent when clicking checkboxes)
        row.addEventListener('click', function(e) {
            // Don't select if clicking directly on a checkbox
            if (e.target.type === 'checkbox') return;
            
            handleRowClick(originalIdx, e);
        });
        
        row.innerHTML = `
            <td>${escapeHtml(item.studentName)}</td>
            <td>${escapeHtml(item.class)}</td>
            <td>${escapeHtml(item.code)}</td>
            <td>${escapeHtml(item.teacher)}</td>
            <td>${escapeHtml(item.essayType)}</td>
            <td>${escapeHtml(item.evaluation)}</td>
            <td>${escapeHtml(item.evaluator || '-')}</td>
            <td>${escapeHtml(item.date)}</td>
            <td style="text-align: center;">
                <input type="checkbox" id="copypaste-${rowId}" class="action-checkbox" data-type="copypaste" data-index="${originalIdx}" ${savedStates.copypaste ? 'checked' : ''}>
            </td>
            <td style="text-align: center;">
                <input type="checkbox" id="reviewed-${rowId}" class="action-checkbox" data-type="reviewed" data-index="${originalIdx}" ${savedStates.reviewed ? 'checked' : ''}>
            </td>
            <td style="text-align: center;">
                <input type="checkbox" id="export-${rowId}" class="action-checkbox" data-type="export" data-index="${originalIdx}" ${savedStates.export ? 'checked' : ''}>
            </td>
        `;
        resultsTableBody.appendChild(row);
        
        // Add event listeners to checkboxes after adding to DOM
        const copypasteCheckbox = document.getElementById(`copypaste-${rowId}`);
        const reviewedCheckbox = document.getElementById(`reviewed-${rowId}`);
        
        copypasteCheckbox.addEventListener('click', function(e) {
            e.stopPropagation(); // Prevent row selection when clicking checkbox
            handleActionCheckboxClick(originalIdx, 'copypaste', this.checked);
        });
        
        reviewedCheckbox.addEventListener('click', function(e) {
            e.stopPropagation(); // Prevent row selection when clicking checkbox
            handleActionCheckboxClick(originalIdx, 'reviewed', this.checked);
        });
        
        const exportCheckbox = document.getElementById(`export-${rowId}`);
        exportCheckbox.addEventListener('click', function(e) {
            e.stopPropagation(); // Prevent row selection when clicking checkbox
            handleActionCheckboxClick(originalIdx, 'export', this.checked);
        });
    });
    
    updateResultsCount();
}

// Toggle row selection
function toggleRowSelection(originalIndex) {
    if (selectedRowIndices.has(originalIndex)) {
        selectedRowIndices.delete(originalIndex);
    } else {
        selectedRowIndices.add(originalIndex);
    }
    updateRowSelectionStyles();
    updateResultsCount();
}

// Update results count with selection info
function updateResultsCount() {
    const selectedCount = Array.from(selectedRowIndices).filter(idx => 
        filteredData.some(item => item.originalIndex === idx)
    ).length;
    
    if (selectedCount > 0) {
        resultsCount.textContent = `Showing ${filteredData.length} of ${allData.length} entries | ${selectedCount} row(s) selected`;
        clearSelectionBtn.style.display = 'inline-block';
    } else {
        resultsCount.textContent = `Showing ${filteredData.length} of ${allData.length} entries`;
        clearSelectionBtn.style.display = 'none';
    }
    
    // Show Select All button when there are results
    if (filteredData.length > 0) {
        selectAllBtn.style.display = 'inline-block';
    } else {
        selectAllBtn.style.display = 'none';
    }
}

// Clear all row selections
function clearSelection() {
    // Get selected indices before clearing
    const selectedIndices = Array.from(selectedRowIndices);
    
    // Clear checkbox states for all selected rows
    selectedIndices.forEach(idx => {
        if (checkboxStates[idx]) {
            checkboxStates[idx] = { copypaste: false, reviewed: false, export: false };
        }
        // Update DOM checkboxes if they exist
        const rowId = `row-${idx}`;
        ['copypaste', 'reviewed', 'export'].forEach(type => {
            const checkbox = document.getElementById(`${type}-${rowId}`);
            if (checkbox) {
                checkbox.checked = false;
            }
        });
    });
    
    selectedRowIndices.clear();
    lastSelectedIndex = null;
    updateRowSelectionStyles();
    updateResultsCount();
}

// Select all visible rows
function selectAllRows() {
    // Get sorted data to maintain order
    const sortedData = [...filteredData].sort((a, b) => {
        return (a.originalIndex || 0) - (b.originalIndex || 0);
    });
    
    sortedData.forEach(item => {
        selectedRowIndices.add(item.originalIndex);
    });
    
    if (sortedData.length > 0) {
        lastSelectedIndex = sortedData[sortedData.length - 1].originalIndex;
    }
    
    updateRowSelectionStyles();
    updateResultsCount();
}

// Transform name from format 1 to format 2
// Format 1: "양시유S(Siyu)	중1 파보 M^2511" (tab-separated)
// Format 2: "양시유S(Siyu)\n중1 파보 M^2511" (single newline between name and class)
function transformNameFormat(item) {
    const studentName = item.studentName;
    const classAndCode = `${item.class} ${item.code}`.trim();
    return `${studentName}\n${classAndCode}`;
}

// Copy export names list to clipboard
function copyExportNames() {
    // Get all checked export checkboxes that are currently in filteredData
    const filteredIndices = new Set(filteredData.map(item => item.originalIndex));
    const checkedIndices = [];
    Object.keys(checkboxStates).forEach(index => {
        const idx = parseInt(index);
        // Only include items that are checked AND currently pass the filters
        if (checkboxStates[index] && checkboxStates[index].export && filteredIndices.has(idx)) {
            checkedIndices.push(idx);
        }
    });
    
    if (checkedIndices.length === 0) {
        showStatus('No names selected for export. Please check the "Export Names List" boxes.', 'error');
        return;
    }
    
    // Get the items in original order (already filtered to only filteredData items)
    const exportItems = checkedIndices
        .map(idx => filteredData.find(item => item.originalIndex === idx))
        .filter(item => item !== undefined)
        .sort((a, b) => (a.originalIndex || 0) - (b.originalIndex || 0));
    
    // Transform each item, append ---, and join with double newline
    const exportText = exportItems
        .map(item => transformNameFormat(item) + '\n\n---')
        .join('\n\n');
    
    // Copy to clipboard
    navigator.clipboard.writeText(exportText).then(() => {
        showStatus(`Copied ${exportItems.length} name(s) to clipboard!`, 'success');
    }).catch(err => {
        showStatus(`Failed to copy: ${err.message}`, 'error');
    });
}

// Update visual styles for selected rows
function updateRowSelectionStyles() {
    const rows = resultsTableBody.querySelectorAll('tr[data-original-index]');
    rows.forEach(row => {
        const originalIdx = parseInt(row.dataset.originalIndex);
        if (selectedRowIndices.has(originalIdx)) {
            row.classList.add('selected-row');
        } else {
            row.classList.remove('selected-row');
        }
    });
}

// Handle action checkbox click (apply to selected rows if any)
function handleActionCheckboxClick(clickedIndex, checkboxType, isChecked) {
    // If there are selected rows and the clicked row is among them, apply to all selected rows
    if (selectedRowIndices.size > 0 && selectedRowIndices.has(clickedIndex)) {
        // Apply to all selected rows
        selectedRowIndices.forEach(originalIdx => {
            if (!checkboxStates[originalIdx]) {
                checkboxStates[originalIdx] = { copypaste: false, reviewed: false, export: false };
            }
            checkboxStates[originalIdx][checkboxType] = isChecked;
            
            // Update the checkbox in the DOM if it exists
            const rowId = `row-${originalIdx}`;
            const checkbox = document.getElementById(`${checkboxType}-${rowId}`);
            if (checkbox) {
                checkbox.checked = isChecked;
            }
        });
    } else {
        // No selected rows or clicked row is not selected - just update this one
        if (!checkboxStates[clickedIndex]) {
            checkboxStates[clickedIndex] = { copypaste: false, reviewed: false, export: false };
        }
        checkboxStates[clickedIndex][checkboxType] = isChecked;
    }
}

// Clear all filters
function clearAllFilters() {
    const checkboxes = document.querySelectorAll('input[type="checkbox"][data-filter-type]');
    checkboxes.forEach(cb => cb.checked = false);
    fromDateInput.value = '';
    toDateInput.value = '';
    applyFilters();
}

// Clear all data
function clearAllData() {
    allData = [];
    filteredData = [];
    selectedRowIndices.clear();
    lastSelectedIndex = null;
    checkboxStates = {};
    pasteArea.value = '';
    fileInput.value = '';
    fileInfo.textContent = '';
    filterSection.style.display = 'none';
    resultsSection.style.display = 'none';
    parseStatus.textContent = '';
    resultsTableBody.innerHTML = '';
    resultsCount.textContent = '';
}

// Show status message
function showStatus(message, type = 'info') {
    parseStatus.textContent = message;
    parseStatus.className = `status-message ${type}`;
    
    if (type === 'success') {
        setTimeout(() => {
            parseStatus.className = 'status-message';
        }, 3000);
    }
}

// Escape HTML to prevent XSS
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

