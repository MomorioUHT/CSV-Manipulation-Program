/**
 * Multiple Filter CSV / XLSX Application
 * Built using OOP and Vanilla JS
 */

class FileManager {
    constructor() {
        this.rawData = []; // Array of arrays (2D array)
        this.headers = [];
        this.rawSheetData = []; // Store the full 2D array including empty rows
        this.fileName = "export";
    }

    async loadFile(file) {
        this.fileName = file.name.split('.').slice(0, -1).join('.');
        const data = await file.arrayBuffer();
        // Use SheetJS to read file
        const workbook = XLSX.read(data);
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Convert sheet to 2D array, preserving blank rows so row indices match exactly
        this.rawSheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", blankrows: true });
        return this.rawSheetData;
    }

    extractDataByHeaderRow(headerRowIndex) {
        if (!this.rawSheetData || this.rawSheetData.length === 0) {
            this.headers = [];
            this.rawData = [];
            return;
        }
        
        // Ensure index is within bounds
        if (headerRowIndex < 0 || headerRowIndex >= this.rawSheetData.length) {
            headerRowIndex = 0;
        }

        this.headers = this.rawSheetData[headerRowIndex] || [];
        // The data is everything below the header
        this.rawData = this.rawSheetData.slice(headerRowIndex + 1);
    }

    exportData(headers, data) {
        // Create a new workbook and add the data
        const wsData = [headers, ...data];
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Filtered Data");

        // Generate download
        XLSX.writeFile(wb, `${this.fileName}_filtered.xlsx`);
    }

    getHeaders() {
        return this.headers;
    }

    getRawData() {
        return this.rawData;
    }
}

class FilterEngine {
    constructor() {
        this.filteredData = [];
    }

    /**
     * Filters the data based on keywords and selected columns.
     */
    filter(rawData, headers, keywords, searchColumns, visibleColumns, rowStart, rowEnd, numericFilters) {
        const rawFilteredData = [];
        const highlightedData = [];

        const searchColsToInclude = searchColumns.length > 0 
            ? searchColumns 
            : headers.map((_, index) => index);

        const visibleColsToInclude = visibleColumns.length > 0 
            ? visibleColumns 
            : headers.map((_, index) => index);
        
        const filteredHeaders = visibleColsToInclude.map(idx => headers[idx]);

        let regexPattern = null;
        if (keywords && keywords.length > 0) {
            const escapedKeywords = keywords.map(kw => kw.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
            regexPattern = new RegExp(`(${escapedKeywords.join('|')})`, 'gi');
        }

        const maxRow = rowEnd ? Math.min(rowEnd, rawData.length) : rawData.length;
        const startRow = Math.max((rowStart || 1) - 1, 0);

        for (let r = startRow; r < maxRow; r++) {
            const row = rawData[r];
            let rowMatches = true;

            if (numericFilters && numericFilters.length > 0) {
                for (const filter of numericFilters) {
                    const cellVal = row[filter.colIndex];
                    if (cellVal === undefined || cellVal === null || cellVal === "") {
                        rowMatches = false;
                        break;
                    }
                    const numCell = Number(cellVal);
                    if (isNaN(numCell)) {
                        rowMatches = false;
                        break;
                    }

                    switch(filter.op) {
                        case '=': if (numCell !== filter.value) rowMatches = false; break;
                        case '>': if (numCell <= filter.value) rowMatches = false; break;
                        case '<': if (numCell >= filter.value) rowMatches = false; break;
                        case '>=': if (numCell < filter.value) rowMatches = false; break;
                        case '<=': if (numCell > filter.value) rowMatches = false; break;
                        case '!=': if (numCell === filter.value) rowMatches = false; break;
                    }
                    if (!rowMatches) break;
                }
            }

            if (!rowMatches) continue;

            if (keywords && keywords.length > 0) {
                rowMatches = false; // Start false, require at least one keyword match (OR logic)
                for (const colIndex of searchColsToInclude) {
                    const cellValue = String(row[colIndex] || "");
                    // regexPattern is already an OR combination of all keywords: (kw1|kw2|...)
                    if (cellValue.match(regexPattern)) {
                        rowMatches = true;
                        break;
                    }
                }
            }

            if (rowMatches) {
                const newHighlightedRow = [];
                for (const colIndex of visibleColsToInclude) {
                    const cellValue = row[colIndex];
                    if (cellValue === undefined || cellValue === null || cellValue === "") {
                        newHighlightedRow.push("");
                    } else {
                        const strVal = String(cellValue);
                        if (regexPattern && regexPattern.test(strVal)) {
                            newHighlightedRow.push(strVal.replace(regexPattern, '<span class="highlight">$1</span>'));
                        } else {
                            newHighlightedRow.push(strVal);
                        }
                    }
                }
                rawFilteredData.push([...row]);
                highlightedData.push(newHighlightedRow);
            }
        }

        return { highlightedData, rawFilteredData, filteredHeaders };
    }
}

class UIManager {
    constructor(appContext) {
        this.app = appContext;

        // DOM Elements
        this.fileInput = document.getElementById('file-upload');
        this.fileNameDisplay = document.getElementById('file-name');
        this.headerRowInput = document.getElementById('header-row');
        this.keywordsInput = document.getElementById('filter-keywords');
        this.rowStartInput = document.getElementById('row-start');
        this.rowEndInput = document.getElementById('row-end');
        this.numericFiltersSection = document.getElementById('numeric-filters-section');
        this.numericFiltersContainer = document.getElementById('numeric-filters-container');
        this.btnAddNumericFilter = document.getElementById('btn-add-numeric-filter');
        this.checkboxContainerSearch = document.getElementById('column-checkboxes-search');
        this.checkboxContainerVisible = document.getElementById('column-checkboxes-visible');
        this.btnFilter = document.getElementById('btn-filter');
        this.btnDownload = document.getElementById('btn-download');
        this.btnClear = document.getElementById('btn-clear');
        
        this.paginationControls = document.getElementById('pagination-controls');
        this.btnPrev = document.getElementById('btn-prev');
        this.btnNext = document.getElementById('btn-next');
        this.pageInfo = document.getElementById('page-info');

        this.tableWrapper = document.getElementById('table-wrapper');
        this.spinnerContainer = document.getElementById('spinner-container');
        this.spinnerText = this.spinnerContainer.querySelector('.spinner-text');
        this.tableHead = document.getElementById('table-head');
        this.tableBody = document.getElementById('table-body');
        this.tableInfo = document.getElementById('table-info');

        this.bindEvents();
    }

    bindEvents() {
        this.fileInput.addEventListener('change', (e) => this.app.handleFileUpload(e));
        this.headerRowInput.addEventListener('change', () => this.app.handleHeaderRowChange());
        this.btnFilter.addEventListener('click', () => this.app.handleFilter());
        this.btnDownload.addEventListener('click', () => this.app.handleDownload());
        this.btnClear.addEventListener('click', () => this.app.handleClear());
        this.btnPrev.addEventListener('click', () => this.app.prevPage());
        this.btnNext.addEventListener('click', () => this.app.nextPage());
        if(this.btnAddNumericFilter) this.btnAddNumericFilter.addEventListener('click', () => this.addNumericFilterRow());

        // Optional: Trigger filter on enter key
        this.keywordsInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                this.app.handleFilter();
            }
        });
    }

    updateFileName(name) {
        this.fileNameDisplay.textContent = name;
    }

    renderCheckboxGroup(container, typeStr, headers) {
        container.innerHTML = ''; 

        if (headers.length === 0) {
            container.innerHTML = '<span class="placeholder-text">ไม่พบคอลัมน์ในไฟล์</span>';
            return;
        }

        const allLabel = document.createElement('label');
        allLabel.className = 'checkbox-item';
        allLabel.innerHTML = `
            <input type="checkbox" id="chk-all-${typeStr}" checked>
            <span>เลือกทั้งหมด (All)</span>
        `;
        container.appendChild(allLabel);

        const chkAll = allLabel.querySelector(`#chk-all-${typeStr}`);
        const colCheckboxes = [];

        headers.forEach((header, index) => {
            const label = document.createElement('label');
            label.className = 'checkbox-item disabled';
            label.innerHTML = `
                <input type="checkbox" value="${index}" class="col-chk-${typeStr}" disabled>
                <span>${header || `Column ${index + 1}`}</span>
            `;
            container.appendChild(label);
            colCheckboxes.push(label.querySelector('input'));
        });

        chkAll.addEventListener('change', (e) => {
            const isChecked = e.target.checked;
            colCheckboxes.forEach(chk => {
                chk.disabled = isChecked;
                chk.checked = false;
                chk.parentElement.classList.toggle('disabled', isChecked);
            });
        });
    }

    renderCheckboxes(headers) {
        if (this.checkboxContainerSearch) this.renderCheckboxGroup(this.checkboxContainerSearch, 'search', headers);
        if (this.checkboxContainerVisible) this.renderCheckboxGroup(this.checkboxContainerVisible, 'visible', headers);
    }

    getSearchColumns() {
        const chkAll = document.getElementById('chk-all-search');
        if (chkAll && chkAll.checked) return [];
        const selected = [];
        document.querySelectorAll('.col-chk-search:checked').forEach(chk => selected.push(parseInt(chk.value, 10)));
        return selected;
    }

    getVisibleColumns() {
        const chkAll = document.getElementById('chk-all-visible');
        if (chkAll && chkAll.checked) return [];
        const selected = [];
        document.querySelectorAll('.col-chk-visible:checked').forEach(chk => selected.push(parseInt(chk.value, 10)));
        return selected;
    }

    setAvailableNumericColumns(headers, numericColIndices) {
        this.availableNumericColumns = numericColIndices.map(index => ({
            index: index,
            name: headers[index] || `Column ${index + 1}`
        }));

        this.numericFiltersContainer.innerHTML = ''; 
        
        if (this.availableNumericColumns.length === 0) {
            this.numericFiltersSection.style.display = 'none';
        } else {
            this.numericFiltersSection.style.display = 'flex';
        }
    }

    addNumericFilterRow() {
        if (!this.availableNumericColumns || this.availableNumericColumns.length === 0) return;

        const row = document.createElement('div');
        row.className = 'numeric-filter-row';

        const colSelect = document.createElement('select');
        colSelect.className = 'num-col';
        this.availableNumericColumns.forEach(col => {
            const option = document.createElement('option');
            option.value = col.index;
            option.textContent = col.name;
            colSelect.appendChild(option);
        });

        row.innerHTML = `
            <select class="num-op">
                <option value="=">=</option>
                <option value=">">&gt;</option>
                <option value="<">&lt;</option>
                <option value=">=">&gt;=</option>
                <option value="<=">&lt;=</option>
                <option value="!=">!=</option>
            </select>
            <input type="number" class="num-val" placeholder="ระบุตัวเลข">
            <button type="button" class="remove-num-filter">X</button>
        `;
        
        row.insertBefore(colSelect, row.firstChild);

        row.querySelector('.remove-num-filter').addEventListener('click', () => {
            row.remove();
        });

        this.numericFiltersContainer.appendChild(row);
    }

    getNumericFilters() {
        const filters = [];
        const rows = this.numericFiltersContainer.querySelectorAll('.numeric-filter-row');
        rows.forEach(row => {
            const colIndex = parseInt(row.querySelector('.num-col').value, 10);
            const op = row.querySelector('.num-op').value;
            const valStr = row.querySelector('.num-val').value;
            
            if (!isNaN(colIndex) && op && valStr !== '') {
                filters.push({
                    colIndex: colIndex,
                    op: op,
                    value: parseFloat(valStr)
                });
            }
        });
        return filters;
    }

    getKeywords() {
        const val = this.keywordsInput.value.trim();
        if (!val) return [];
        return val.split(',').map(k => k.trim()).filter(k => k.length > 0);
    }

    renderTable(headers, data, isHtmlData = false) {
        // Render headers
        this.tableHead.innerHTML = '';
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header || 'Unknown';
            this.tableHead.appendChild(th);
        });

        // Render body
        this.tableBody.innerHTML = '';
        data.forEach(row => {
            const tr = document.createElement('tr');
            headers.forEach((_, index) => {
                const td = document.createElement('td');
                let val = row[index];
                
                if (val === undefined || val === null || val === "") {
                    td.innerHTML = '<span style="color: red;">null</span>';
                } else if (isHtmlData) {
                    td.innerHTML = val;
                } else {
                    td.textContent = val;
                }
                tr.appendChild(td);
            });
            this.tableBody.appendChild(tr);
        });
    }

    setTableInfo(start, end, total) {
        if (total === 0) {
            this.tableInfo.textContent = "ไม่พบรายการข้อมูล";
            return;
        }
        this.tableInfo.textContent = `แสดงรายการที่ ${start}-${end} จาก ${total} รายการ`;
    }

    updatePagination(page, maxPage) {
        if (maxPage <= 1) {
            this.paginationControls.style.display = 'none';
        } else {
            this.paginationControls.style.display = 'flex';
            this.btnPrev.disabled = page <= 1;
            this.btnNext.disabled = page >= maxPage;
            this.pageInfo.textContent = `หน้า ${page} / ${maxPage}`;
        }
    }

    setDownloadState(enabled) {
        this.btnDownload.disabled = !enabled;
    }

    showError(msg) {
        alert(msg);
    }

    showSpinner(text) {
        if (this.tableWrapper) this.tableWrapper.style.display = 'none';
        if (this.spinnerContainer) {
            this.spinnerContainer.style.display = 'flex';
            if (this.spinnerText) this.spinnerText.textContent = text || "กำลังประมวลผล...";
        }
    }
    
    hideSpinner() {
        if (this.spinnerContainer) this.spinnerContainer.style.display = 'none';
        if (this.tableWrapper) this.tableWrapper.style.display = 'block';
    }
}

class App {
    constructor() {
        this.fileManager = new FileManager();
        this.filterEngine = new FilterEngine();
        this.uiManager = new UIManager(this);

        this.currentRawFilteredData = [];
        this.currentHighlightedData = [];
        this.currentFilteredHeaders = [];

        this.currentPage = 1;
        this.itemsPerPage = 50;
        this.isHtmlView = false;
    }

    async handleFileUpload(event) {
        const file = event.target.files[0];
        if (!file) return;

        this.uiManager.updateFileName(file.name);
        this.uiManager.tableInfo.textContent = "";
        this.uiManager.showSpinner("กำลังโหลดและอ่านไฟล์...");

        setTimeout(async () => {
            try {
                const rawSheetData = await this.fileManager.loadFile(file);
                
                // Auto detect header row: Find the row in the first 20 rows that has the most non-empty cells
                let detectedIndex = 0;
                let maxCount = -1;
                for (let i = 0; i < Math.min(20, rawSheetData.length); i++) {
                    const count = rawSheetData[i] ? rawSheetData[i].filter(c => c !== "").length : 0;
                    if (count > maxCount) {
                        maxCount = count;
                        detectedIndex = i;
                    }
                }

                this.uiManager.headerRowInput.value = detectedIndex + 1; // Convert to 1-based index for user
                this.updateDataView(detectedIndex);
                this.uiManager.hideSpinner();
            } catch (error) {
                console.error(error);
                this.uiManager.hideSpinner();
                this.uiManager.showError("เกิดข้อผิดพลาดในการอ่านไฟล์ โปรดตรวจสอบว่าไฟล์ถูกต้อง");
            }
        }, 50);
    }

    handleHeaderRowChange() {
        const val = parseInt(this.uiManager.headerRowInput.value, 10);
        if (isNaN(val) || val < 1) return;
        
        const headerRowIndex = val - 1; // Convert back to 0-based index
        this.updateDataView(headerRowIndex);
    }

    updateDataView(headerRowIndex) {
        this.fileManager.extractDataByHeaderRow(headerRowIndex);
        const headers = this.fileManager.getHeaders();
        const rawData = this.fileManager.getRawData();

        this.uiManager.renderCheckboxes(headers);
        
        // Auto-detect numeric columns
        const numericColIndices = [];
        if (rawData.length > 0) {
            const rowsToCheck = Math.min(100, rawData.length);
            for (let c = 0; c < headers.length; c++) {
                let hasNumber = false;
                let isAllNumeric = true;
                
                for (let r = 0; r < rowsToCheck; r++) {
                    if (!rawData[r]) continue;
                    const val = rawData[r][c];
                    if (val === undefined || val === null || val === "") continue;
                    
                    const numVal = Number(val);
                    if (isNaN(numVal)) {
                        isAllNumeric = false;
                        break;
                    } else {
                        hasNumber = true;
                    }
                }
                
                if (hasNumber && isAllNumeric) {
                    numericColIndices.push(c);
                }
            }
        }
        this.uiManager.setAvailableNumericColumns(headers, numericColIndices);

        this.currentRawFilteredData = rawData.map(r => [...r]);
        this.currentHighlightedData = rawData.map(r => [...r]);
        this.currentFilteredHeaders = headers;
        this.currentPage = 1;
        this.isHtmlView = false;
        
        this.renderCurrentPage();
        this.uiManager.setDownloadState(false);
    }

    renderCurrentPage() {
        const total = this.currentHighlightedData.length;
        const maxPage = Math.ceil(total / this.itemsPerPage) || 1;
        if (this.currentPage < 1) this.currentPage = 1;
        if (this.currentPage > maxPage) this.currentPage = maxPage;

        const startIdx = (this.currentPage - 1) * this.itemsPerPage;
        const endIdx = Math.min(startIdx + this.itemsPerPage, total);

        const pageData = this.currentHighlightedData.slice(startIdx, endIdx);

        this.uiManager.renderTable(this.currentFilteredHeaders, pageData, this.isHtmlView);

        const displayStart = total === 0 ? 0 : startIdx + 1;
        this.uiManager.setTableInfo(displayStart, endIdx, total);
        this.uiManager.updatePagination(this.currentPage, maxPage);
    }
    
    prevPage() {
        if (this.currentPage > 1) {
            this.currentPage--;
            this.renderCurrentPage();
        }
    }
    
    nextPage() {
        const maxPage = Math.ceil(this.currentHighlightedData.length / this.itemsPerPage) || 1;
        if (this.currentPage < maxPage) {
            this.currentPage++;
            this.renderCurrentPage();
        }
    }

    handleFilter() {
        const headers = this.fileManager.getHeaders();
        const rawData = this.fileManager.getRawData();

        if (headers.length === 0 || rawData.length === 0) {
            this.uiManager.showError("กรุณาอัปโหลดไฟล์ก่อนทำการ Filter");
            return;
        }

        this.uiManager.showSpinner("กำลังกรองข้อมูล...");

        setTimeout(() => {
            const keywords = this.uiManager.getKeywords();
            const searchColumns = this.uiManager.getSearchColumns();
            const visibleColumns = this.uiManager.getVisibleColumns();
            const numericFilters = this.uiManager.getNumericFilters();
            const rowStart = parseInt(this.uiManager.rowStartInput.value, 10) || null;
            const rowEnd = parseInt(this.uiManager.rowEndInput.value, 10) || null;

            const { highlightedData, rawFilteredData, filteredHeaders } = this.filterEngine.filter(
                rawData,
                headers,
                keywords,
                searchColumns,
                visibleColumns,
                rowStart,
                rowEnd,
                numericFilters
            );

            if (rawFilteredData.length === 0) {
                alert("ไม่พบข้อมูลที่ตรงกับเงื่อนไขการค้นหา");
                
                // Fallback table to initial data, but keep UI inputs intact
                this.currentRawFilteredData = rawData.map(r => [...r]);
                this.currentHighlightedData = rawData.map(r => [...r]);
                this.currentFilteredHeaders = headers;
                
                this.currentPage = 1;
                this.isHtmlView = false;
                this.renderCurrentPage();
                
                this.uiManager.setDownloadState(false);
                this.uiManager.hideSpinner();
                return;
            }

            this.currentRawFilteredData = rawFilteredData;
            this.currentHighlightedData = highlightedData;
            this.currentFilteredHeaders = filteredHeaders;
            
            this.currentPage = 1;
            this.isHtmlView = true;
            this.renderCurrentPage();
            
            this.uiManager.setDownloadState(true);
            this.uiManager.hideSpinner();
        }, 50);
    }

    handleClear() {
        this.uiManager.keywordsInput.value = '';
        this.uiManager.rowStartInput.value = '';
        this.uiManager.rowEndInput.value = '';
        
        // Remove all dynamically added numeric filter rows
        this.uiManager.numericFiltersContainer.innerHTML = '';
        
        const chkAllSearch = document.getElementById('chk-all-search');
        if (chkAllSearch) {
            chkAllSearch.checked = true;
            chkAllSearch.dispatchEvent(new Event('change'));
        }

        const chkAllVisible = document.getElementById('chk-all-visible');
        if (chkAllVisible) {
            chkAllVisible.checked = true;
            chkAllVisible.dispatchEvent(new Event('change'));
        }

        const headerRowIndex = parseInt(this.uiManager.headerRowInput.value, 10) - 1 || 0;
        this.updateDataView(headerRowIndex);
    }

    handleDownload() {
        if (this.currentRawFilteredData.length === 0) {
            this.uiManager.showError("ไม่มีข้อมูลสำหรับดาวน์โหลด");
            return;
        }

        const fullHeaders = this.fileManager.getHeaders();
        this.fileManager.exportData(fullHeaders, this.currentRawFilteredData);
        alert("ดาวน์โหลดไฟล์สำเร็จเรียบร้อยแล้ว!");
    }
}

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
    new App();
});
