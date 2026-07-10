/**
 * Multiple Filter CSV / XLSX Application
 * Built using OOP and Vanilla JS
 */

class FileManager {
    constructor() {
        this.rawData = []; // Array of arrays (2D array)
        this.headers = [];
        this.headerObjs = []; // Stores top/bottom parts for double headers
        this.rawSheetData = []; // Store the full 2D array including empty rows
        this.fileName = "export";
    }

    clear() {
        this.rawData = [];
        this.headers = [];
        this.headerObjs = [];
        this.rawSheetData = [];
        this.fileName = "export";
    }

    async loadFile(file) {
        this.fileName = file.name.split('.').slice(0, -1).join('.');
        const data = await file.arrayBuffer();
        // Use SheetJS to read file
        const workbook = XLSX.read(data);
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        this.merges = worksheet['!merges'] || [];
        // Convert sheet to 2D array, preserving blank rows so row indices match exactly
        this.rawSheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", blankrows: true });
        return this.rawSheetData;
    }

    getCellValueStr(val) {
        if (val === undefined || val === null) return "";
        if (typeof val === 'object') {
            if (val instanceof Date) return val.toISOString();
            if (val.v !== undefined) return String(val.v).trim();
            if (val.w !== undefined) return String(val.w).trim();
            if (val.t !== undefined) return String(val.t).trim();
            try {
                return JSON.stringify(val);
            } catch (e) {
                return "[Object]";
            }
        }
        return String(val).trim();
    }

    isCellInMerge(r, c) {
        if (!this.merges) return null;
        for (let i = 0; i < this.merges.length; i++) {
            const m = this.merges[i];
            if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
                return m;
            }
        }
        return null;
    }

    extractDataByHeaderRow(headerRowIndex, isDoubleHeader = false) {
        if (!this.rawSheetData || this.rawSheetData.length === 0) {
            this.headers = [];
            this.headerObjs = [];
            this.rawData = [];
            return;
        }

        // Ensure index is within bounds
        if (headerRowIndex < 0 || headerRowIndex >= this.rawSheetData.length) {
            headerRowIndex = 0;
        }

        if (isDoubleHeader) {
            const row1 = this.rawSheetData[headerRowIndex] || [];
            const row2 = this.rawSheetData[headerRowIndex + 1] || [];
            this.headers = [];
            this.headerObjs = [];

            const maxLen = Math.max(row1.length, row2.length);

            for (let i = 0; i < maxLen; i++) {
                let topVal = row1[i];
                const merge = this.isCellInMerge(headerRowIndex, i);
                if (merge) {
                    topVal = this.rawSheetData[merge.s.r] ? this.rawSheetData[merge.s.r][merge.s.c] : "";
                }

                const topStr = this.getCellValueStr(topVal);
                const bottomStr = this.getCellValueStr(row2[i]);

                let combined = topStr ? `${topStr}_${bottomStr}` : bottomStr;
                if (!topStr && !bottomStr) combined = `Column ${i + 1}`;
                if (topStr && !bottomStr) combined = topStr;
                if (topStr === bottomStr) combined = topStr;

                this.headers.push(combined);
                this.headerObjs.push({ top: topStr, bottom: bottomStr, combined });
            }
            this.rawData = this.rawSheetData.slice(headerRowIndex + 2);
        } else {
            this.headers = (this.rawSheetData[headerRowIndex] || []).map(c => this.getCellValueStr(c));
            this.headerObjs = this.headers.map((h, i) => ({ top: null, bottom: h, combined: h }));
            this.rawData = this.rawSheetData.slice(headerRowIndex + 1);
        }
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

    getHeaderObjs() {
        return this.headerObjs || [];
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
    filter(rawData, headers, headerObjs, keywords, searchColumns, visibleColumns, rowStart, rowEnd, numericFilters) {
        const rawFilteredData = [];
        const highlightedData = [];

        const searchColsToInclude = searchColumns.length > 0
            ? searchColumns
            : headers.map((_, index) => index);

        const visibleColsToInclude = visibleColumns.length > 0
            ? visibleColumns
            : headers.map((_, index) => index);

        const filteredHeaders = visibleColsToInclude.map(idx => headers[idx]);
        const filteredHeaderObjs = headerObjs ? visibleColsToInclude.map(idx => headerObjs[idx]) : null;

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

                    switch (filter.op) {
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

        return { highlightedData, rawFilteredData, filteredHeaders, filteredHeaderObjs };
    }
}

class UIManager {
    constructor(appContext) {
        this.app = appContext;

        // DOM Elements
        this.fileInput = document.getElementById('file-upload');
        this.fileNameDisplay = document.getElementById('file-name');
        this.btnRemoveFile = document.getElementById('btn-remove-file');
        this.headerRowInput = document.getElementById('header-row');
        this.doubleHeaderCheckbox = document.getElementById('double-header');
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
        this.tableInfo = document.getElementById('table-info');

        this.topScrollbar = document.getElementById('top-scrollbar');
        this.topScrollbarContent = document.getElementById('top-scrollbar-content');



        this.resultSummaryCard = document.getElementById('result-summary-card');
        this.resultCount = document.getElementById('result-count');

        this.tableWrapper = document.getElementById('table-wrapper');
        this.tableOverlay = document.getElementById('table-overlay');

        if (this.topScrollbar && this.tableWrapper) {
            let isSyncingLeftScroll = false;
            let isSyncingRightScroll = false;

            this.topScrollbar.addEventListener('scroll', () => {
                if (!isSyncingLeftScroll) {
                    isSyncingRightScroll = true;
                    this.tableWrapper.scrollLeft = this.topScrollbar.scrollLeft;
                }
                isSyncingLeftScroll = false;
            });
            this.tableWrapper.addEventListener('scroll', () => {
                if (!isSyncingRightScroll) {
                    isSyncingLeftScroll = true;
                    this.topScrollbar.scrollLeft = this.tableWrapper.scrollLeft;
                }
                isSyncingRightScroll = false;
            });
        }
        this.tableOverlayText = document.getElementById('table-overlay-text');
        this.tableHead = document.getElementById('table-head');
        this.tableBody = document.getElementById('table-body');
        this.tableInfo = document.getElementById('table-info');

        this.bindEvents();
    }

    bindEvents() {
        this.fileInput.addEventListener('change', (e) => this.app.handleFileUpload(e));
        this.headerRowInput.addEventListener('change', () => this.app.handleHeaderRowChange());
        if (this.doubleHeaderCheckbox) this.doubleHeaderCheckbox.addEventListener('change', () => this.app.handleHeaderRowChange());
        this.btnFilter.addEventListener('click', () => this.app.handleFilter());
        this.btnDownload.addEventListener('click', () => this.app.handleDownload());
        this.btnClear.addEventListener('click', () => this.app.handleClear());
        this.btnPrev.addEventListener('click', () => this.app.prevPage());
        this.btnNext.addEventListener('click', () => this.app.nextPage());
        if (this.btnAddNumericFilter) this.btnAddNumericFilter.addEventListener('click', () => this.addNumericFilterRow());
        if (this.btnRemoveFile) this.btnRemoveFile.addEventListener('click', () => this.app.handleRemoveFile());

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

    insertLineBreaks(str, maxLen) {
        if (!str || typeof str !== 'string') return str;

        // Replace actual newline characters with a space so they don't break prematurely.
        // This allows our custom maxLen logic to dictate the line breaks.
        str = str.replace(/\n/g, ' ');

        // If string is purely short text (even without tags), return early
        if (str.length <= maxLen && !str.includes('<')) return str;

        let result = '';
        let count = 0;
        let inTag = false;

        for (let i = 0; i < str.length; i++) {
            const char = str[i];

            if (char === '<') inTag = true;

            result += char;

            if (char === '>') {
                inTag = false;
                continue;
            }

            if (!inTag) {
                count++;
                if (count >= maxLen) {
                    result += '<br>';
                    count = 0;
                }
            }
        }
        return result;
    }

    renderTable(headerObjs, data, rawData, fullHeaders, isHtmlData = false, isDoubleHeader = false) {
        // Render headers
        this.tableHead.innerHTML = '';
        if (isDoubleHeader && headerObjs) {
            const trTop = document.createElement('tr');
            const trBottom = document.createElement('tr');

            let currentTop = null;
            let currentTopTh = null;
            let currentTopColSpan = 0;

            headerObjs.forEach((hObj) => {
                const thBottom = document.createElement('th');
                thBottom.textContent = hObj.bottom || '';
                trBottom.appendChild(thBottom);

                if (hObj.top === currentTop && currentTop !== null) {
                    currentTopColSpan++;
                    currentTopTh.colSpan = currentTopColSpan;
                } else {
                    currentTop = hObj.top;
                    currentTopColSpan = 1;
                    currentTopTh = document.createElement('th');
                    currentTopTh.textContent = currentTop || '';
                    currentTopTh.colSpan = currentTopColSpan;
                    trTop.appendChild(currentTopTh);
                }
            });
            this.tableHead.appendChild(trTop);
            this.tableHead.appendChild(trBottom);
        } else {
            const tr = document.createElement('tr');
            headerObjs.forEach(hObj => {
                const th = document.createElement('th');
                th.textContent = (hObj && hObj.combined) ? hObj.combined : (hObj || 'Unknown');
                tr.appendChild(th);
            });
            this.tableHead.appendChild(tr);
        }

        // Find Link column index in fullHeaders
        let linkColIndex = -1;
        if (fullHeaders) {
            linkColIndex = fullHeaders.findIndex(h => h && h.toLowerCase() === 'link');
        }

        // Render body
        this.tableBody.innerHTML = '';
        data.forEach((row, rowIndex) => {
            const tr = document.createElement('tr');

            // Check for link
            if (linkColIndex !== -1 && rawData && rawData[rowIndex]) {
                const linkUrl = rawData[rowIndex][linkColIndex];
                if (linkUrl && typeof linkUrl === 'string' && (linkUrl.startsWith('http') || linkUrl.startsWith('www'))) {
                    tr.style.cursor = 'pointer';
                    tr.title = `ไปที่ลิงก์: ${linkUrl}`;
                    tr.addEventListener('click', (e) => {
                        window.open(linkUrl.startsWith('http') ? linkUrl : `http://${linkUrl}`, '_blank');
                    });
                    tr.addEventListener('mouseenter', () => tr.style.backgroundColor = 'rgba(255, 107, 0, 0.05)');
                    tr.addEventListener('mouseleave', () => tr.style.backgroundColor = '');
                }
            }

            headerObjs.forEach((_, index) => {
                const td = document.createElement('td');
                let val = row[index];

                if (val === undefined || val === null || val === "") {
                    td.innerHTML = '<span style="color: red;">null</span>';
                } else if (isHtmlData) {
                    td.innerHTML = this.insertLineBreaks(val, 100);
                } else {
                    td.innerHTML = this.insertLineBreaks(String(val), 100);
                }
                tr.appendChild(td);
            });
            this.tableBody.appendChild(tr);
        });

        // Sync top scrollbar width
        if (this.topScrollbarContent && this.tableHead) {
            setTimeout(() => {
                const table = document.getElementById('data-table');
                if (table) {
                    this.topScrollbarContent.style.width = table.offsetWidth + 'px';
                }
            }, 50);
        }
    }

    setTableInfo(start, end, total) {
        if (total === 0) {
            this.tableInfo.textContent = "ไม่พบรายการข้อมูล";
        } else {
            this.tableInfo.textContent = `แสดงรายการที่ ${start}-${end} จาก ${total} รายการ`;
        }
    }

    showResultCard(total) {
        if (this.resultSummaryCard && this.resultCount && total > 0) {
            this.resultCount.textContent = total.toLocaleString();
            this.resultSummaryCard.style.display = 'flex';
        }
    }

    hideResultCard() {
        if (this.resultSummaryCard) {
            this.resultSummaryCard.style.display = 'none';
        }
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
        if (this.tableOverlay) {
            this.tableOverlay.style.display = 'flex';
            if (this.tableOverlayText) this.tableOverlayText.textContent = text || "กำลังประมวลผล...";
        }
    }

    hideSpinner() {
        if (this.tableOverlay) this.tableOverlay.style.display = 'none';
    }

    scrollToCard() {
        if (this.resultSummaryCard && this.resultSummaryCard.style.display !== 'none') {
            this.resultSummaryCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
    }

    scrollToTable() {
        if (this.tableWrapper) {
            this.tableWrapper.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
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
        this.currentFilteredHeaderObjs = [];

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

                // Auto detect header row
                let detectedIndex = 0;
                let maxCount = -1;
                for (let i = 0; i < Math.min(20, rawSheetData.length); i++) {
                    const count = rawSheetData[i] ? rawSheetData[i].filter(c => c !== undefined && c !== null && String(c).trim() !== "").length : 0;
                    if (count > maxCount) {
                        maxCount = count;
                        detectedIndex = i;
                    }
                }

                let isDoubleHeader = false;
                if (detectedIndex > 0) {
                    const prevRow = rawSheetData[detectedIndex - 1] || [];
                    const hasText = prevRow.some(c => c !== undefined && c !== null && String(c).trim() !== "");
                    if (hasText) {
                        isDoubleHeader = true;
                        detectedIndex = detectedIndex - 1; // set header row to the top row of the double header
                    }
                }

                this.uiManager.headerRowInput.value = detectedIndex + 1; // Convert to 1-based index for user
                if (this.uiManager.doubleHeaderCheckbox) {
                    this.uiManager.doubleHeaderCheckbox.checked = isDoubleHeader;
                }

                this.updateDataView(detectedIndex, isDoubleHeader);
                if (this.uiManager.btnRemoveFile) {
                    this.uiManager.btnRemoveFile.style.display = 'block';
                }
                this.uiManager.hideSpinner();
                this.uiManager.scrollToTable();
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
        const isDoubleHeader = this.uiManager.doubleHeaderCheckbox ? this.uiManager.doubleHeaderCheckbox.checked : false;
        this.updateDataView(headerRowIndex, isDoubleHeader);
    }

    updateDataView(headerRowIndex, isDoubleHeader = false) {
        this.fileManager.extractDataByHeaderRow(headerRowIndex, isDoubleHeader);
        const headers = this.fileManager.getHeaders();
        const headerObjs = this.fileManager.getHeaderObjs();
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
        this.currentFilteredHeaderObjs = headerObjs;
        this.currentPage = 1;
        this.isHtmlView = false;

        this.uiManager.hideResultCard();
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
        const rawPageData = this.currentRawFilteredData.slice(startIdx, endIdx);
        const isDoubleHeader = this.uiManager.doubleHeaderCheckbox ? this.uiManager.doubleHeaderCheckbox.checked : false;

        this.uiManager.renderTable(
            this.currentFilteredHeaderObjs || this.currentFilteredHeaders,
            pageData,
            rawPageData,
            this.fileManager.getHeaders(),
            this.isHtmlView,
            isDoubleHeader
        );

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
        const headerObjs = this.fileManager.getHeaderObjs();
        const rawData = this.fileManager.getRawData();

        if (headers.length === 0 || rawData.length === 0) {
            this.uiManager.showError("กรุณาอัปโหลดไฟล์ก่อนทำการ Filter");
            return;
        }

        this.uiManager.hideResultCard();
        this.uiManager.showSpinner("กำลังกรองข้อมูล...");

        setTimeout(() => {
            const keywords = this.uiManager.getKeywords();
            const searchColumns = this.uiManager.getSearchColumns();
            const visibleColumns = this.uiManager.getVisibleColumns();
            const numericFilters = this.uiManager.getNumericFilters();
            const rowStart = parseInt(this.uiManager.rowStartInput.value, 10) || null;
            const rowEnd = parseInt(this.uiManager.rowEndInput.value, 10) || null;

            const { highlightedData, rawFilteredData, filteredHeaders, filteredHeaderObjs } = this.filterEngine.filter(
                rawData,
                headers,
                headerObjs,
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
                this.currentFilteredHeaderObjs = headerObjs;

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
            this.currentFilteredHeaderObjs = filteredHeaderObjs;

            this.currentPage = 1;
            this.isHtmlView = true;
            this.renderCurrentPage();

            this.uiManager.setDownloadState(true);
            this.uiManager.hideSpinner();

            this.uiManager.showResultCard(rawFilteredData.length);
            this.uiManager.scrollToCard();
        }, 50);
    }

    handleClear() {
        this.uiManager.keywordsInput.value = '';
        this.uiManager.rowStartInput.value = '';
        this.uiManager.rowEndInput.value = '';
        this.uiManager.numericFiltersContainer.innerHTML = '';

        const headers = this.fileManager.getHeaders();
        const headerObjs = this.fileManager.getHeaderObjs();
        const rawData = this.fileManager.getRawData();

        if (headers.length === 0 || rawData.length === 0) return;

        this.uiManager.renderCheckboxes(headers);

        this.currentRawFilteredData = rawData.map(r => [...r]);
        this.currentHighlightedData = rawData.map(r => [...r]);
        this.currentFilteredHeaders = headers;
        this.currentFilteredHeaderObjs = headerObjs;
        this.currentPage = 1;
        this.isHtmlView = false;

        this.uiManager.hideResultCard();
        this.renderCurrentPage();
        this.uiManager.setDownloadState(false);
    }

    handleRemoveFile() {
        // Clear file input
        this.uiManager.fileInput.value = '';
        this.uiManager.fileNameDisplay.textContent = 'เลือกไฟล์ CSV หรือ XLSX...';
        if (this.uiManager.btnRemoveFile) {
            this.uiManager.btnRemoveFile.style.display = 'none';
        }

        // Clear filter inputs
        this.uiManager.keywordsInput.value = '';
        this.uiManager.rowStartInput.value = '';
        this.uiManager.rowEndInput.value = '';
        this.uiManager.numericFiltersContainer.innerHTML = '';
        this.uiManager.checkboxContainerSearch.innerHTML = '';
        this.uiManager.checkboxContainerVisible.innerHTML = '';

        // Clear data state
        this.fileManager.clear();
        this.currentRawFilteredData = [];
        this.currentHighlightedData = [];
        this.currentFilteredHeaders = [];
        this.currentFilteredHeaderObjs = [];
        this.currentPage = 1;
        this.isHtmlView = false;

        // Clear UI
        this.uiManager.hideResultCard();
        this.uiManager.tableHead.innerHTML = '';
        this.uiManager.tableBody.innerHTML = '';
        this.uiManager.updatePagination(1, 1);
        this.uiManager.tableInfo.textContent = '';
        this.uiManager.setDownloadState(false);
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
