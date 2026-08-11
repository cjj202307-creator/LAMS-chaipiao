/**
 * Web拆票应用 - UI处理逻辑
 */

let currentResult = null;
let currentHeaders = null;

// ========================
// 文件上传处理
// ========================
function initFileUpload() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');

    // 点击上传
    dropZone.addEventListener('click', () => fileInput.click());

    // 文件选择
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) handleFile(e.target.files[0]);
    });

    // 拖拽
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
    });
}

// ========================
// 处理上传的Excel文件
// ========================
async function handleFile(file) {
    const fileName = file.name;
    const fileExt = fileName.split('.').pop().toLowerCase();

    if (!['xlsx', 'xls', 'xlsm'].includes(fileExt)) {
        showAlert('请上传 Excel 文件（.xlsx / .xls / .xlsm）', 'error');
        return;
    }

    showProcessing(true);
    updateProgress('正在读取文件...');

    try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

        updateProgress('正在解析数据...');

        // 读取数据sheet
        const sheetName = SPLIT_CONFIG.sourceSheet;
        if (!workbook.SheetNames.includes(sheetName)) {
            showAlert(`未找到 "${sheetName}" sheet，请确认文件格式正确`, 'error');
            showProcessing(false);
            return;
        }

        const ws = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

        if (rows.length === 0) {
            showAlert('数据为空，请检查文件内容', 'error');
            showProcessing(false);
            return;
        }

        // 读取豁免清单
        let exemptionCodes = new Set(SPLIT_CONFIG.defaultExemptionCodes);
        if (workbook.SheetNames.includes(SPLIT_CONFIG.exemptionSheet)) {
            const wsEx = workbook.Sheets[SPLIT_CONFIG.exemptionSheet];
            const exRows = XLSX.utils.sheet_to_json(wsEx, { defval: '' });
            for (const row of exRows) {
                const code = cleanHSCodeForLoad(row['豁免税号'] || row[Object.keys(row)[0]] || '');
                if (code.length === 8) exemptionCodes.add(code);
            }
        }

        updateProgress(`正在执行拆票逻辑（${rows.length} 行数据）...`);

        // 拆票前深拷贝原始数据快照（只有原始列，无新增列）
        // 用快照做比对基准，确保即使 process() 内部修改了行对象也能检测到
        const originalSnapshot = rows.map(r => ({ ...r }));

        // 执行拆票
        const engine = new SplitEngine(SPLIT_CONFIG);
        const result = engine.process(rows, exemptionCodes);

        // uodId数据完整性校验：用快照（拆票前）和结果（拆票后）逐单元格比对
        updateProgress('正在执行uodId数据完整性校验...');
        const validationReport = engine.validateByUodId(originalSnapshot, result.resultRows);
        result.validation = validationReport;

        currentResult = result;
        currentHeaders = Object.keys(rows[0]);

        updateProgress('正在生成预览...');

        // 显示结果
        displayResults(result, rows);

        // 显示下载按钮
        document.getElementById('downloadSection').style.display = 'block';
        document.getElementById('uploadInfo').textContent =
            `${fileName} → ${rows.length} 行 → ${Object.keys(result.ticketNumbers).length} 票`;

        // 自动留底：生成结果副本存入浏览器并自动下载
        try {
            await autoSaveAudit(result, fileName, rows, file);
        } catch (e) {
            console.warn('留底保存失败', e);
        }

        showProcessing(false);

    } catch (err) {
        console.error(err);
        showAlert('处理文件时出错：' + err.message, 'error');
        showProcessing(false);
    }
}

// ========================
// HSCODE清洗（用于加载豁免清单）
// ========================
function cleanHSCodeForLoad(hsCode) {
    if (!hsCode) return '';
    hsCode = String(hsCode);
    let result = '';
    for (let i = 0; i < hsCode.length; i++) {
        const c = hsCode[i];
        if (c >= '0' && c <= '9') result += c;
    }
    if (result.length > 8) result = result.substring(0, 8);
    if (result.length < 8 && result.length > 0) {
        result = '0'.repeat(8 - result.length) + result;
    }
    return result;
}

// ========================
// 显示结果
// ========================
function displayResults(result, originalRows) {
    // 统计信息
    const stats = document.getElementById('stats');
    const ticketCount = Object.keys(result.ticketNumbers).length;
    const rowCount = result.resultRows.filter(r => r !== null).length;
    const inconsistentCount = result.inconsistentInfo.highlightedRows.size;
    const validationPassed = result.validation ? result.validation.passed : false;

    stats.innerHTML = `
        <div class="stat-card">
            <div class="stat-value">${rowCount}</div>
            <div class="stat-label">数据行</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">${ticketCount}</div>
            <div class="stat-label">分票数</div>
        </div>
        <div class="stat-card">
            <div class="stat-value">${inconsistentCount}</div>
            <div class="stat-label">不一致行</div>
        </div>
        <div class="stat-card">
            <div class="stat-value" style="color: ${validationPassed ? '#0F6E56' : '#A32D2D'};">${validationPassed ? '\u2705' : '\u274c'}</div>
            <div class="stat-label">uodId校验</div>
        </div>
    `;

    // 预览表格
    displayTable(result);

    // 拆分理由
    displayReasons(result.reasons);

    // 数据校验
    displayValidation(result.validation);
}

// ========================
// 显示预览表格
// ========================
function displayTable(result) {
    const container = document.getElementById('resultTable');
    const displayColumns = [
        'uodId', '分票编号', '分票', '产品编号', '产品名称', '客户料号',
        '大PO', '小PO', '备案单位', '原产国', '出库备注',
        '账册号', '业务申报表号', '客户指令号'
    ];

    // 限制预览行数
    const maxPreviewRows = 200;
    const rows = result.resultRows;
    const previewRows = rows.length > maxPreviewRows ? rows.slice(0, maxPreviewRows) : rows;

    let html = '<table class="result-table"><thead><tr>';
    for (const col of displayColumns) {
        html += `<th>${col}</th>`;
    }
    html += '</tr></thead><tbody>';

    for (const row of previewRows) {
        if (row === null) {
            html += '<tr class="blank-row"><td colspan="' + displayColumns.length + '"></td></tr>';
            continue;
        }

        const isHighlighted = row._isInconsistent || false;
        const rowClass = isHighlighted ? 'highlighted-row' : '';

        html += `<tr class="${rowClass}">`;
        for (const col of displayColumns) {
            const value = row[col] != null ? String(row[col]) : '';
            const isRed = isHighlighted && row._redColumns && row._redColumns.has(col);
            const cellClass = isRed ? 'red-cell' : '';
            html += `<td class="${cellClass}">${escapeHtml(value)}</td>`;
        }
        html += '</tr>';
    }

    html += '</tbody></table>';

    if (rows.length > maxPreviewRows) {
        html += `<p class="preview-note">仅显示前 ${maxPreviewRows} 行，完整结果请下载Excel文件</p>`;
    }

    // 图例
    const hlDimLabels = SPLIT_CONFIG.highlightDimensions
        .map(function (d) { return d.label; }).join(' / ');
    html += `
        <div class="legend">
            <span class="legend-item"><span class="legend-color yellow-bg"></span>黄色行：同票内同产品编号，${hlDimLabels} 任一不一致</span>
            <span class="legend-item"><span class="legend-color red-font"></span>红色字体：具体不一致的单元格</span>
        </div>
    `;

    container.innerHTML = html;
}

// ========================
// 显示拆分理由
// ========================
function displayReasons(reasons) {
    const container = document.getElementById('reasonsTable');

    if (!reasons || reasons.length === 0) {
        container.innerHTML = '<p class="no-data">无拆分理由</p>';
        return;
    }

    let html = '<table class="reasons-table"><thead><tr>';
    html += '<th>分票编号</th><th>对应分票标记</th><th>拆分理由</th>';
    html += '</tr></thead><tbody>';

    for (const r of reasons) {
        html += '<tr>';
        html += `<td class="ticket-no">${escapeHtml(r.ticketNo)}</td>`;
        html += `<td class="mark-text">${escapeHtml(r.mark)}</td>`;
        html += `<td class="reason-text"><pre>${escapeHtml(r.reason)}</pre></td>`;
        html += '</tr>';
    }

    html += '</tbody></table>';
    container.innerHTML = html;
}

// ========================
// 显示数据校验结果
// ========================
function displayValidation(report) {
    const container = document.getElementById('validationTable');
    if (!report) {
        container.innerHTML = '<p class="no-data">校验未执行</p>';
        return;
    }

    let html = '';

    // 总体结论
    const statusClass = report.passed ? 'validation-pass' : 'validation-fail';
    const statusIcon = report.passed ? '\u2705' : '\u274c';
    const statusText = report.passed
        ? '数据完整性校验通过：无行丢失、无行多余、无重复、原始列全部一致'
        : `数据完整性校验未通过：发现 ${report.totalMismatches} 个问题`;

    html += `<div class="validation-summary ${statusClass}">
        <span class="validation-icon">${statusIcon}</span>
        <span class="validation-text">${statusText}</span>
    </div>`;

    // 检查项明细
    html += '<table class="validation-table"><tbody>';
    html += `<tr><td class="validation-label">原始数据行数</td><td>${report.originalCount}</td></tr>`;
    html += `<tr><td class="validation-label">结果数据行数</td><td>${report.resultCount}</td></tr>`;
    html += `<tr><td class="validation-label">行数匹配</td><td>${report.countMatch ? '\u2705 通过' : '\u274c 不一致'}</td></tr>`;
    html += `<tr><td class="validation-label">丢失的uodId</td><td>${report.missingIds.length === 0 ? '\u2705 无' : '\u274c ' + report.missingIds.length + ' 个'}</td></tr>`;
    html += `<tr><td class="validation-label">多余的uodId</td><td>${report.extraIds.length === 0 ? '\u2705 无' : '\u274c ' + report.extraIds.length + ' 个'}</td></tr>`;
    html += `<tr><td class="validation-label">重复的uodId</td><td>${report.duplicateIds.length === 0 ? '\u2705 无' : '\u274c ' + report.duplicateIds.length + ' 个'}</td></tr>`;
    html += `<tr><td class="validation-label">单元格错位</td><td>${report.cellMismatches.length === 0 ? '\u2705 无' : '\u274c ' + report.cellMismatches.length + ' 处'}</td></tr>`;
    html += '</tbody></table>';

    // 丢失的uodId明细
    if (report.missingIds.length > 0) {
        html += '<h4 class="validation-detail-title">丢失的uodId明细</h4>';
        html += '<table class="validation-table"><thead><tr><th>#</th><th>uodId</th></tr></thead><tbody>';
        for (let i = 0; i < Math.min(report.missingIds.length, 50); i++) {
            html += `<tr><td>${i + 1}</td><td>${escapeHtml(report.missingIds[i])}</td></tr>`;
        }
        html += '</tbody></table>';
        if (report.missingIds.length > 50) {
            html += `<p class="preview-note">仅显示前50条，共 ${report.missingIds.length} 条</p>`;
        }
    }

    // 多余的uodId明细
    if (report.extraIds.length > 0) {
        html += '<h4 class="validation-detail-title">多余的uodId明细</h4>';
        html += '<table class="validation-table"><thead><tr><th>#</th><th>uodId</th></tr></thead><tbody>';
        for (let i = 0; i < Math.min(report.extraIds.length, 50); i++) {
            html += `<tr><td>${i + 1}</td><td>${escapeHtml(report.extraIds[i])}</td></tr>`;
        }
        html += '</tbody></table>';
        if (report.extraIds.length > 50) {
            html += `<p class="preview-note">仅显示前50条，共 ${report.extraIds.length} 条</p>`;
        }
    }

    // 重复的uodId明细
    if (report.duplicateIds.length > 0) {
        html += '<h4 class="validation-detail-title">重复的uodId明细</h4>';
        html += '<table class="validation-table"><thead><tr><th>uodId</th><th>出现次数</th></tr></thead><tbody>';
        for (const d of report.duplicateIds) {
            html += `<tr><td>${escapeHtml(d.uodId)}</td><td>${d.count}</td></tr>`;
        }
        html += '</tbody></table>';
    }

    // 单元格错位明细
    if (report.cellMismatches.length > 0) {
        html += '<h4 class="validation-detail-title">单元格错位明细（前50条）</h4>';
        html += '<table class="validation-table"><thead><tr><th>uodId</th><th>列名</th><th>原始值</th><th>结果值</th></tr></thead><tbody>';
        for (let i = 0; i < Math.min(report.cellMismatches.length, 50); i++) {
            const m = report.cellMismatches[i];
            html += `<tr><td>${escapeHtml(m.uodId)}</td><td>${escapeHtml(m.column)}</td><td>${escapeHtml(m.originalValue)}</td><td>${escapeHtml(m.resultValue)}</td></tr>`;
        }
        html += '</tbody></table>';
        if (report.cellMismatches.length > 50) {
            html += `<p class="preview-note">仅显示前50条，共 ${report.cellMismatches.length} 条</p>`;
        }
    }

    container.innerHTML = html;
}

// ========================
// 下载结果
// ========================
// ========================
// Excel 样式工具（需 xlsx-js-style 支持写入单元格样式）
// ========================
const THIN_BORDER = { style: 'thin', color: { rgb: 'FFB0B0B0' } };

function fullBorder() {
    return { top: THIN_BORDER, bottom: THIN_BORDER, left: THIN_BORDER, right: THIN_BORDER };
}

// 遍历所有单元格并应用样式（styleFn(r, c, cell) -> 样式对象，写入 cell.s）
function applyCellStyle(ws, styleFn) {
    if (!ws || !ws['!ref']) return;
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
            const addr = XLSX.utils.encode_cell({ r, c });
            const cell = ws[addr];
            if (!cell) continue;
            const st = styleFn(r, c, cell);
            if (st) cell.s = st;
        }
    }
}

// 设置所有行高（磅/pt，1磅=1pt），不换行
function setRowHeights(ws, heightPt) {
    if (!ws || !ws['!ref']) return;
    const range = XLSX.utils.decode_range(ws['!ref']);
    ws['!rows'] = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
        ws['!rows'][r] = { hpt: heightPt, customHeight: true };
    }
}

// 主表样式：表头加粗+蓝底；高亮行整行黄底；差异单元格红字加粗；全部带细框线；行高14.5、不换行
function styleMainSheet(ws, outputHeaders, rowRefs) {
    applyCellStyle(ws, (r, c, cell) => {
        const style = { border: fullBorder(), alignment: { vertical: 'center', wrapText: false } };
        if (r === 0) {
            style.font = { bold: true, color: { rgb: 'FF1A1A1A' } };
            style.fill = { fgColor: { rgb: 'FFDCE6F1' } };
            style.alignment = { vertical: 'center', horizontal: 'center', wrapText: false };
            return style;
        }
        const rowObj = rowRefs[r - 1];
        if (rowObj && rowObj._isInconsistent) {
            style.fill = { fgColor: { rgb: 'FFFFFF00' } };  // 黄色高亮整行
            const colName = outputHeaders[c];
            if (rowObj._redColumns && rowObj._redColumns.has(colName)) {
                style.font = { bold: true, color: { rgb: 'FFFF0000' } };  // 红色加粗=差异单元格
            }
        }
        return style;
    });
}

// 通用网格样式：表头加粗+底色，其余带细框线（分票理由 / 数据校验 sheet）；行高14.5、不换行
function applyGridStyle(ws, headerRows) {
    applyCellStyle(ws, (r, c, cell) => {
        const style = { border: fullBorder(), alignment: { vertical: 'center', wrapText: false } };
        if (r < (headerRows || 1)) {
            style.font = { bold: true, color: { rgb: 'FF1A1A1A' } };
            style.fill = { fgColor: { rgb: 'FFDCE6F1' } };
            style.alignment = { vertical: 'center', horizontal: 'center', wrapText: false };
        }
        return style;
    });
}

// 构建结果工作簿（主表 + 分票理由 + 数据校验），含高亮/框线/行高14.5/不换行
function buildWorkbook(result) {
    const headers = currentHeaders;

    // 主表
    const outputHeaders = ['分票编号', '分票', ...headers];
    const aoa = [outputHeaders];
    const rowRefs = [];  // 与 aoa 数据行一一对应，指向原始行对象（用于还原高亮）

    for (const row of result.resultRows) {
        if (row === null) {
            aoa.push(new Array(outputHeaders.length).fill(''));
            rowRefs.push(null);
        } else {
            const rowData = [row['分票编号'] || '', row['分票'] || ''];
            for (const h of headers) {
                rowData.push(row[h] != null ? row[h] : '');
            }
            aoa.push(rowData);
            rowRefs.push(row);
        }
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 14 }, { wch: 40 }];
    for (let i = 2; i < outputHeaders.length; i++) ws['!cols'].push({ wch: 16 });
    styleMainSheet(ws, outputHeaders, rowRefs);
    setRowHeights(ws, 14.5);
    XLSX.utils.book_append_sheet(wb, ws, '中芯');

    // 分票理由sheet（带框线、行高14.5、不换行）
    const reasonHeaders = ['分票编号', '对应分票标记', '拆分理由'];
    const reasonAoa = [reasonHeaders];
    for (const r of result.reasons) {
        reasonAoa.push([r.ticketNo, r.mark, r.reason]);
    }
    const wsReason = XLSX.utils.aoa_to_sheet(reasonAoa);
    wsReason['!cols'] = [{ wch: 16 }, { wch: 40 }, { wch: 80 }];
    applyGridStyle(wsReason, 1);
    setRowHeights(wsReason, 14.5);
    XLSX.utils.book_append_sheet(wb, wsReason, '分票理由');

    // 数据校验sheet（带框线、行高14.5、不换行）
    const vReport = result.validation;
    if (vReport) {
        const vAoa = [
            ['数据完整性校验报告'],
            [],
            ['校验结果', vReport.passed ? '通过' : '未通过'],
            ['原始数据行数', vReport.originalCount],
            ['结果数据行数', vReport.resultCount],
            ['行数匹配', vReport.countMatch ? '是' : '否'],
            ['丢失uodId数', vReport.missingIds.length],
            ['多余uodId数', vReport.extraIds.length],
            ['重复uodId数', vReport.duplicateIds.length],
            ['单元格错位数', vReport.cellMismatches.length],
            ['总问题数', vReport.totalMismatches],
            []
        ];

        if (vReport.missingIds.length > 0) {
            vAoa.push(['=== 丢失的uodId ===']);
            vAoa.push(['uodId']);
            for (const id of vReport.missingIds) {
                vAoa.push([id]);
            }
            vAoa.push([]);
        }

        if (vReport.extraIds.length > 0) {
            vAoa.push(['=== 多余的uodId ===']);
            vAoa.push(['uodId']);
            for (const id of vReport.extraIds) {
                vAoa.push([id]);
            }
            vAoa.push([]);
        }

        if (vReport.duplicateIds.length > 0) {
            vAoa.push(['=== 重复的uodId ===']);
            vAoa.push(['uodId', '出现次数']);
            for (const d of vReport.duplicateIds) {
                vAoa.push([d.uodId, d.count]);
            }
            vAoa.push([]);
        }

        if (vReport.cellMismatches.length > 0) {
            vAoa.push(['=== 单元格错位明细 ===']);
            vAoa.push(['uodId', '列名', '原始值', '结果值']);
            for (const m of vReport.cellMismatches) {
                vAoa.push([m.uodId, m.column, m.originalValue, m.resultValue]);
            }
        }

        const wsValidation = XLSX.utils.aoa_to_sheet(vAoa);
        wsValidation['!cols'] = [{ wch: 22 }, { wch: 22 }, { wch: 30 }, { wch: 30 }];
        applyGridStyle(wsValidation, 1);
        setRowHeights(wsValidation, 14.5);
        XLSX.utils.book_append_sheet(wb, wsValidation, '数据校验');
    }

    return wb;
}

function downloadResult() {
    if (!currentResult) {
        showAlert('请先上传文件并完成拆票', 'error');
        return;
    }

    const wb = buildWorkbook(currentResult);
    const now = new Date();
    const ts = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    const fileName = `拆票结果_${ts}.xlsx`;

    XLSX.writeFile(wb, fileName);
    showAlert(`已下载：${fileName}`, 'success');
}

// ========================
// 留底（审计）功能：每次上传拆票后自动保存「处理后的结果文件」+「原始上传文件」两份副本
// 纯前端实现（无后端）：两份文件均存浏览器 IndexedDB，元数据存 localStorage，并自动下载到本机"下载"文件夹
// ========================
const AUDIT_LOG_KEY = 'lams_audit_log';
const AUDIT_LIMIT = 200;

function loadAuditLog() {
    try {
        const raw = localStorage.getItem(AUDIT_LOG_KEY);
        return raw ? JSON.parse(raw) : [];
    } catch (e) { return []; }
}
function saveAuditLog(log) {
    try { localStorage.setItem(AUDIT_LOG_KEY, JSON.stringify(log.slice(0, AUDIT_LIMIT))); } catch (e) {}
}

function openAuditDB() {
    return new Promise((resolve, reject) => {
        if (!window.indexedDB) { reject(new Error('浏览器不支持IndexedDB')); return; }
        const req = indexedDB.open('lams_audit_db', 1);
        req.onupgradeneeded = () => {
            if (!req.result.objectStoreNames.contains('results')) req.result.createObjectStore('results');
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}
function auditDBPut(key, blob) {
    return openAuditDB().then(db => new Promise((res, rej) => {
        const tx = db.transaction('results', 'readwrite');
        tx.objectStore('results').put(blob, key);
        tx.oncomplete = () => res();
        tx.onerror = () => rej(tx.error);
    }));
}
function auditDBGet(key) {
    return openAuditDB().then(db => new Promise((res, rej) => {
        const tx = db.transaction('results', 'readonly');
        const r = tx.objectStore('results').get(key);
        r.onsuccess = () => res(r.result);
        r.onerror = () => rej(r.error);
    }));
}

function workbookToBlob(wb) {
    const arr = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    return new Blob([arr], { type: 'application/octet-stream' });
}

// 每次拆票后调用：生成留底文件 + 留存原始上传文件 -> 存IndexedDB -> 写日志 -> 自动下载
async function autoSaveAudit(result, fileName, rows, file) {
    const wb = buildWorkbook(result);
    const blob = workbookToBlob(wb);

    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const ts = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
    const id = 'audit_' + now.getTime() + '_' + Math.random().toString(36).slice(2, 8);
    const downloadName = `拆票留底_${ts}.xlsx`;

    // 原始上传文件：尽量原样留存（用于防扯皮，可追溯到最初上传数据）
    let origId = null, origName = null, hasOrig = false;
    if (file && file.size) {
        origId = 'orig_' + id;
        origName = `${ts}_原始_${fileName || '上传文件'}`;
        hasOrig = true;
    }

    const record = {
        id: id,
        time: now.toLocaleString('zh-CN', { hour12: false }),
        fileName: fileName,
        fileSize: file && file.size ? file.size : 0,
        rowCount: rows.length,
        ticketCount: Object.keys(result.ticketNumbers).length,
        inconsistentCount: result.inconsistentInfo.highlightedRows.size,
        validationPassed: result.validation ? result.validation.passed : false,
        downloadName: downloadName,
        origId: origId,
        origName: origName,
        hasOrig: hasOrig
    };

    try { await auditDBPut(id, blob); } catch (e) { console.warn('留底文件存储失败', e); }
    if (hasOrig) {
        try { await auditDBPut(origId, file); } catch (e) { console.warn('原始文件存储失败', e); record.hasOrig = false; }
    }

    const log = loadAuditLog();
    log.unshift(record);
    saveAuditLog(log);

    // 自动下载留底文件（最佳努力；若浏览器拦截，可到"留底记录"页重新下载）
    try {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = downloadName;
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (e) { console.warn('自动下载留底失败', e); }

    // 自动下载原始上传文件（防扯皮：本地永久留存最初数据）
    if (hasOrig) {
        try {
            const url = URL.createObjectURL(file);
            const a = document.createElement('a');
            a.href = url; a.download = origName;
            document.body.appendChild(a); a.click(); a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 1000);
        } catch (e) { console.warn('自动下载原始文件失败', e); }
    }

    if (document.getElementById('auditTable')) displayAuditLog();
    showAlert('已自动留底：' + downloadName + (hasOrig ? '（含原始上传文件）' : ''), 'success');
}

function downloadAuditBlob(id, name) {
    auditDBGet(id).then(blob => {
        if (!blob) { showAlert('留底文件已不存在（可能清理了浏览器数据）', 'error'); return; }
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = name || '拆票留底.xlsx';
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }).catch(e => showAlert('读取留底失败：' + e.message, 'error'));
}

function downloadOrigBlob(origId, name) {
    auditDBGet(origId).then(blob => {
        if (!blob) { showAlert('原始文件已不存在（可能清理了浏览器数据）', 'error'); return; }
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = name || '原始上传文件';
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }).catch(e => showAlert('读取原始文件失败：' + e.message, 'error'));
}

function displayAuditLog() {
    const container = document.getElementById('auditTable');
    if (!container) return;
    const log = loadAuditLog();
    if (log.length === 0) {
        container.innerHTML = '<p class="no-data">暂无留底记录。每次上传并拆票后，系统会自动生成「处理后的结果文件」与「原始上传文件」两份留底，并保存在本浏览器、自动下载到本机"下载"文件夹。</p>';
        return;
    }
    let html = '<table class="audit-table"><thead><tr>'
        + '<th>时间</th><th>原始文件</th><th>大小</th><th>数据行</th><th>分票数</th><th>不一致行</th><th>校验</th><th>留底文件</th><th>原始上传</th><th>操作</th>'
        + '</tr></thead><tbody>';
    for (const r of log) {
        const vClass = r.validationPassed ? 'ok' : 'bad';
        const vText = r.validationPassed ? '✅通过' : '❌未通过';
        const origBtn = r.hasOrig
            ? '<button class="mini-btn" data-orig-id="' + r.origId + '" data-orig-name="' + escapeHtml(r.origName) + '">下载原始</button>'
            : '<span class="audit-none">—</span>';
        html += '<tr>'
            + '<td>' + escapeHtml(r.time) + '</td>'
            + '<td>' + escapeHtml(r.fileName) + '</td>'
            + '<td>' + formatSize(r.fileSize) + '</td>'
            + '<td>' + r.rowCount + '</td>'
            + '<td>' + r.ticketCount + '</td>'
            + '<td>' + r.inconsistentCount + '</td>'
            + '<td class="' + vClass + '">' + vText + '</td>'
            + '<td class="audit-file">' + escapeHtml(r.downloadName) + '</td>'
            + '<td>' + origBtn + '</td>'
            + '<td><button class="mini-btn" data-audit-id="' + r.id + '" data-audit-name="' + escapeHtml(r.downloadName) + '">重新下载</button></td>'
            + '</tr>';
    }
    html += '</tbody></table>';
    html += '<div class="audit-actions">'
        + '<button class="mini-btn danger" id="clearAuditBtn">清空留底记录</button>'
        + '<span class="audit-hint">留底文件与原始上传文件均保存在本浏览器（IndexedDB），仅本机可查看/下载；每次上传会自动下载这两个文件到"下载"文件夹作为永久备份。清空仅删除记录与文件索引，不影响已导出的文件。</span>'
        + '</div>';
    container.innerHTML = html;

    container.querySelectorAll('[data-audit-id]').forEach(btn => {
        btn.addEventListener('click', () => downloadAuditBlob(btn.dataset.auditId, btn.dataset.auditName));
    });
    container.querySelectorAll('[data-orig-id]').forEach(btn => {
        btn.addEventListener('click', () => downloadOrigBlob(btn.dataset.origId, btn.dataset.origName));
    });
    const clearBtn = document.getElementById('clearAuditBtn');
    if (clearBtn) {
        clearBtn.addEventListener('click', () => {
            if (confirm('确认清空所有留底记录？此操作不可恢复（已导出的Excel不受影响）。')) {
                localStorage.removeItem(AUDIT_LOG_KEY);
                openAuditDB().then(db => {
                    const tx = db.transaction('results', 'readwrite');
                    tx.objectStore('results').clear();
                    tx.oncomplete = () => { displayAuditLog(); showAlert('留底记录已清空', 'success'); };
                }).catch(() => { displayAuditLog(); });
            }
        });
    }
}

// ========================
// UI辅助函数
// ========================
function showProcessing(show) {
    document.getElementById('processingOverlay').style.display = show ? 'flex' : 'none';
}

function updateProgress(msg) {
    document.getElementById('progressText').textContent = msg;
}

function showAlert(msg, type) {
    const alert = document.getElementById('alertBox');
    alert.textContent = msg;
    alert.className = 'alert ' + (type || 'info');
    alert.style.display = 'block';
    setTimeout(() => { alert.style.display = 'none'; }, 5000);
}

function formatSize(bytes) {
    if (!bytes || bytes <= 0) return '—';
    const kb = bytes / 1024;
    if (kb < 1024) return kb.toFixed(1) + ' KB';
    return (kb / 1024).toFixed(2) + ' MB';
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ========================
// 显示规则配置
// ========================
function displayConfig() {
    const container = document.getElementById('configDisplay');
    const cfg = SPLIT_CONFIG;
    const hlDims = cfg.highlightDimensions.map(function (d) { return '「' + d.label + '」'; }).join('、');

    // 一、整体流程
    const steps = [
        ['1', '读取数据', '从「' + cfg.sourceSheet + '」sheet 读取明细行；若含「' + cfg.exemptionSheet + '」sheet 则加载豁免HSCODE清单（否则用内置默认清单 ' + cfg.defaultExemptionCodes.length + ' 条）'],
        ['2', '生成基础标记', '按 账册号后三位 + 业务申报表号 + 客户类型 生成每行的基础分票标记'],
        ['3', '征税 / 免表判定', '根据客户类型、原产国、品名判定每行是「征税」还是「免表」'],
        ['4', '131豁免判定', '原产国为美国时，按HSCODE是否在豁免清单判定「131豁免外 / 非131豁免外」'],
        ['5', 'PO分组', '同一基础标记下按PO分组，每票PO数量不超过上限，超出自动拆成多票'],
        ['6', '行数限制', '永芯每票最多40行，超出自动拆成多票'],
        ['7', '一致性高亮', '同票内相同产品编号，若 ' + hlDims + ' 任一不同 → 整行黄色、差异单元格红色加粗'],
        ['8', '生成分票编号', '免表前缀MB、征税前缀TAX，格式为 前缀 + 月日 + 序号'],
        ['9', '生成分票理由', '逐票汇总本票包含的标记与拆分原因，便于追溯'],
        ['10', '数据完整性校验', '按uodId逐单元格比对拆票前后，确保无行丢失 / 多余 / 错位']
    ];

    const customerRows = cfg.customerTypes.map(function (ct) {
        return '<tr><td>' + escapeHtml(ct.label) + '</td><td>' + (ct.isTianjin ? '是（天津规则）' : '否') + '</td><td>' + (ct.isYongxin ? '是（40行上限）' : '否') + '</td></tr>';
    }).join('');

    function ruleRows(rules, def) {
        let rows = rules.map(function (r) {
            let cond;
            if (r.match.originCountry) {
                cond = '原产国包含「' + r.match.originCountry + '」' + (r.match.originCountryExclude ? ' 且不含「' + r.match.originCountryExclude + '」' : '');
            } else {
                cond = '品名包含 ' + r.match.productName.map(function (p) { return '「' + p + '」'; }).join('、');
            }
            const type = r.type === '征税' ? '征税' : '免表';
            return '<tr><td>' + type + '</td><td>' + cond + '</td><td>' + escapeHtml(r.reason) + '</td></tr>';
        }).join('');
        rows += '<tr><td>' + (def.type === '征税' ? '征税' : '免表') + '</td><td>其他情况</td><td>' + escapeHtml(def.reason) + '</td></tr>';
        return rows;
    }

    const beijingRows = ruleRows(cfg.beijingTaxRule.rules, cfg.beijingTaxRule.default);
    const tianjinRows = ruleRows(cfg.tianjinTaxRule.rules, cfg.tianjinTaxRule.default);

    const poRows = Object.keys(cfg.maxPOPerTicket).map(function (k) {
        const label = k === 'default' ? '其他客户（北京/北方/京城/永芯/未知）' : k;
        return '<tr><td>' + label + '</td><td>' + cfg.maxPOPerTicket[k] + ' 个PO / 票</td></tr>';
    }).join('');

    const rowRows = Object.keys(cfg.maxRowsPerTicket).map(function (k) {
        const label = k === 'default' ? '其他客户' : k;
        const val = cfg.maxRowsPerTicket[k] === null ? '无限制' : cfg.maxRowsPerTicket[k] + ' 行 / 票';
        return '<tr><td>' + label + '</td><td>' + val + '</td></tr>';
    }).join('');

    const ex = cfg.exemption131;
    const exText = '原产国为美国：HSCODE在豁免清单 → ' + ex.usOriginRule.inList.status + '；不在清单 → ' + ex.usOriginRule.notInList.status + '。原产国非美国 → ' + ex.nonUsRule.status + '。';

    const prefixText = '免表票前缀「' + cfg.ticketPrefix['免表'] + '」，征税票前缀「' + cfg.ticketPrefix['default'] + '」；编号格式：前缀 + 月日(MMDD) + 两位序号。';

    let html = '';

    html += '<div class="logic-section"><h3>一、整体处理流程</h3><ol class="logic-steps">';
    steps.forEach(function (s) {
        html += '<li><span class="step-no">' + s[0] + '</span><div><b>' + s[1] + '</b><br><span class="step-desc">' + s[2] + '</span></div></li>';
    });
    html += '</ol></div>';

    html += '<div class="logic-section"><h3>二、客户类型识别</h3><table class="logic-table"><thead><tr><th>客户类型</th><th>是否天津</th><th>是否永芯</th></tr></thead><tbody>' + customerRows + '</tbody></table></div>';

    html += '<div class="logic-section"><h3>三、征税 / 免表判定（北京 / 北方 / 京城 / 永芯）</h3><table class="logic-table"><thead><tr><th>结果</th><th>触发条件</th><th>说明</th></tr></thead><tbody>' + beijingRows + '</tbody></table></div>';

    html += '<div class="logic-section"><h3>四、征税 / 免表判定（天津中芯）</h3><table class="logic-table"><thead><tr><th>结果</th><th>触发条件</th><th>说明</th></tr></thead><tbody>' + tianjinRows + '</tbody></table></div>';

    html += '<div class="logic-section"><h3>五、PO分组上限</h3><table class="logic-table"><thead><tr><th>客户类型</th><th>上限</th></tr></thead><tbody>' + poRows + '</tbody></table></div>';

    html += '<div class="logic-section"><h3>六、单票行数上限</h3><table class="logic-table"><thead><tr><th>客户类型</th><th>上限</th></tr></thead><tbody>' + rowRows + '</tbody></table></div>';

    html += '<div class="logic-section"><h3>七、一致性高亮规则</h3><p class="logic-text">同一分票内，若<b>相同产品编号</b>的各行在 ' + hlDims + ' 维度上出现<b>不一致</b>，则：整行标记为<span class="yellow-bg">黄色</span>，具体不一致的单元格用<span class="red-font">红色加粗</span>标注。高亮仅为风险提示，不改变分票结果。</p></div>';

    html += '<div class="logic-section"><h3>八、131豁免判定</h3><p class="logic-text">' + exText + '</p></div>';

    html += '<div class="logic-section"><h3>九、分票编号规则</h3><p class="logic-text">' + prefixText + '</p></div>';

    html += '<div class="logic-section"><h3>十、数据完整性校验</h3><p class="logic-text">拆分完成后，按 <b>uodId</b> 将结果与原始数据逐行逐单元格比对：行数是否一致、是否有丢失 / 多余 / 重复的uodId、原始列（除新增"分票编号/分票"两列外）是否完全一致。任一异常都会在"数据校验"页与下载的Excel中提示。</p></div>';

    html += '<div class="logic-section logic-note"><p>以上逻辑为当前拆分数据实际使用逻辑</p></div>';

    container.innerHTML = html;
    // 同步渲染到首页常驻的「当前拆票规则」面板（业务人员一眼可见）
    const rp = document.getElementById('rulesPanel');
    if (rp) rp.innerHTML = html;
}

// ========================
// 初始化
// ========================

document.addEventListener('DOMContentLoaded', () => {
    initFileUpload();
    displayConfig();
    displayAuditLog();

    // 下载按钮
    document.getElementById('downloadBtn').addEventListener('click', downloadResult);

    // 标签页切换
    const tabs = document.querySelectorAll('.tab');
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            tabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            document.getElementById(tab.dataset.tab).classList.add('active');
        });
    });
});
