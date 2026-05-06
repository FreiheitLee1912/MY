/**
 * Gantt Chart Generator - CSV to PPT-Ready Gantt Chart
 * Parses Jira CSV exports and renders interactive Gantt charts
 */

// ========================================
// Global State
// ========================================
const state = {
    files: [],
    tasks: [],
    filteredTasks: [],
    viewMode: 'month', // month, quarter, year
    groupBy: 'none',   // default to no grouping
    selectedProject: 'all',
    theme: 'light',
    dayWidth: 4, // pixels per day
    customStart: null,  // user-set timeline start (Date or null)
    customEnd: null,    // user-set timeline end (Date or null)
};

// ========================================
// Task Type Configuration
// ========================================
const TYPE_CONFIG = {
    'Milestone':   { color: 'var(--milestone-color)',   hex: '#f6ad55', label: 'MS' },
    'Goal':        { color: 'var(--goal-color)',         hex: '#68d391', label: 'Goal' },
    'Review':      { color: 'var(--review-color)',       hex: '#63b3ed', label: 'Rev' },
    'Event':       { color: 'var(--event-color)',        hex: '#b794f4', label: 'Evt' },
    'Task':        { color: 'var(--task-color)',         hex: '#4299e1', label: 'Task' },
    'Sub task':    { color: 'var(--task-color)',         hex: '#4299e1', label: 'Sub' },
    'Validation':  { color: 'var(--validation-color)',   hex: '#f56565', label: 'Val' },
    'SOP':         { color: 'var(--sop-color)',          hex: '#ecc94b', label: 'SOP' },
    'Certificate': { color: 'var(--certificate-color)',  hex: '#4299e1', label: 'Cert' },
    'PPAP':        { color: 'var(--ppap-color)',         hex: '#fc8181', label: 'PPAP' },
    'Top-level initiative': { color: 'var(--accent-primary)', hex: '#3182ce', label: 'Init' },
};

const STATUS_COLORS = {
    'Done': '#48bb78',
    'Closed': '#48bb78',
    'To Do': '#f6ad55',
    'In Progress': '#63b3ed',
    'Start': '#ecc94b',
    'Process': '#68d391',
};

// ========================================
// CSV Parser
// ========================================
function* parseCSVGenerator(text) {
    let current = '';
    let inQuotes = false;
    let fields = [];
    
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (ch === '"') {
            if (inQuotes && text[i + 1] === '"') {
                current += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (ch === ',' && !inQuotes) {
            fields.push(current.trim());
            current = '';
        } else if ((ch === '\n' || ch === '\r') && !inQuotes) {
            if (ch === '\r' && text[i + 1] === '\n') i++;
            fields.push(current.trim());
            current = '';
            if (fields.some(f => f !== '')) {
                yield fields;
            }
            fields = [];
        } else {
            current += ch;
        }
    }
    if (current || fields.length > 0) {
        fields.push(current.trim());
        if (fields.some(f => f !== '')) {
            yield fields;
        }
    }
}

// ========================================
// Date Parser
// ========================================
function parseDate(dateStr) {
    if (!dateStr || dateStr.trim() === '') return null;
    
    let str = dateStr.trim();
    
    // Remove everything from the first HH:MM onwards.
    // e.g. "06/12/26 12:00 AM" -> "06/12/26"
    str = str.replace(/\s+\d{1,2}:\d{2}.*$/, '').trim();
    
    // Format: DD/M/YY  (Jira export: day first, 2-digit year last  e.g. "06/12/26" = Dec 6 2026)
    const slashMatch = str.match(/^(\d{1,4})\/(\d{1,2})\/(\d{1,4})$/);
    if (slashMatch) {
        const a = parseInt(slashMatch[1]);
        const b = parseInt(slashMatch[2]);
        const c = parseInt(slashMatch[3]);
        let year, month, day;
        if (a > 31) {
            // YYYY/M/D  (4-digit year first)
            year = a; month = b - 1; day = c;
        } else if (c <= 99) {
            // DD/M/YY  (2-digit year last) 窶・Jira format
            day = a; month = b - 1; year = c + 2000;
        } else {
            // D/M/YYYY (4-digit year last)
            day = a; month = b - 1; year = c;
        }
        const result = new Date(year, month, day);
        if (!isNaN(result.getTime())) return result;
        return null;
    }
    
    // Format: YYYY-MM-DD
    const dashMatch = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (dashMatch) {
        return new Date(parseInt(dashMatch[1]), parseInt(dashMatch[2]) - 1, parseInt(dashMatch[3]));
    }
    
    // Try native parse as last resort
    const d = new Date(str);
    if (!isNaN(d.getTime())) return d;
    
    return null;
}

// Format a Date object to YYYY-MM-DD string for server API
function formatDate(date) {
    if (!date) return '';
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
}

// ========================================
// File Loading
// ========================================
function loadCSVFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const text = e.target.result;
            const tasks = parseCSVToTasks(text, file.name);
            resolve(tasks);
        };
        reader.onerror = reject;
        // Try UTF-8 first
        reader.readAsText(file, 'UTF-8');
    });
}

function parseCSVToTasks(text, filename) {
    const tasks = [];
    let headers = null;
    
    for (const row of parseCSVGenerator(text)) {
        if (!headers) {
            headers = row.map(h => h.toLowerCase().replace(/\s+/g, ''));
            continue;
        }
        
        // A=Key, B=Type, C=Summary, D=Status, E=Assignee, F=StartDate, G=EstStart, H=DueDate, I=EstDue
        const key      = row[0] || '';
        const type     = row[1] || '';
        const summary  = row[2] || '';
        const status   = row[3] || '';
        const assignee = row[4] || '';
        const startDate   = row[5] || '';
        const estStart    = row[6] || '';
        const deadline    = row[7] || '';
        const estDeadline = row[8] || '';
        
        // Determine effective start and end dates
        const effectiveStart = parseDate(startDate) || parseDate(estStart);
        const effectiveEnd = parseDate(deadline) || parseDate(estDeadline);
        
        if (key) {
            tasks.push({
                key,
                type: type || 'Task',
                parent: '',
                summary: summary || key,
                status: status || 'To Do',
                assignee: assignee || '',
                startDate: effectiveStart,
                endDate: effectiveEnd,
                startDateRaw: startDate,
                endDateRaw: deadline || estDeadline,
                source: filename,
                project: key.split('-')[0],
            });
        }
    }
    
    return tasks;
}

// ========================================
// Task Filtering & Grouping
// ========================================
function filterAndGroupTasks() {
    let tasks = [...state.tasks];
    
    // Filter by project
    if (state.selectedProject !== 'all') {
        tasks = tasks.filter(t => t.project === state.selectedProject);
    }
    
    // Only show tasks with dates
    const withDates = tasks.filter(t => t.startDate || t.endDate);
    // Group tasks
    let grouped = [];
    
    if (state.groupBy === 'type') {
        const typeOrder = ['Milestone', 'Top-level initiative', 'Goal', 'Review', 'Event', 'SOP', 'Certificate', 'Validation', 'PPAP', 'Task', 'Sub task'];
        const groups = {};
        
        for (const task of withDates) {
            const g = task.type || 'Other';
            if (!groups[g]) groups[g] = [];
            groups[g].push(task);
        }
        
        // Sort groups by type order
        const sortedKeys = Object.keys(groups).sort((a, b) => {
            const ia = typeOrder.indexOf(a);
            const ib = typeOrder.indexOf(b);
            return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
        });
        
        for (const groupName of sortedKeys) {
            grouped.push({ isGroup: true, label: groupName, count: groups[groupName].length });
            // Sort tasks within group by start date
            const sorted = groups[groupName].sort((a, b) => {
                const da = a.startDate || new Date(9999, 0);
                const db = b.startDate || new Date(9999, 0);
                return da - db;
            });
            grouped.push(...sorted);
        }
    } else if (state.groupBy === 'parent') {
        const rootTasks = withDates.filter(t => !t.parent);
        const childTasks = withDates.filter(t => t.parent);
        
        // Sort root tasks by start date
        rootTasks.sort((a, b) => {
            const da = a.startDate || new Date(9999, 0);
            const db = b.startDate || new Date(9999, 0);
            return da - db;
        });
        
        for (const root of rootTasks) {
            grouped.push(root);
            const children = childTasks.filter(c => c.parent === root.key);
            children.sort((a, b) => {
                const da = a.startDate || new Date(9999, 0);
                const db = b.startDate || new Date(9999, 0);
                return da - db;
            });
            for (const child of children) {
                child._isChild = true;
                grouped.push(child);
            }
        }
        
        // Add orphan children
        const usedChildren = new Set(grouped.filter(t => t._isChild).map(t => t.key));
        const orphans = childTasks.filter(c => !usedChildren.has(c.key));
        if (orphans.length > 0) {
            grouped.push({ isGroup: true, label: 'Other Tasks', count: orphans.length });
            grouped.push(...orphans);
        }
    } else {
        // No grouping - sort by start date
        withDates.sort((a, b) => {
            const da = a.startDate || new Date(9999, 0);
            const db = b.startDate || new Date(9999, 0);
            return da - db;
        });
        grouped = withDates;
    }
    
    state.filteredTasks = grouped;
    return grouped;
}

// ========================================
// Timeline Calculations
// ========================================
function getTimelineBounds() {
    const tasks = state.filteredTasks.filter(t => !t.isGroup);
    if (tasks.length === 0) return { start: new Date(), end: new Date() };
    
    let earliest = null;
    let latest = null;
    
    for (const task of tasks) {
        const s = task.startDate;
        const e = task.endDate;
        
        if (s && (!earliest || s < earliest)) earliest = s;
        if (e && (!latest || e > latest)) latest = e;
        if (s && (!latest || s > latest)) latest = s;
        if (e && (!earliest || e < earliest)) earliest = e;
    }
    
    if (!earliest) earliest = new Date();
    if (!latest) latest = new Date();
    
    // Use custom bounds if set by user, otherwise auto-detect with padding
    const start = state.customStart 
        ? new Date(state.customStart.getFullYear(), state.customStart.getMonth(), 1)
        : new Date(earliest.getFullYear(), earliest.getMonth() - 1, 1);
    const end = state.customEnd 
        ? new Date(state.customEnd.getFullYear(), state.customEnd.getMonth() + 1, 0)
        : new Date(latest.getFullYear(), latest.getMonth() + 2, 0);
    
    return { start, end };
}

function getMonthsBetween(start, end) {
    const months = [];
    const current = new Date(start.getFullYear(), start.getMonth(), 1);
    
    while (current <= end) {
        months.push(new Date(current));
        current.setMonth(current.getMonth() + 1);
    }
    
    return months;
}

function daysBetween(a, b) {
    return Math.round((b - a) / (1000 * 60 * 60 * 24));
}

function daysInMonth(date) {
    return new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
}

// ========================================
// Rendering
// ========================================
function renderGantt() {
    const tasks = filterAndGroupTasks();
    if (tasks.length === 0) return;
    
    const bounds = getTimelineBounds();
    const months = getMonthsBetween(bounds.start, bounds.end);
    
    // Adjust day width based on view mode
    switch (state.viewMode) {
        case 'month':  state.dayWidth = 4; break;
        case 'quarter': state.dayWidth = 2; break;
        case 'year':   state.dayWidth = 1; break;
    }
    
    renderLabels(tasks);
    renderTimelineHeader(months);
    renderTimelineBody(tasks, bounds, months);
    
    // Show sections
    document.getElementById('gantt-section').classList.remove('hidden');
    document.getElementById('controls-section').classList.remove('hidden');
}

function renderLabels(tasks) {
    const container = document.getElementById('gantt-labels');
    container.innerHTML = '';
    
    // Header
    const header = document.createElement('div');
    header.className = 'gantt-label-header';
    header.innerHTML = `
        <span style="min-width:30px"></span>
        <span style="min-width:65px">Key</span>
        <span style="min-width:45px">Type</span>
        <span style="flex:1">Summary</span>
        <span style="min-width:85px">Start</span>
        <span style="min-width:85px">End</span>
    `;
    container.appendChild(header);
    
    // Rows
    tasks.forEach((item, index) => {
        const row = document.createElement('div');
        
        if (item.isGroup) {
            row.className = 'gantt-label-row group-header';
            row.innerHTML = `<span style="flex:1">${item.label} (${item.count})</span>`;
        } else {
            const isChild = item._isChild || item.parent;
            row.className = `gantt-label-row${isChild ? ' child-row' : ''}`;
            
            const typeConf = TYPE_CONFIG[item.type] || { color: 'var(--default-color)', hex: '#828282', label: '?' };
            row.innerHTML = `
                <div class="move-controls">
                    <button class="move-btn" onclick="moveTask('${item.key}', -1)">^</button>
                    <button class="move-btn" onclick="moveTask('${item.key}', 1)">v</button>
                </div>
                <span class="task-key">${item.key}</span>
                <span class="task-type-badge" style="background:${typeConf.hex}22;color:${typeConf.hex}">${typeConf.label}</span>
                <span class="task-name editable" contenteditable="true" onblur="updateTask('${item.key}', 'summary', this.innerText)">${item.summary}</span>
                <span class="task-date editable" contenteditable="true" onblur="updateTask('${item.key}', 'startDate', this.innerText)">${formatDate(item.startDate)}</span>
                <span class="task-date editable" contenteditable="true" onblur="updateTask('${item.key}', 'endDate', this.innerText)">${formatDate(item.endDate)}</span>
            `;
        }
        
        container.appendChild(row);
    });
}

function moveTask(key, direction) {
    const idx = state.tasks.findIndex(t => t.key === key);
    if (idx === -1) return;
    const newIdx = idx + direction;
    if (newIdx < 0 || newIdx >= state.tasks.length) return;
    
    const temp = state.tasks[idx];
    state.tasks[idx] = state.tasks[newIdx];
    state.tasks[newIdx] = temp;
    
    renderGantt();
}

function updateTask(key, field, value) {
    const task = state.tasks.find(t => t.key === key);
    if (!task) return;
    
    if (field === 'startDate' || field === 'endDate') {
        const d = parseDate(value);
        if (d) {
            task[field] = d;
            renderGantt(); // Re-render to update bar position
        } else {
            // Restore if invalid
            renderGantt();
        }
    } else {
        task[field] = value;
        // Don't necessarily need to re-render everything for summary edit unless bar has label
        renderGantt();
    }
}

function renderTimelineHeader(months) {
    const container = document.getElementById('gantt-timeline-header');
    container.innerHTML = '';
    
    // Group months by year
    const years = {};
    months.forEach(m => {
        const y = m.getFullYear();
        if (!years[y]) years[y] = [];
        years[y].push(m);
    });

    // Create Year Row
    const yearRow = document.createElement('div');
    yearRow.className = 'timeline-year-row';
    for (const year in years) {
        const yearMonths = years[year];
        const width = yearMonths.reduce((sum, m) => sum + daysInMonth(m) * state.dayWidth, 0);
        const yearEl = document.createElement('div');
        yearEl.className = 'timeline-year-label';
        yearEl.style.width = `${width}px`;
        yearEl.textContent = year;
        yearRow.appendChild(yearEl);
    }
    container.appendChild(yearRow);

    // Create Month Row
    const monthRow = document.createElement('div');
    monthRow.className = 'timeline-month-row';
    for (const month of months) {
        const days = daysInMonth(month);
        const width = days * state.dayWidth;
        const monthEl = document.createElement('div');
        monthEl.className = 'timeline-month-label-hier';
        monthEl.style.width = `${width}px`;
        monthEl.textContent = month.getMonth() + 1;
        monthRow.appendChild(monthEl);
    }
    container.appendChild(monthRow);
}

function renderTimelineBody(tasks, bounds, months) {
    const container = document.getElementById('gantt-timeline-body');
    container.innerHTML = '';
    
    // Calculate total width
    let totalWidth = 0;
    for (const month of months) {
        totalWidth += daysInMonth(month) * state.dayWidth;
    }
    container.style.width = `${totalWidth}px`;
    
    // Draw grid lines (month boundaries)
    let xOffset = 0;
    for (const month of months) {
        const line = document.createElement('div');
        line.className = 'gantt-grid-line month-line';
        line.style.left = `${xOffset}px`;
        container.appendChild(line);
        xOffset += daysInMonth(month) * state.dayWidth;
    }
    
    // Today line
    const today = new Date();
    const todayOffset = daysBetween(bounds.start, today) * state.dayWidth;
    if (todayOffset >= 0 && todayOffset <= totalWidth) {
        const todayLine = document.createElement('div');
        todayLine.className = 'gantt-today-line';
        todayLine.style.left = `${todayOffset}px`;
        container.appendChild(todayLine);
    }
    
    // Task bars
    for (let i = 0; i < tasks.length; i++) {
        const item = tasks[i];
        const rowEl = document.createElement('div');
        rowEl.className = `gantt-row${item.isGroup ? ' group-header' : ''}`;
        
        if (!item.isGroup && (item.startDate || item.endDate)) {
            const start = item.startDate || item.endDate;
            const end = item.endDate || item.startDate;
            
            const startOffset = daysBetween(bounds.start, start) * state.dayWidth;
            const duration = Math.max(daysBetween(start, end), 1) * state.dayWidth;
            
            const typeConf = TYPE_CONFIG[item.type] || { color: 'var(--default-color)', hex: '#828282' };
            const isMilestone = item.type === 'Milestone' || (daysBetween(start, end) === 0 && item.type !== 'Task' && item.type !== 'Sub task');
            
            const bar = document.createElement('div');
            bar.className = `gantt-bar${isMilestone ? ' milestone' : ''}`;
            bar.style.left = `${startOffset}px`;
            
            if (isMilestone) {
                bar.style.background = typeConf.hex;
            } else {
                bar.style.width = `${Math.max(duration, 6)}px`;
                bar.style.background = `linear-gradient(135deg, ${typeConf.hex}, ${typeConf.hex}bb)`;
                
                // Show label if bar is wide enough
                if (duration > 60) {
                    const label = document.createElement('span');
                    label.className = 'gantt-bar-label';
                    label.textContent = item.summary;
                    bar.appendChild(label);
                }
            }
            
            // Completed tasks get a different style
            if (item.status === 'Done' || item.status === 'Closed') {
                bar.style.background = `linear-gradient(135deg, #27ae60, #27ae60bb)`;
                bar.style.opacity = '0.8';
            }
            
            // Tooltip data
            bar.dataset.key = item.key;
            bar.dataset.summary = item.summary;
            bar.dataset.type = item.type;
            bar.dataset.status = item.status;
            bar.dataset.start = formatDate(item.startDate);
            bar.dataset.end = formatDate(item.endDate);
            bar.dataset.assignee = item.assignee;
            
            bar.addEventListener('mouseenter', showTooltip);
            bar.addEventListener('mouseleave', hideTooltip);
            bar.addEventListener('mousemove', moveTooltip);
            
            // Stagger animation
            bar.style.animationDelay = `${i * 0.02}s`;
            
            rowEl.appendChild(bar);
        }
        
        container.appendChild(rowEl);
    }
}

// ========================================
// Tooltip
// ========================================
let tooltipEl = null;

function createTooltip() {
    tooltipEl = document.createElement('div');
    tooltipEl.className = 'gantt-tooltip';
    tooltipEl.id = 'gantt-tooltip';
    document.body.appendChild(tooltipEl);
}

function showTooltip(e) {
    if (!tooltipEl) createTooltip();
    
    const bar = e.currentTarget;
    tooltipEl.innerHTML = `
        <div class="tooltip-title">${bar.dataset.summary}</div>
        <div class="tooltip-row"><span>Key:</span><span>${bar.dataset.key}</span></div>
        <div class="tooltip-row"><span>Type:</span><span>${bar.dataset.type}</span></div>
        <div class="tooltip-row"><span>Status:</span><span>${bar.dataset.status}</span></div>
        <div class="tooltip-row"><span>Start:</span><span>${bar.dataset.start || 'Not set'}</span></div>
        <div class="tooltip-row"><span>End:</span><span>${bar.dataset.end || 'Not set'}</span></div>
        ${bar.dataset.assignee ? `<div class="tooltip-row"><span>Assignee:</span><span>${bar.dataset.assignee}</span></div>` : ''}
    `;
    tooltipEl.classList.add('visible');
}

function hideTooltip() {
    if (tooltipEl) tooltipEl.classList.remove('visible');
}

function moveTooltip(e) {
    if (!tooltipEl) return;
    const x = e.clientX + 12;
    const y = e.clientY + 12;
    tooltipEl.style.left = `${x}px`;
    tooltipEl.style.top = `${y}px`;
}

// ========================================
// Date Formatting
// ========================================
function formatDate(date) {
    if (!date) return '';
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}/${m}/${d}`;
}

function isLocalServer() {
    return ['localhost', '127.0.0.1', '::1'].includes(window.location.hostname);
}

function loadScript(src) {
    return new Promise((resolve, reject) => {
        const existing = document.querySelector(`script[src="${src}"]`);
        if (existing) {
            existing.addEventListener('load', resolve, { once: true });
            existing.addEventListener('error', reject, { once: true });
            if (window.PptxGenJS || window.pptxgen || window.pptxgenjs) resolve();
            return;
        }

        const script = document.createElement('script');
        script.src = src;
        script.onload = resolve;
        script.onerror = () => reject(new Error(`Failed to load ${src}`));
        document.head.appendChild(script);
    });
}

async function getPptxConstructor() {
    if (window.PptxGenJS || window.pptxgen || window.pptxgenjs) {
        return window.PptxGenJS || window.pptxgen || window.pptxgenjs;
    }

    const sourceSets = [
        ['vendor/jszip.min.js', 'vendor/pptxgen.min.js'],
        [
            'https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.12.0/libs/jszip.min.js',
            'https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.12.0/dist/pptxgen.min.js',
        ],
        ['https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.12.0/dist/pptxgen.bundle.js'],
    ];

    for (const sources of sourceSets) {
        try {
            for (const src of sources) {
                await loadScript(src);
            }
            const constructor = window.PptxGenJS || window.pptxgen || window.pptxgenjs;
            if (constructor) return constructor;
        } catch (err) {
            console.warn(err.message);
        }
    }

    throw new Error('PptxGenJS library could not be loaded.');
}

async function captureGanttImage() {
    const ganttWrapper = document.getElementById('gantt-wrapper');
    const container = document.getElementById('gantt-container');
    const origWrapperOverflow = ganttWrapper.style.overflow;
    const origContainerOverflow = container.style.overflow;
    const origWrapperWidth = ganttWrapper.style.width;

    ganttWrapper.style.overflow = 'visible';
    container.style.overflow = 'visible';
    ganttWrapper.style.width = `${container.scrollWidth}px`;

    try {
        const canvas = await html2canvas(ganttWrapper, {
            backgroundColor: getComputedStyle(document.documentElement).getPropertyValue('--bg-primary').trim(),
            scale: 2,
            logging: false,
            useCORS: true,
            windowWidth: Math.max(document.documentElement.clientWidth, container.scrollWidth),
        });
        return canvas.toDataURL('image/png');
    } finally {
        ganttWrapper.style.overflow = origWrapperOverflow;
        container.style.overflow = origContainerOverflow;
        ganttWrapper.style.width = origWrapperWidth;
    }
}

async function generatePPTXInBrowser(tasks, title, filename) {
    const PptxConstructor = await getPptxConstructor();
    const pptx = new PptxConstructor();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = 'Gantt Chart Generator';
    pptx.subject = 'CSV Gantt export';
    pptx.title = title;
    pptx.company = 'Browser export';

    const slide = pptx.addSlide();
    slide.background = { color: 'F8FAFC' };
    slide.addText(title, {
        x: 0.35, y: 0.18, w: 8.0, h: 0.35,
        fontFace: 'Aptos',
        fontSize: 18,
        bold: true,
        color: '1F2937',
        margin: 0,
    });
    slide.addText(`Exported ${new Date().toLocaleDateString()}`, {
        x: 0.35, y: 0.55, w: 3.0, h: 0.2,
        fontFace: 'Aptos',
        fontSize: 8,
        color: '64748B',
        margin: 0,
    });
    slide.addImage({
        data: await captureGanttImage(),
        x: 0.25,
        y: 0.85,
        w: 12.85,
        h: 6.25,
        sizing: { type: 'contain', x: 0.25, y: 0.85, w: 12.85, h: 6.25 },
    });

    await pptx.writeFile({ fileName: filename });
}

// ========================================
// Export Functions
// ========================================
async function exportPNG() {
    try {
        const imgData = await captureGanttImage();
        const filename = `gantt_${new Date().toISOString().slice(0,10)}.png`;
        
        // Save to browser download as fallback
        const link = document.createElement('a');
        link.download = filename;
        link.href = imgData;
        link.click();
        
        if (isLocalServer()) {
            try {
                const response = await fetch('/api/save-image', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ image: imgData, filename }),
                });

                if (response.ok) {
                    const result = await response.json();
                    if (result.success) {
                        showNotification(`PNG also saved to local output folder:\n${result.filename}`, 'success');
                        return;
                    }
                }
            } catch (serverErr) {
                console.warn('Local PNG save skipped:', serverErr);
            }
        }
        showNotification(`PNG downloaded:\n${filename}`, 'success');
    } catch (err) {
        console.error('PNG export failed:', err);
        alert('PNG export failed: ' + err.message);
    }
}

async function exportPPTX() {
    const tasks = state.filteredTasks.filter(t => !t.isGroup);
    if (tasks.length === 0) {
        alert('No visible tasks to export.');
        return;
    }
    
    // Get project name for title
    const projects = [...new Set(tasks.map(t => t.project))];
    const title = projects.join(' / ') + ' Gantt Chart';
    const filename = `gantt_${projects.join('_')}_${new Date().toISOString().slice(0,10)}.pptx`;
    
    // Prepare task data for server
    const taskData = tasks.map(t => ({
        key: t.key,
        type: t.type,
        summary: t.summary,
        status: t.status,
        assignee: t.assignee || '',
        parent: t.parent || '',
        startDate: t.startDate ? formatDate(t.startDate) : '',
        endDate: t.endDate ? formatDate(t.endDate) : '',
    }));
    
    // Show loading state on button
    const btn = document.getElementById('btn-export-pptx');
    const origHTML = btn.innerHTML;
    btn.innerHTML = '<span class="btn-spinner"></span> Generating...';
    btn.disabled = true;
    
    try {
        // Include custom timeline bounds if set
        const payload = { tasks: taskData, title, filename, groupBy: state.groupBy };
        if (state.customStart) payload.timelineStart = formatDate(state.customStart);
        if (state.customEnd) payload.timelineEnd = formatDate(state.customEnd);

        if (isLocalServer()) {
            try {
                const response = await fetch('/api/generate-pptx', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload),
                });

                if (response.ok) {
                    const result = await response.json();

                    if (result.success) {
                        showNotification(`PPTX saved to output folder:\n${result.filename}`, 'success');
                        return;
                    }

                    console.warn('Server PPTX export failed:', result.error);
                }
            } catch (serverErr) {
                console.warn('Server PPTX export skipped:', serverErr);
            }
        }

        await generatePPTXInBrowser(tasks, title, filename);
        showNotification(`PPTX downloaded:\n${filename}`, 'success');
    } catch (err) {
        console.error('PPTX export failed:', err);
        showNotification(`PPTX export failed: ${err.message}`, 'error');
    } finally {
        btn.innerHTML = origHTML;
        btn.disabled = false;
    }
}

// ========================================
// Notification
// ========================================
function showNotification(message, type = 'success') {
    // Remove existing notification
    const existing = document.getElementById('notification');
    if (existing) existing.remove();
    
    const notif = document.createElement('div');
    notif.id = 'notification';
    notif.className = `notification notification-${type}`;
    notif.innerHTML = `
        <div class="notif-icon">
            ${type === 'success' 
                ? '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg>'
                : '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>'
            }
        </div>
        <div class="notif-text">${message}</div>
        <button class="notif-close" onclick="this.parentElement.remove()">&times;</button>
    `;
    document.body.appendChild(notif);
    
    // Auto-remove after 5 seconds
    setTimeout(() => {
        if (notif.parentElement) {
            notif.classList.add('notification-hide');
            setTimeout(() => notif.remove(), 300);
        }
    }, 5000);
}

// ========================================
// Event Handlers
// ========================================
function initEventListeners() {
    // File input
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    
    dropZone.addEventListener('click', () => fileInput.click());
    
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });
    
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });
    
    dropZone.addEventListener('drop', async (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        const files = Array.from(e.dataTransfer.files).filter(f => f.name.endsWith('.csv'));
        await handleFiles(files);
    });
    
    fileInput.addEventListener('change', async (e) => {
        const files = Array.from(e.target.files);
        await handleFiles(files);
    });
    
    // View mode buttons
    document.querySelectorAll('.view-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.view-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.viewMode = btn.dataset.view;
            renderGantt();
        });
    });
    
    // Group by
    document.getElementById('group-by').addEventListener('change', (e) => {
        state.groupBy = e.target.value;
        renderGantt();
    });
    
    // Project select
    document.getElementById('project-select').addEventListener('change', (e) => {
        state.selectedProject = e.target.value;
        renderGantt();
    });
    
    // Export buttons
    document.getElementById('btn-export-png').addEventListener('click', exportPNG);
    document.getElementById('btn-export-pptx').addEventListener('click', exportPPTX);
    
    // Timeline date range controls
    document.getElementById('timeline-start').addEventListener('change', (e) => {
        state.customStart = e.target.value ? new Date(e.target.value) : null;
        renderGantt();
    });
    
    document.getElementById('timeline-end').addEventListener('change', (e) => {
        state.customEnd = e.target.value ? new Date(e.target.value) : null;
        renderGantt();
    });
    
    document.getElementById('btn-reset-range').addEventListener('click', () => {
        state.customStart = null;
        state.customEnd = null;
        document.getElementById('timeline-start').value = '';
        document.getElementById('timeline-end').value = '';
        renderGantt();
    });
    
    // Theme toggle
    document.getElementById('btn-toggle-theme').addEventListener('click', () => {
        state.theme = state.theme === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', state.theme);
    });
}

async function handleFiles(files) {
    const loadedFilesEl = document.getElementById('loaded-files');
    
    for (const file of files) {
        try {
            const tasks = await loadCSVFile(file);
            state.tasks.push(...tasks);
            state.files.push(file.name);
            
            // Show loaded file indicator
            const fileEl = document.createElement('div');
            fileEl.className = 'loaded-file';
            fileEl.innerHTML = `
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <polyline points="20 6 9 17 4 12"/>
                </svg>
                <span>${file.name} (${tasks.length} tasks)</span>
            `;
            loadedFilesEl.appendChild(fileEl);
        } catch (err) {
            console.error(`Failed to load ${file.name}:`, err);
        }
    }
    
    if (state.tasks.length > 0) {
        // Populate project selector
        const projects = [...new Set(state.tasks.map(t => t.project))];
        const select = document.getElementById('project-select');
        select.innerHTML = '<option value="all">All</option>';
        for (const proj of projects) {
            select.innerHTML += `<option value="${proj}">${proj}</option>`;
        }
        
        // Render after short delay for animation
        setTimeout(() => {
            document.getElementById('upload-section').style.minHeight = 'auto';
            renderGantt();
        }, 500);
    }
}

// ========================================
// Initialize
// ========================================
document.addEventListener('DOMContentLoaded', () => {
    // Set default light theme
    document.documentElement.setAttribute('data-theme', 'light');
    initEventListeners();
    createTooltip();
});


