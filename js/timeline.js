// Excel風タイムライン描画エンジン（ドラッグ操作対応・マルチモード版）
// 表示モード: 週(week) / 5日(fiveday) / 1日(day) / 全体(fit)

// ========================================
// 定数
// ========================================
const ROW_HEIGHT = 32;
const FIXED_COL_WIDTH_CAT = 120;
const FIXED_COL_WIDTH_EVT = 120;
const FIXED_COLS_TOTAL = FIXED_COL_WIDTH_CAT + FIXED_COL_WIDTH_EVT;
const MONTH_NAMES = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"];
const RESIZE_ZONE_PX = 10;
const DRAG_THRESHOLD = 5;
const _isMobile = () => window.innerWidth <= 480;

// モードごとの日付開始配列
const MODE_STARTS = {
    week: [1, 8, 15, 22],
    fiveday: [1, 6, 11, 16, 21, 26],
    fit: [1, 8, 15, 22],
};
// モードごとの列幅 (px)。fit は動的計算のため null
const MODE_COL_WIDTHS = { week: 80, fiveday: 48, day: 24, fit: null };

// ========================================
// タイムラインモード状態
// ========================================
let _tlMode = "day";
let _tlColOffsets = [0]; // 初期化時に再構築
let _tlTotalCols = 365;
let _tlColWidth = 24;
let _tlYear = null; // 初回描画でsetTimelineModeを確実に呼ぶ

// モードを設定し、列オフセットテーブルを再構築
function setTimelineMode(modeId, year) {
    _tlMode = modeId;
    _tlYear = year;
    _tlColOffsets = [0];
    for (let m = 0; m < 12; m++) {
        _tlColOffsets.push(_tlColOffsets[m] + _colsPerMonth(year, m));
    }
    _tlTotalCols = _tlColOffsets[12];
    _tlColWidth = MODE_COL_WIDTHS[modeId] || 80;
}

// 月あたりの列数
function _colsPerMonth(year, month) {
    if (_tlMode === "day") return daysInMonth(year, month);
    const starts = MODE_STARTS[_tlMode];
    return starts ? starts.length : 4;
}

// ========================================
// ユーティリティ（共通関数は utils.js を参照）
// ========================================

// ========================================
// 汎用列マッピング関数（モード対応）
// ========================================

// dateStr → 絶対列番号
function absCol(dateStr) {
    const d = parseDateStr(dateStr);
    const month = d.getMonth();
    const day = d.getDate();

    if (_tlMode === "day") {
        return _tlColOffsets[month] + day - 1;
    }

    const starts = MODE_STARTS[_tlMode];
    let wi = 0;
    for (let i = starts.length - 1; i >= 0; i--) {
        if (day >= starts[i]) { wi = i; break; }
    }
    return _tlColOffsets[month] + wi;
}

// 列 → 月番号
function _colToMonth(colIdx) {
    for (let m = 0; m < 12; m++) {
        if (colIdx < _tlColOffsets[m + 1]) return m;
    }
    return 11;
}

// 列 → 月内インデックス
function _colToLocalIdx(colIdx) {
    const m = _colToMonth(colIdx);
    return colIdx - _tlColOffsets[m];
}

// 列 → 開始日
function colToStartDate(colIdx, year) {
    const month = _colToMonth(colIdx);
    const localIdx = colIdx - _tlColOffsets[month];

    if (_tlMode === "day") {
        const day = localIdx + 1;
        return `${year}-${String(month + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }

    const starts = MODE_STARTS[_tlMode];
    const day = starts[localIdx];
    return `${year}-${String(month + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

// 列 → 終了日
function colToEndDate(colIdx, year) {
    const month = _colToMonth(colIdx);
    const localIdx = colIdx - _tlColOffsets[month];

    if (_tlMode === "day") {
        // 1日モード: 開始日 = 終了日
        return colToStartDate(colIdx, year);
    }

    const starts = MODE_STARTS[_tlMode];
    const colsInMonth = starts.length;
    let lastDay;
    if (localIdx < colsInMonth - 1) {
        lastDay = starts[localIdx + 1] - 1;
    } else {
        lastDay = daysInMonth(year, month);
    }
    return `${year}-${String(month + 1).padStart(2, "0")}-${String(lastDay).padStart(2, "0")}`;
}

// ========================================
// カテゴリグループ化
// ========================================
function groupByCategory(events) {
    const groups = {};
    CATEGORIES.forEach(cat => { groups[cat.name] = []; });
    groups["__other__"] = [];
    events.forEach(ev => {
        if (ev.type === "holiday") return;
        const catNames = ev.categories || [];
        let placed = false;
        for (const cn of catNames) {
            if (CATEGORIES.find(c => c.name === cn)) {
                groups[cn].push(ev);
                placed = true;
                break;
            }
        }
        if (!placed) groups["__other__"].push(ev);
    });
    return groups;
}

// ========================================
// DragManager: ドラッグ操作の状態管理（パフォーマンス最適化版）
// ========================================
const DragManager = {
    active: false,
    started: false,
    mode: null,     // "create" | "move" | "resize-start" | "resize-end"
    startX: 0,
    startY: 0,
    startCol: -1,
    currentCol: -1,
    _prevCol: -1,
    row: null,
    eventId: null,
    eventObj: null,
    categoryName: null,
    originalStartCol: -1,
    originalEndCol: -1,
    year: null,

    _cells: null,
    _highlighted: [],
    _rafId: 0,
    _colGeometry: null,

    begin(mode, row, col, mouseX, mouseY, eventId, eventObj, catName, startCol, endCol, year) {
        PopoverManager.hide();
        this.active = true;
        this.started = false;
        this.mode = mode;
        this.startX = mouseX;
        this.startY = mouseY;
        this.startCol = col;
        this.currentCol = col;
        this._prevCol = -1;
        this.row = row;
        this.eventId = eventId;
        this.eventObj = eventObj;
        this.categoryName = catName;
        this.originalStartCol = startCol;
        this.originalEndCol = endCol;
        this.year = year;
        this._highlighted = [];

        const tds = row.querySelectorAll("td.gs-week-td");
        this._cells = new Array(_tlTotalCols).fill(null);
        for (let i = 0; i < tds.length; i++) {
            const c = parseInt(tds[i].dataset.col, 10);
            if (c < _tlTotalCols) this._cells[c] = tds[i];
        }

        this._cacheColGeometry();
        document.body.classList.add("gs-dragging");
    },

    _cacheColGeometry() {
        for (let i = 0; i < _tlTotalCols; i++) {
            if (this._cells[i]) {
                const rect = this._cells[i].getBoundingClientRect();
                this._colGeometry = { left: rect.left, width: rect.width, firstCol: i };
                return;
            }
        }
        this._colGeometry = null;
    },

    colFromClientX(clientX) {
        if (!this._colGeometry) return this.currentCol;
        const { width, firstCol } = this._colGeometry;
        const cell0 = this._cells[firstCol];
        if (!cell0) return this.currentCol;
        const currentLeft = cell0.getBoundingClientRect().left;
        const col = firstCol + Math.floor((clientX - currentLeft) / width);
        return Math.max(0, Math.min(_tlTotalCols - 1, col));
    },

    move(clientX, clientY) {
        if (!this.active) return;
        if (!this.started) {
            const dx = clientX - this.startX;
            const dy = clientY - this.startY;
            if (dx * dx + dy * dy < DRAG_THRESHOLD * DRAG_THRESHOLD) return;
            this.started = true;
        }
        const col = this.colFromClientX(clientX);
        this.currentCol = col;

        if (col === this._prevCol) return;
        this._prevCol = col;

        if (this._rafId) cancelAnimationFrame(this._rafId);
        this._rafId = requestAnimationFrame(() => {
            this._rafId = 0;
            this._updateVisual();
        });
    },

    end() {
        if (!this.active) return null;
        const wasStarted = this.started;
        if (this._rafId) { cancelAnimationFrame(this._rafId); this._rafId = 0; }
        const result = wasStarted ? this.computeResult() : null;
        this.cleanup();
        return result;
    },

    cancel() {
        if (this._rafId) { cancelAnimationFrame(this._rafId); this._rafId = 0; }
        this.cleanup();
    },

    cleanup() {
        const hl = this._highlighted;
        for (let i = 0; i < hl.length; i++) {
            hl[i].classList.remove("gs-drag-select", "gs-drag-ghost", "gs-drag-original");
        }
        this._highlighted = [];
        document.body.classList.remove("gs-dragging");
        this.active = false;
        this.started = false;
        this.mode = null;
        this.row = null;
        this.eventObj = null;
        this._cells = null;
        this._colGeometry = null;
    },

    computeResult() {
        const year = this.year;
        const maxCol = _tlTotalCols - 1;
        if (this.mode === "create") {
            const s = Math.min(this.startCol, this.currentCol);
            const e = Math.max(this.startCol, this.currentCol);
            return {
                action: "create",
                categoryName: this.categoryName,
                startDate: colToStartDate(s, year),
                endDate: colToEndDate(e, year),
            };
        }
        if (this.mode === "move") {
            const delta = this.currentCol - this.startCol;
            const newStart = this.originalStartCol + delta;
            const newEnd = this.originalEndCol + delta;
            if (newStart < 0 || newEnd > maxCol) return null;
            return {
                action: "move",
                eventId: this.eventId,
                eventObj: this.eventObj,
                startDate: colToStartDate(newStart, year),
                endDate: colToEndDate(newEnd, year),
            };
        }
        if (this.mode === "resize-start") {
            const newStart = Math.min(this.currentCol, this.originalEndCol);
            return {
                action: "resize",
                eventId: this.eventId,
                eventObj: this.eventObj,
                startDate: colToStartDate(newStart, year),
                endDate: colToEndDate(this.originalEndCol, year),
            };
        }
        if (this.mode === "resize-end") {
            const newEnd = Math.max(this.currentCol, this.originalStartCol);
            return {
                action: "resize",
                eventId: this.eventId,
                eventObj: this.eventObj,
                startDate: colToStartDate(this.originalStartCol, year),
                endDate: colToEndDate(newEnd, year),
            };
        }
        return null;
    },

    _updateVisual() {
        const hl = this._highlighted;
        for (let i = 0; i < hl.length; i++) {
            hl[i].classList.remove("gs-drag-select", "gs-drag-ghost", "gs-drag-original");
        }
        this._highlighted = [];

        const cells = this._cells;
        if (!cells) return;
        const maxCol = _tlTotalCols - 1;

        if (this.mode === "create") {
            const s = Math.min(this.startCol, this.currentCol);
            const e = Math.max(this.startCol, this.currentCol);
            for (let c = s; c <= e; c++) {
                if (cells[c]) { cells[c].classList.add("gs-drag-select"); this._highlighted.push(cells[c]); }
            }
        } else if (this.mode === "move") {
            const delta = this.currentCol - this.startCol;
            const ns = this.originalStartCol + delta;
            const ne = this.originalEndCol + delta;
            for (let c = this.originalStartCol; c <= this.originalEndCol; c++) {
                if (cells[c]) { cells[c].classList.add("gs-drag-original"); this._highlighted.push(cells[c]); }
            }
            if (ns >= 0 && ne <= maxCol) {
                for (let c = ns; c <= ne; c++) {
                    if (cells[c]) { cells[c].classList.add("gs-drag-ghost"); this._highlighted.push(cells[c]); }
                }
            }
        } else if (this.mode === "resize-start") {
            const ns = Math.min(this.currentCol, this.originalEndCol);
            for (let c = ns; c <= this.originalEndCol; c++) {
                if (cells[c]) { cells[c].classList.add("gs-drag-ghost"); this._highlighted.push(cells[c]); }
            }
        } else if (this.mode === "resize-end") {
            const ne = Math.max(this.currentCol, this.originalStartCol);
            for (let c = this.originalStartCol; c <= ne; c++) {
                if (cells[c]) { cells[c].classList.add("gs-drag-ghost"); this._highlighted.push(cells[c]); }
            }
        }
    },
};

// ========================================
// PopoverManager: ホバーポップオーバー管理
// ========================================
const PopoverManager = {
    _el: null,
    _showTimer: null,
    _hideTimer: null,
    _currentEventId: null,
    _bound: false,

    init() {
        this._el = document.getElementById("event-popover");
        if (!this._el || this._bound) return;
        this._bound = true;

        // ポップオーバー自体のホバーで消えないようにする
        this._el.addEventListener("mouseenter", () => {
            this._clearHideTimer();
        });
        this._el.addEventListener("mouseleave", () => {
            this._scheduleHide();
        });

        // ポップオーバークリックで編集モーダルを開く
        this._el.addEventListener("click", () => {
            const ev = this._currentEventId ? _eventsById[this._currentEventId] : null;
            if (ev) {
                this.hide();
                const catName = (ev.categories && ev.categories[0]) || null;
                openEventModal(ev, catName);
            }
        });

        // ポップオーバーにカーソルポインタを設定
        this._el.style.cursor = "pointer";

        // タッチデバイス: 外タップで閉じる
        document.addEventListener("touchstart", (e) => {
            if (!this._el || !this._el.classList.contains("gs-popover-visible")) return;
            if (this._el.contains(e.target)) return;
            if (e.target.closest && e.target.closest(".gs-bar-cell")) return;
            this.hide();
        }, { passive: true });
    },

    /** バーセルのmouseenter時に呼ばれる */
    scheduleShow(eventId, barCellEl) {
        if (DragManager.active) return;
        if (this._currentEventId === eventId && this._el && this._el.classList.contains("gs-popover-visible")) {
            this._clearHideTimer();
            return;
        }
        this._clearShowTimer();
        this._clearHideTimer();
        this._showTimer = setTimeout(() => {
            if (DragManager.active) return;
            this._show(eventId, barCellEl);
        }, 300);
    },

    /** バーセルのmouseleave時に呼ばれる */
    scheduleHide() {
        this._clearShowTimer();
        this._scheduleHide();
    },

    /** 即座に非表示（ドラッグ開始時、スクロール時など） */
    hide() {
        this._clearShowTimer();
        this._clearHideTimer();
        if (this._el) {
            this._el.classList.remove("gs-popover-visible");
            this._el.setAttribute("aria-hidden", "true");
        }
        this._currentEventId = null;
    },

    _show(eventId, anchorEl) {
        const ev = _eventsById[eventId];
        if (!ev) return;

        const cat = getCategoryDef((ev.categories && ev.categories[0]) || "");
        this._currentEventId = eventId;

        const el = this._el;
        el.querySelector(".gs-popover-accent").style.background = cat.color;
        el.querySelector(".gs-popover-title").textContent = ev.title;

        const badge = el.querySelector(".gs-popover-badge");
        badge.textContent = cat.name;
        badge.style.background = cat.color;

        const startFmt = this._formatDateJP(ev.startDate);
        const endFmt = this._formatDateJP(ev.endDate);
        el.querySelector(".gs-popover-dates").textContent =
            ev.startDate === ev.endDate ? startFmt : `${startFmt} 〜 ${endFmt}`;

        el.querySelector(".gs-popover-notes").textContent = ev.bodyPreview || "";

        this._position(anchorEl);
        el.classList.add("gs-popover-visible");
        el.setAttribute("aria-hidden", "false");
    },

    _position(anchorEl) {
        const el = this._el;
        const rect = anchorEl.getBoundingClientRect();
        const popW = _isMobile() ? Math.min(280, window.innerWidth - 24) : 280;
        el.style.width = popW + "px";

        // 一時的に表示して高さを計測
        el.style.visibility = "hidden";
        el.style.display = "flex";
        const popH = el.offsetHeight;
        el.style.visibility = "";

        const gap = 8;
        const vw = window.innerWidth;
        const vh = window.innerHeight;

        // 水平: バー中央に合わせ、viewport内にクランプ
        let left = rect.left + (rect.width / 2) - (popW / 2);
        left = Math.max(8, Math.min(left, vw - popW - 8));

        // 垂直: 上に配置(デフォルト)、上にスペースがなければ下にフリップ
        let top;
        if (rect.top - popH - gap > 8) {
            top = rect.top - popH - gap;
        } else {
            top = rect.bottom + gap;
        }
        top = Math.max(8, Math.min(top, vh - popH - 8));

        el.style.left = left + "px";
        el.style.top = top + "px";
    },

    _formatDateJP(dateStr) {
        if (!dateStr) return "";
        const parts = dateStr.split("-");
        return `${parseInt(parts[1])}/${parseInt(parts[2])}`;
    },

    _clearShowTimer() {
        if (this._showTimer) { clearTimeout(this._showTimer); this._showTimer = null; }
    },

    _clearHideTimer() {
        if (this._hideTimer) { clearTimeout(this._hideTimer); this._hideTimer = null; }
    },

    _scheduleHide() {
        this._hideTimer = setTimeout(() => { this.hide(); }, 150);
    },
};

// ---- 全イベントをIDで引けるようにキャッシュ ----
let _eventsById = {};

// ========================================
// メイン描画
// ========================================
function renderTimeline(events, year, holidaySet, options) {
    const container = document.getElementById("timeline-container");

    // バウンス防止
    const prevHeight = container.offsetHeight;
    if (prevHeight > 0) container.style.minHeight = prevHeight + "px";

    _eventsById = {};

    // モードが未初期化 or 年が変わった場合に再構築
    if (_tlYear !== year) setTimelineMode(_tlMode, year);

    const totalCols = _tlTotalCols;
    const today = new Date();
    const todayStr = formatDateYMD(today);
    const todayCol = (today.getFullYear() === year) ? absCol(todayStr) : -1;
    const grouped = groupByCategory(events);

    events.forEach(ev => { if (ev.id) _eventsById[ev.id] = ev; });

    // fitモード: 列幅を動的計算
    // 旧テーブルが残っていると wrapper.clientWidth が膨らむため先にクリアする
    let colWidth = _tlColWidth;
    if (_tlMode === "fit") {
        container.replaceChildren();
        container.style.minHeight = "";
        const wrapper = document.getElementById("timeline-wrapper");
        const fixedWidth = _isMobile() ? 120 : FIXED_COLS_TOTAL;
        // window.innerWidth をフォールバックに使用（より信頼性が高い）
        const wrapperW = wrapper ? Math.min(wrapper.clientWidth, window.innerWidth) : window.innerWidth;
        const available = wrapperW - fixedWidth;
        colWidth = Math.max(10, Math.floor(available / totalCols));
        _tlColWidth = colWidth;
    }

    const table = document.createElement("table");
    table.className = "gs-table";
    if (_tlMode === "fit") table.classList.add("gs-fit-mode");
    table.dataset.year = year;
    table.dataset.mode = _tlMode;
    table.style.setProperty("--col-w", colWidth + "px");
    table.setAttribute("role", "grid");
    table.setAttribute("aria-label", `${year}年 年間スケジュール タイムライン`);

    // ===== thead =====
    const thead = document.createElement("thead");

    // 月ヘッダー行
    const monthRow = document.createElement("tr");
    monthRow.className = "gs-month-row";
    const thCatM = document.createElement("th");
    thCatM.className = "gs-fixed-col gs-col-cat gs-corner";
    thCatM.textContent = "カテゴリ";
    thCatM.rowSpan = 2;
    monthRow.appendChild(thCatM);
    const thEvtM = document.createElement("th");
    thEvtM.className = "gs-fixed-col gs-col-evt gs-corner";
    thEvtM.textContent = "イベント名";
    thEvtM.rowSpan = 2;
    monthRow.appendChild(thEvtM);

    for (let m = 0; m < 12; m++) {
        const colsInMonth = _tlColOffsets[m + 1] - _tlColOffsets[m];
        const th = document.createElement("th");
        th.className = "gs-month-cell";
        th.colSpan = colsInMonth;
        // fitモードでは min-width を設定しない（table-layout:fixed; width:100% に任せる）
        if (_tlMode !== "fit") th.style.minWidth = (colWidth * colsInMonth) + "px";
        th.textContent = MONTH_NAMES[m];
        if (today.getFullYear() === year && today.getMonth() === m) th.classList.add("gs-current-month");
        monthRow.appendChild(th);
    }
    thead.appendChild(monthRow);

    // サブヘッダー行
    const subRow = document.createElement("tr");
    subRow.className = "gs-week-row";

    for (let m = 0; m < 12; m++) {
        const colsInMonth = _tlColOffsets[m + 1] - _tlColOffsets[m];
        for (let li = 0; li < colsInMonth; li++) {
            const globalCol = _tlColOffsets[m] + li;
            const th = document.createElement("th");
            th.className = "gs-week-cell";
            th.textContent = _subHeaderLabel(m, li);

            if (todayCol === globalCol) th.classList.add("gs-current-week");

            // 月末セルに区切りクラスを付与
            if (li === colsInMonth - 1) th.classList.add("gs-month-end");

            // 1日モード: 土日祝の色
            if (_tlMode === "day") {
                const dayNum = li + 1;

                const dateStr = `${year}-${String(m + 1).padStart(2, "0")}-${String(dayNum).padStart(2, "0")}`;
                const d = new Date(year, m, dayNum);
                const dow = d.getDay();
                if (holidaySet && holidaySet.has(dateStr)) {
                    th.classList.add("gs-holiday-col");
                } else if (dow === 0) {
                    th.classList.add("gs-sunday-col");
                } else if (dow === 6) {
                    th.classList.add("gs-saturday-col");
                }
            }

            subRow.appendChild(th);
        }
    }
    thead.appendChild(subRow);
    table.appendChild(thead);

    // ===== tbody =====
    const tbody = document.createElement("tbody");

    CATEGORIES.forEach(cat => {
        const catEvents = grouped[cat.name] || [];
        if (catEvents.length === 0) {
            const tr = createEmptyRow(cat, totalCols, year, holidaySet, todayCol);
            tbody.appendChild(tr);
            return;
        }
        catEvents.forEach((ev, idx) => {
            const tr = createEventRow(ev, cat, idx, catEvents.length, totalCols, year, holidaySet, todayCol);
            tbody.appendChild(tr);
        });
    });

    const other = grouped["__other__"] || [];
    if (other.length > 0) {
        const defCat = { id: "other", name: "その他", color: "#94a3b8", bg: "#f8fafc", border: "#cbd5e1" };
        other.forEach((ev, idx) => {
            const tr = createEventRow(ev, defCat, idx, other.length, totalCols, year, holidaySet, todayCol);
            tbody.appendChild(tr);
        });
    }

    table.appendChild(tbody);
    container.replaceChildren(table);
    container.style.minHeight = "";

    // カテゴリバッジの縦中央揃え（DOM挿入後に実測）
    _centerCatBadges(tbody);

    // fitモード: DOM挿入後に実測列幅を取得してバー位置計算用に更新
    if (_tlMode === "fit") {
        const firstTd = table.querySelector(".gs-week-td");
        if (firstTd) _tlColWidth = firstTd.offsetWidth;
    }

    setupDragHandlers(table, year);

    // ポップオーバー初期化 + スクロールで非表示
    PopoverManager.init();
    const wrapper = document.getElementById("timeline-wrapper");
    if (wrapper && !wrapper._popoverScrollBound) {
        wrapper.addEventListener("scroll", () => { PopoverManager.hide(); }, { passive: true });
        wrapper._popoverScrollBound = true;
    }

    if (!options || !options.skipScroll) {
        scrollToCurrentMonth(year);
    }
}

// カテゴリバッジの縦中央揃え
function _centerCatBadges(tbody) {
    const rows = tbody.querySelectorAll(".gs-event-row");
    let groupRows = [];
    let currentCat = null;

    const adjustGroup = (group) => {
        if (group.length <= 1) return;
        const catCell = group[0].querySelector(".gs-col-cat");
        const badge = catCell && catCell.querySelector(".gs-cat-badge");
        if (!badge) return;

        // グループ全行の合計高さを計算
        let totalH = 0;
        group.forEach(r => totalH += r.offsetHeight);

        const badgeH = badge.offsetHeight;
        const topOffset = (totalH - badgeH) / 2;

        // absolute配置で縦中央に
        badge.style.position = "absolute";
        badge.style.top = topOffset + "px";
        badge.style.left = "50%";
        badge.style.transform = "translateX(-50%)";

        // 先頭セルのz-indexを上げて後続行の上に描画
        const baseZ = parseInt(window.getComputedStyle(catCell).zIndex) || 3;
        catCell.style.zIndex = String(baseZ + 1);
    };

    rows.forEach(row => {
        const cat = row.dataset.categoryName;
        if (cat !== currentCat) {
            if (groupRows.length > 0) adjustGroup(groupRows);
            groupRows = [row];
            currentCat = cat;
        } else {
            groupRows.push(row);
        }
    });
    if (groupRows.length > 0) adjustGroup(groupRows);
}

// サブヘッダーラベル生成
function _subHeaderLabel(month, localIdx) {
    if (_tlMode === "day") {
        return String(localIdx + 1);
    }
    const starts = MODE_STARTS[_tlMode];
    return `${month + 1}/${starts[localIdx]}`;
}

// ---- 空カテゴリ行 ----
function createEmptyRow(cat, totalCols, year, holidaySet, todayCol) {
    const tr = document.createElement("tr");
    tr.className = "gs-event-row gs-cat-first-row gs-cat-last-row";
    tr.dataset.category = cat.id;
    tr.dataset.categoryName = cat.name;

    const tdCat = document.createElement("td");
    tdCat.className = "gs-fixed-col gs-col-cat";
    tdCat.innerHTML = `<span class="gs-cat-badge" style="background:${cat.color};color:#fff">${cat.name}</span>`;
    tr.appendChild(tdCat);

    const tdEvt = document.createElement("td");
    tdEvt.className = "gs-fixed-col gs-col-evt gs-empty-evt";
    tdEvt.setAttribute("role", "button");
    tdEvt.setAttribute("tabindex", "0");
    tdEvt.setAttribute("aria-label", `${cat.name}にイベントを追加`);
    tdEvt.addEventListener("click", () => openEventModal(null, cat.name));
    tdEvt.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); openEventModal(null, cat.name); }
    });

    const evtInner = document.createElement("div");
    evtInner.className = "gs-evt-inner";

    const addBtn = document.createElement("button");
    addBtn.className = "gs-row-add-btn gs-row-add-btn-empty";
    addBtn.textContent = "＋";
    addBtn.title = `${cat.name}にイベントを追加`;
    addBtn.setAttribute("aria-label", `${cat.name}にイベントを追加`);
    addBtn.addEventListener("click", (e) => { e.stopPropagation(); openEventModal(null, cat.name); });
    evtInner.appendChild(addBtn);

    tdEvt.appendChild(evtInner);
    tr.appendChild(tdEvt);

    for (let c = 0; c < totalCols; c++) {
        tr.appendChild(createTimelineCell(c, year, holidaySet, todayCol));
    }
    return tr;
}


// ---- イベント行 ----
// NOTE: カテゴリセルはrowSpanを使わず全行に個別配置。
// rowSpanはborder-collapse:separateでボーダー整合性の問題を起こすため削除。
// バッジは先頭行のみ表示し、CSSで非先頭行のborder-bottomを透明にして
// 視覚的なグループ化を実現。
function createEventRow(ev, cat, idx, totalInCat, totalCols, year, holidaySet, todayCol) {
    const tr = document.createElement("tr");
    tr.className = "gs-event-row";
    tr.dataset.category = cat.id;
    tr.dataset.categoryName = cat.name;
    tr.dataset.eventId = ev.id;

    // 全行に個別のカテゴリセルを配置（rowSpan不使用）
    const tdCat = document.createElement("td");
    tdCat.className = "gs-fixed-col gs-col-cat";
    if (idx === 0) {
        tr.classList.add("gs-cat-first-row");
        tdCat.innerHTML = `<span class="gs-cat-badge" style="background:${cat.color};color:#fff">${cat.name}</span>`;
    }
    if (idx === totalInCat - 1) {
        tr.classList.add("gs-cat-last-row");
    }
    tr.appendChild(tdCat);

    const tdEvt = document.createElement("td");
    tdEvt.className = "gs-fixed-col gs-col-evt";

    const evtInner = document.createElement("div");
    evtInner.className = "gs-evt-inner";

    const evtLabel = document.createElement("span");
    evtLabel.className = "gs-evt-label";
    evtLabel.textContent = ev.title;
    evtLabel.title = ev.title;
    evtLabel.setAttribute("role", "button");
    evtLabel.setAttribute("tabindex", "0");
    evtLabel.setAttribute("aria-label", `${ev.title}を編集（${ev.startDate}〜${ev.endDate}）`);
    evtLabel.addEventListener("click", (e) => { e.stopPropagation(); openEventModal(ev, cat.name); });
    evtLabel.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); openEventModal(ev, cat.name); }
    });
    evtInner.appendChild(evtLabel);

    const addBtn = document.createElement("button");
    addBtn.className = "gs-row-add-btn";
    addBtn.textContent = "＋";
    addBtn.title = `${cat.name}にイベントを追加`;
    addBtn.setAttribute("aria-label", `${cat.name}にイベントを追加`);
    addBtn.addEventListener("click", (e) => { e.stopPropagation(); openEventModal(null, cat.name); });
    evtInner.appendChild(addBtn);

    tdEvt.appendChild(evtInner);
    tr.appendChild(tdEvt);

    const startCol = absCol(ev.startDate);
    const endCol = absCol(ev.endDate);

    for (let c = 0; c < totalCols; c++) {
        const td = createTimelineCell(c, year, holidaySet, todayCol);
        if (c >= startCol && c <= endCol) {
            td.classList.add("gs-bar-cell");
            td.dataset.bar = "true";
            td.dataset.eventId = ev.id;
            td.style.backgroundColor = cat.bg;
            td.style.borderTop = `1px solid ${cat.border}`;
            td.style.borderBottom = `1px solid ${cat.border}`;
            if (c === startCol) {
                td.dataset.barEdge = "start";
                td.style.borderLeft = `var(--bar-accent-width) solid ${cat.color}`;
                td.style.borderTopLeftRadius = "var(--bar-radius)";
                td.style.borderBottomLeftRadius = "var(--bar-radius)";
                const handle = document.createElement("div");
                handle.className = "gs-resize-handle gs-resize-handle-left";
                td.appendChild(handle);
            }
            if (c === endCol) {
                td.dataset.barEdge = (c === startCol) ? "both" : "end";
                td.style.borderRight = `2px solid ${cat.border}`;
                td.style.borderTopRightRadius = "var(--bar-radius)";
                td.style.borderBottomRightRadius = "var(--bar-radius)";
                const handle = document.createElement("div");
                handle.className = "gs-resize-handle gs-resize-handle-right";
                td.appendChild(handle);
            }

            // ポップオーバー: hover リスナー
            td.addEventListener("mouseenter", () => { PopoverManager.scheduleShow(ev.id, td); });
            td.addEventListener("mouseleave", () => { PopoverManager.scheduleHide(); });

            // ポップオーバー: タッチタップ リスナー
            ((barTd, evId) => {
                let _touchStartPos = null;
                barTd.addEventListener("touchstart", (te) => {
                    const t = te.touches[0];
                    _touchStartPos = { x: t.clientX, y: t.clientY };
                }, { passive: true });
                barTd.addEventListener("touchend", (te) => {
                    if (!_touchStartPos || DragManager.started) { _touchStartPos = null; return; }
                    const t = te.changedTouches[0];
                    const dx = t.clientX - _touchStartPos.x;
                    const dy = t.clientY - _touchStartPos.y;
                    _touchStartPos = null;
                    // タップ判定: 移動距離 < 10px
                    if (dx * dx + dy * dy < 100) {
                        // ドラッグ中でなければポップオーバーをトグル
                        if (!DragManager.active) {
                            if (PopoverManager._currentEventId === evId &&
                                PopoverManager._el && PopoverManager._el.classList.contains("gs-popover-visible")) {
                                PopoverManager.hide();
                            } else {
                                PopoverManager._show(evId, barTd);
                            }
                        }
                    }
                }, { passive: true });
            })(td, ev.id);
        }
        td.dataset.col = c;
        td.dataset.startCol = startCol;
        td.dataset.endCol = endCol;
        tr.appendChild(td);
    }
    return tr;
}

// ---- タイムラインセル ----
function createTimelineCell(colIdx, year, holidaySet, todayCol) {
    const td = document.createElement("td");
    td.className = "gs-week-td";
    td.dataset.col = colIdx;

    if (todayCol === colIdx) td.classList.add("gs-today-col");

    // 月末判定: この列が月の最後の列かどうか
    const month = _colToMonth(colIdx);
    if (colIdx === _tlColOffsets[month + 1] - 1) td.classList.add("gs-month-end");

    // 土日・祝日の背景色（1日モード時のみ正確に判定）
    if (_tlMode === "day") {
        const localIdx = colIdx - _tlColOffsets[month];
        const dayNum = localIdx + 1;
        const dateStr = `${year}-${String(month + 1).padStart(2, "0")}-${String(dayNum).padStart(2, "0")}`;
        const d = new Date(year, month, dayNum);
        const dow = d.getDay(); // 0=日, 6=土
        if (holidaySet && holidaySet.has(dateStr)) {
            td.classList.add("gs-holiday-col");
        } else if (dow === 0) {
            td.classList.add("gs-sunday-col");
        } else if (dow === 6) {
            td.classList.add("gs-saturday-col");
        }
    }

    return td;
}

// ========================================
// ドラッグハンドラ設置
// ========================================
let _dragDocListenersAttached = false;
let _currentDragTable = null;

function setupDragHandlers(table, year) {
    _currentDragTable = table;

    // --- マウスイベント ---
    table.addEventListener("mousedown", (e) => {
        const td = e.target.closest("td.gs-week-td");
        const handle = e.target.closest(".gs-resize-handle");
        if (!td) return;

        const tr = td.closest("tr.gs-event-row");
        if (!tr) return;

        const col = parseInt(td.dataset.col, 10);
        const eventId = tr.dataset.eventId || null;
        const catName = tr.dataset.categoryName || null;
        const evObj = eventId ? _eventsById[eventId] : null;

        let barStart = -1, barEnd = -1;
        if (eventId && evObj) {
            barStart = absCol(evObj.startDate);
            barEnd = absCol(evObj.endDate);
        }

        if (handle) {
            e.preventDefault();
            const isLeft = handle.classList.contains("gs-resize-handle-left");
            DragManager.begin(
                isLeft ? "resize-start" : "resize-end",
                tr, col, e.clientX, e.clientY,
                eventId, evObj, catName, barStart, barEnd, year
            );
            return;
        }

        if (td.dataset.bar === "true" && eventId) {
            e.preventDefault();
            DragManager.begin(
                "move", tr, col, e.clientX, e.clientY,
                eventId, evObj, catName, barStart, barEnd, year
            );
        } else if (!td.dataset.bar && catName) {
            e.preventDefault();
            DragManager.begin(
                "create", tr, col, e.clientX, e.clientY,
                null, null, catName, -1, -1, year
            );
        }
    });

    // --- タッチイベント ---
    table.addEventListener("touchstart", (e) => {
        const t = e.touches[0];
        const td = document.elementFromPoint(t.clientX, t.clientY);
        if (!td || !td.closest) return;
        const cell = td.closest("td.gs-week-td");
        const handle = td.closest(".gs-resize-handle");
        if (!cell) return;

        const tr = cell.closest("tr.gs-event-row");
        if (!tr) return;

        const col = parseInt(cell.dataset.col, 10);
        const eventId = tr.dataset.eventId || null;
        const catName = tr.dataset.categoryName || null;
        const evObj = eventId ? _eventsById[eventId] : null;

        let barStart = -1, barEnd = -1;
        if (eventId && evObj) {
            barStart = absCol(evObj.startDate);
            barEnd = absCol(evObj.endDate);
        }

        if (handle) {
            e.preventDefault(); // リサイズ時はスクロール抑制
            const isLeft = handle.classList.contains("gs-resize-handle-left");
            DragManager.begin(
                isLeft ? "resize-start" : "resize-end",
                tr, col, t.clientX, t.clientY,
                eventId, evObj, catName, barStart, barEnd, year
            );
            return;
        }

        if (cell.dataset.bar === "true" && eventId) {
            // バー移動: まだpreventDefaultしない（スクロールの可能性）
            DragManager.begin(
                "move", tr, col, t.clientX, t.clientY,
                eventId, evObj, catName, barStart, barEnd, year
            );
        } else if (!cell.dataset.bar && catName) {
            // 新規作成: まだpreventDefaultしない
            DragManager.begin(
                "create", tr, col, t.clientX, t.clientY,
                null, null, catName, -1, -1, year
            );
        }
    }, { passive: false });

    if (_dragDocListenersAttached) return;
    _dragDocListenersAttached = true;

    let _cachedWrapper = null;
    const getWrapper = () => _cachedWrapper || (_cachedWrapper = document.getElementById("timeline-wrapper"));

    document.addEventListener("mousemove", (e) => {
        if (!DragManager.active) return;
        DragManager.move(e.clientX, e.clientY);
        const wrapper = getWrapper();
        if (wrapper) {
            const wr = wrapper.getBoundingClientRect();
            if (e.clientX > wr.right - 40) wrapper.scrollLeft += 10;
            else if (e.clientX < wr.left + 40 + FIXED_COLS_TOTAL) wrapper.scrollLeft -= 10;
        }
    }, { passive: true });

    document.addEventListener("mouseup", (e) => {
        if (!DragManager.active) return;
        const result = DragManager.end();
        if (!result) return;

        if (result.action === "create") {
            onDragCreate(result.categoryName, result.startDate, result.endDate);
        } else if (result.action === "move" || result.action === "resize") {
            onDragMoveOrResize(result.eventId, result.eventObj, result.startDate, result.endDate);
        }
    });

    document.addEventListener("keydown", (e) => {
        if (e.key === "Escape" && DragManager.active) {
            DragManager.cancel();
        }
    });

    // --- タッチ: document レベルリスナー ---
    document.addEventListener("touchmove", (e) => {
        if (!DragManager.active) return;
        const t = e.touches[0];
        DragManager.move(t.clientX, t.clientY);
        // ドラッグ開始済みならスクロール抑制
        if (DragManager.started) {
            e.preventDefault();
            const wrapper = getWrapper();
            if (wrapper) {
                const wr = wrapper.getBoundingClientRect();
                const fixedW = _isMobile() ? 120 : FIXED_COLS_TOTAL;
                if (t.clientX > wr.right - 40) wrapper.scrollLeft += 10;
                else if (t.clientX < wr.left + 40 + fixedW) wrapper.scrollLeft -= 10;
            }
        }
    }, { passive: false });

    document.addEventListener("touchend", (e) => {
        if (!DragManager.active) return;
        const result = DragManager.end();
        if (!result) return;

        if (result.action === "create") {
            onDragCreate(result.categoryName, result.startDate, result.endDate);
        } else if (result.action === "move" || result.action === "resize") {
            onDragMoveOrResize(result.eventId, result.eventObj, result.startDate, result.endDate);
        }
    });

    document.addEventListener("touchcancel", () => {
        if (DragManager.active) DragManager.cancel();
    });
}

// ---- 凡例 ----
function renderLegend() {
    const legend = document.getElementById("legend");
    legend.innerHTML = CATEGORIES.map(cat =>
        `<span class="legend-item">
            <span class="legend-color" style="background:${cat.color}" aria-hidden="true"></span>
            ${cat.name}
        </span>`
    ).join("");
}

// ---- スクロール ----
function scrollToCurrentMonth(year) {
    const wrapper = document.getElementById("timeline-wrapper");
    if (!wrapper) return;
    const now = new Date();
    if (now.getFullYear() !== year) { wrapper.scrollLeft = 0; return; }
    if (_tlMode === "fit") { wrapper.scrollLeft = 0; return; }
    const scrollTarget = _tlColOffsets[now.getMonth()] * _tlColWidth;
    wrapper.scrollLeft = Math.max(0, scrollTarget - 50);
}

function scrollToToday() {
    scrollToCurrentMonth(graphConfig.year);
}

