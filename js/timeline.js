// Excel風タイムライン描画エンジン（ドラッグ操作対応・マルチモード版）
// 表示モード: 週(week) / 5日(fiveday) / 1日(day) / 全体(fit)

// ========================================
// 定数
// ========================================
const ROW_HEIGHT = 36;
const FIXED_COL_WIDTH_CAT = 120;
const FIXED_COL_WIDTH_EVT = 250;
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
const MODE_COL_WIDTHS = { week: 80, fiveday: 48, day: 28, fit: null };

// ========================================
// タイムラインモード状態
// ========================================
let _tlMode = "day";
let _tlColOffsets = [0]; // 初期化時に再構築
let _tlTotalCols = 365;
let _tlColWidth = 24;
let _tlYear = null; // 初回描画でsetTimelineModeを確実に呼ぶ
// 週モード用: 各月の月曜日開始日リスト（setTimelineModeで構築）
let _weekMondayStarts = null; // [month] => [day1, day2, ...]

// 月内の月曜日の日付リストを返す
function _getMondaysInMonth(year, month) {
    const days = daysInMonth(year, month);
    const mondays = [];
    for (let d = 1; d <= days; d++) {
        if (new Date(year, month, d).getDay() === 1) mondays.push(d);
    }
    return mondays;
}

// モードを設定し、列オフセットテーブルを再構築
function setTimelineMode(modeId, year) {
    _tlMode = modeId;
    _tlYear = year;

    // 週/全体モード: 各月の実際の月曜日を計算
    if (modeId === "week" || modeId === "fit") {
        _weekMondayStarts = [];
        for (let m = 0; m < 12; m++) {
            _weekMondayStarts.push(_getMondaysInMonth(year, m));
        }
    } else {
        _weekMondayStarts = null;
    }

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
    if ((_tlMode === "week" || _tlMode === "fit") && _weekMondayStarts) return _weekMondayStarts[month].length;
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

    // 週モード: 実際の月曜日リストから該当週を特定
    if ((_tlMode === "week" || _tlMode === "fit") && _weekMondayStarts) {
        const mondays = _weekMondayStarts[month];
        let wi = 0;
        for (let i = mondays.length - 1; i >= 0; i--) {
            if (day >= mondays[i]) { wi = i; break; }
        }
        return _tlColOffsets[month] + wi;
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

    if ((_tlMode === "week" || _tlMode === "fit") && _weekMondayStarts) {
        const day = _weekMondayStarts[month][localIdx];
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
        return colToStartDate(colIdx, year);
    }

    if ((_tlMode === "week" || _tlMode === "fit") && _weekMondayStarts) {
        const mondays = _weekMondayStarts[month];
        let lastDay;
        if (localIdx < mondays.length - 1) {
            lastDay = mondays[localIdx + 1] - 1;
        } else {
            lastDay = daysInMonth(year, month);
        }
        return `${year}-${String(month + 1).padStart(2, "0")}-${String(lastDay).padStart(2, "0")}`;
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
// タイトルグループ化（同タイトルのイベントを1行にまとめる）
// ========================================
// 親タイトルを取得（"イベント名_サブイベント名" → "イベント名"）
function _getParentTitle(title) {
    if (!title) return "(無題)";
    const idx = title.indexOf("_");
    return idx >= 0 ? title.substring(0, idx) : title;
}

// サブイベント名を取得（"イベント名_サブイベント名" → "サブイベント名"、なければnull）
function _getSubEventName(title) {
    if (!title) return null;
    const idx = title.indexOf("_");
    return idx >= 0 ? title.substring(idx + 1) : null;
}

function _groupByTitle(events) {
    const groups = {};
    const order = [];
    events.forEach(ev => {
        const key = _getParentTitle(ev.title);
        if (!groups[key]) { groups[key] = []; order.push(key); }
        groups[key].push(ev);
    });
    return { groups, order };
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
    eventTitle: null,
    originalStartCol: -1,
    originalEndCol: -1,
    year: null,

    _cells: null,
    _highlighted: [],
    _rafId: 0,
    _colGeometry: null,

    begin(mode, row, col, mouseX, mouseY, eventId, eventObj, catName, startCol, endCol, year, eventTitle) {
        PopoverManager.hide();
        document.querySelectorAll(".gs-bar-hover").forEach(c => c.classList.remove("gs-bar-hover"));
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
        this.eventTitle = eventTitle || null;
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
                eventTitle: this.eventTitle,
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
        // タイトルをそのまま表示（「イベント名_サブ名」形式）
        el.querySelector(".gs-popover-title").textContent = ev.title;

        const badge = el.querySelector(".gs-popover-badge");
        badge.textContent = cat.name;
        badge.style.background = cat.color;

        const startFmt = this._formatDateJP(ev.startDate);
        const endFmt = this._formatDateJP(ev.endDate);
        el.querySelector(".gs-popover-dates").textContent =
            ev.startDate === ev.endDate ? startFmt : `${startFmt} 〜 ${endFmt}`;

        el.querySelector(".gs-popover-notes").textContent = "";

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
        if (m % 2 === 1) th.classList.add("gs-month-even");
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
            if (m % 2 === 1) th.classList.add("gs-month-even");
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

    // マーカー表示対象のカテゴリID
    const MARKER_CATEGORIES = new Set(["morning_meeting"]);
    const PINNED_ORDER = ["morning_meeting", "gyro_holiday"];
    const pinnedSet = new Set(PINNED_ORDER);
    const _markerExceptions = _loadMarkerExceptions(year);

    // ---- カテゴリ行を描画するヘルパー ----
    const _renderCatRows = (cat) => {
        if (MARKER_CATEGORIES.has(cat.id)) {
            const markerEvents = grouped[cat.name] || [];
            const tr = createMarkerRow(cat, totalCols, year, holidaySet, todayCol, _markerExceptions, markerEvents);
            tbody.appendChild(tr);
            return;
        }
        const catEvents = grouped[cat.name] || [];
        if (catEvents.length === 0) {
            const tr = createEmptyRow(cat, totalCols, year, holidaySet, todayCol);
            tbody.appendChild(tr);
            return;
        }
        const { groups: titleGroups, order: titleOrder } = _groupByTitle(catEvents);
        // カスタム順序があれば適用
        const customOrder = _rawConfig?.titleOrders?.[String(year)]?.[cat.name];
        if (customOrder) {
            titleOrder.sort((a, b) => {
                const ia = customOrder.indexOf(a);
                const ib = customOrder.indexOf(b);
                if (ia === -1 && ib === -1) return 0;
                if (ia === -1) return 1;
                if (ib === -1) return -1;
                return ia - ib;
            });
        }
        titleOrder.forEach((title, idx) => {
            const evGroup = titleGroups[title];
            const tr = createEventRow(evGroup, cat, idx, titleOrder.length, totalCols, year, holidaySet, todayCol);
            tbody.appendChild(tr);
        });
    };

    // 1パス目: 固定カテゴリを指定順で先頭に描画
    PINNED_ORDER.forEach(pinId => {
        const cat = CATEGORIES.find(c => c.id === pinId);
        if (cat) _renderCatRows(cat);
    });

    // 2パス目: 通常カテゴリ
    CATEGORIES.forEach(cat => {
        if (pinnedSet.has(cat.id)) return;
        _renderCatRows(cat);
    });

    const other = grouped["__other__"] || [];
    if (other.length > 0) {
        const defCat = { id: "other", name: "その他", color: "#94a3b8", bg: "#f8fafc", border: "#cbd5e1" };
        const { groups: otherGroups, order: otherOrder } = _groupByTitle(other);
        otherOrder.forEach((title, idx) => {
            const evGroup = otherGroups[title];
            const tr = createEventRow(evGroup, defCat, idx, otherOrder.length, totalCols, year, holidaySet, todayCol);
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
    if ((_tlMode === "week" || _tlMode === "fit") && _weekMondayStarts) {
        return `${month + 1}/${_weekMondayStarts[month][localIdx]}`;
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
    const evtInner = document.createElement("div");
    evtInner.className = "gs-evt-inner";

    if (!!getActiveAccount()) {
        tdEvt.setAttribute("role", "button");
        tdEvt.setAttribute("tabindex", "0");
        tdEvt.setAttribute("aria-label", `${cat.name}にイベントを追加`);
        tdEvt.addEventListener("click", () => openEventModal(null, cat.name));
        tdEvt.addEventListener("keydown", (e) => {
            if (e.key === "Enter" || e.key === " ") { e.preventDefault(); openEventModal(null, cat.name); }
        });

        const addBtn = document.createElement("button");
        addBtn.className = "gs-row-add-btn gs-row-add-btn-empty";
        addBtn.textContent = "＋";
        addBtn.title = `${cat.name}にイベントを追加`;
        addBtn.setAttribute("aria-label", `${cat.name}にイベントを追加`);
        addBtn.addEventListener("click", (e) => { e.stopPropagation(); openEventModal(null, cat.name); });
        evtInner.appendChild(addBtn);
    }

    tdEvt.appendChild(evtInner);
    tr.appendChild(tdEvt);

    for (let c = 0; c < totalCols; c++) {
        tr.appendChild(createTimelineCell(c, year, holidaySet, todayCol));
    }
    return tr;
}

// ---- マーカー行（休日・朝会など、日ごとにマーカーを表示する行） ----
function createMarkerRow(cat, totalCols, year, holidaySet, todayCol, exceptions, catEvents) {
    const tr = document.createElement("tr");
    tr.className = "gs-event-row gs-cat-first-row gs-cat-last-row gs-marker-row";
    tr.dataset.category = cat.id;
    tr.dataset.categoryName = cat.name;

    // カテゴリセル
    const tdCat = document.createElement("td");
    tdCat.className = "gs-fixed-col gs-col-cat";
    tdCat.innerHTML = `<span class="gs-cat-badge" style="background:${cat.color};color:#fff">${cat.name}</span>`;
    tr.appendChild(tdCat);

    // イベント名セル（マーカー行は固定ラベル）
    const tdEvt = document.createElement("td");
    tdEvt.className = "gs-fixed-col gs-col-evt";
    const evtInner = document.createElement("div");
    evtInner.className = "gs-evt-inner";
    const evtLabel = document.createElement("span");
    evtLabel.className = "gs-evt-label gs-marker-label";
    evtLabel.textContent = cat.name;
    evtInner.appendChild(evtLabel);
    tdEvt.appendChild(evtInner);
    tr.appendChild(tdEvt);

    // マーカー判定関数を取得（候補日の判定用）
    const shouldMark = _getMarkerPredicate(cat, year, holidaySet, exceptions);
    // イベントがある日を高速検索用にSetで持つ
    const eventDates = new Set();
    (catEvents || []).forEach(e => {
        let d = new Date(e.startDate + "T00:00:00");
        const end = new Date(e.endDate + "T00:00:00");
        while (d <= end) {
            eventDates.add(formatDateYMD(d));
            d.setDate(d.getDate() + 1);
        }
    });

    const isSignedIn = !!getActiveAccount();

    for (let c = 0; c < totalCols; c++) {
        const td = createTimelineCell(c, year, holidaySet, todayCol);
        const result = shouldMark(c);
        if (result) {
            td.classList.add("gs-marker-cell");
            const marker = document.createElement("span");
            marker.className = "gs-marker";

            const isHoliday = holidaySet && holidaySet.has(result.dateStr);
            const hasEvent = eventDates.has(result.dateStr);

            let markerActive;
            if (isHoliday) {
                // 祝日 → 常に✕
                td.classList.add("gs-marker-excluded");
                marker.textContent = "✕";
                marker.style.color = "#94a3b8";
                markerActive = false;
            } else if (isSignedIn) {
                // サインイン済み → Outlookイベントの有無で決定
                if (hasEvent) {
                    marker.textContent = "●";
                    marker.style.color = cat.color;
                    markerActive = true;
                } else {
                    td.classList.add("gs-marker-excluded");
                    marker.textContent = "✕";
                    marker.style.color = "#94a3b8";
                    markerActive = false;
                }
            } else {
                // 未サインイン → デフォルト●
                marker.textContent = "●";
                marker.style.color = cat.color;
                markerActive = true;
            }
            td.appendChild(marker);

            // クリックで簡易モーダルを開く
            td.style.cursor = "pointer";
            td.addEventListener("click", () => {
                openSimpleEventModal(result.dateStr, cat.name, markerActive);
            });
        }
        tr.appendChild(td);
    }
    return tr;
}

// マーカー判定関数を返す
// 戻り値: null（マーカーなし）| "active"（通常マーカー）| "excluded"（除外日）
// dateStr も返す: { status, dateStr }
function _getMarkerPredicate(cat, year, holidaySet, exceptions) {
    const exceptSet = new Set((exceptions && exceptions[cat.id]) || []);

    if (cat.id === "holiday") {
        // 休日: holidaySetに含まれる日にマーカー
        return (colIdx) => {
            if (_tlMode === "day") {
                const month = _colToMonth(colIdx);
                const localIdx = colIdx - _tlColOffsets[month];
                const dayNum = localIdx + 1;
                const dateStr = `${year}-${String(month + 1).padStart(2, "0")}-${String(dayNum).padStart(2, "0")}`;
                if (!(holidaySet && holidaySet.has(dateStr))) return null;
                return { status: exceptSet.has(dateStr) ? "excluded" : "active", dateStr };
            }
            const start = colToStartDate(colIdx, year);
            const end = colToEndDate(colIdx, year);
            if (!holidaySet) return null;
            const s = parseDateStr(start);
            const e = parseDateStr(end);
            for (let d = new Date(s); d <= e; d.setDate(d.getDate() + 1)) {
                if (holidaySet.has(formatDateYMD(d))) {
                    const ds = formatDateYMD(d);
                    return { status: exceptSet.has(ds) ? "excluded" : "active", dateStr: ds };
                }
            }
            return null;
        };
    }
    if (cat.id === "morning_meeting") {
        // 朝会: 毎週月曜日にマーカー（除外日チェック付き）
        return (colIdx) => {
            if (_tlMode === "day") {
                const month = _colToMonth(colIdx);
                const localIdx = colIdx - _tlColOffsets[month];
                const dayNum = localIdx + 1;
                const d = new Date(year, month, dayNum);
                if (d.getDay() !== 1) return null;
                const dateStr = formatDateYMD(d);
                return { status: exceptSet.has(dateStr) ? "excluded" : "active", dateStr };
            }
            const start = colToStartDate(colIdx, year);
            const end = colToEndDate(colIdx, year);
            const s = parseDateStr(start);
            const e = parseDateStr(end);
            for (let d = new Date(s); d <= e; d.setDate(d.getDate() + 1)) {
                if (d.getDay() === 1) {
                    const ds = formatDateYMD(d);
                    return { status: exceptSet.has(ds) ? "excluded" : "active", dateStr: ds };
                }
            }
            return null;
        };
    }
    // デフォルト: マーカーなし
    return () => null;
}

// ---- マーカー除外日の永続化 ----
function _loadMarkerExceptions(year) {
    try {
        const raw = localStorage.getItem(`gyro_marker_exceptions_${year}`);
        return raw ? JSON.parse(raw) : {};
    } catch { return {}; }
}
function _saveMarkerExceptions(year, exceptions) {
    try {
        localStorage.setItem(`gyro_marker_exceptions_${year}`, JSON.stringify(exceptions));
    } catch (e) { console.warn("除外日の保存に失敗:", e); }
}
function _toggleMarkerException(catId, dateStr, year) {
    const exceptions = _loadMarkerExceptions(year);
    if (!exceptions[catId]) exceptions[catId] = [];
    const idx = exceptions[catId].indexOf(dateStr);
    if (idx >= 0) {
        exceptions[catId].splice(idx, 1);
    } else {
        exceptions[catId].push(dateStr);
    }
    _saveMarkerExceptions(year, exceptions);
    rerenderFromCache();
}

// 除外リストから指定日を確実に削除（トグルではなく一方向）
function _removeMarkerException(catId, dateStr, year) {
    const exceptions = _loadMarkerExceptions(year);
    if (!exceptions[catId]) return;
    const idx = exceptions[catId].indexOf(dateStr);
    if (idx >= 0) {
        exceptions[catId].splice(idx, 1);
        _saveMarkerExceptions(year, exceptions);
    }
}

// 除外リストに指定日を確実に追加（トグルではなく一方向）
function _addMarkerException(catId, dateStr, year) {
    const exceptions = _loadMarkerExceptions(year);
    if (!exceptions[catId]) exceptions[catId] = [];
    if (!exceptions[catId].includes(dateStr)) {
        exceptions[catId].push(dateStr);
        _saveMarkerExceptions(year, exceptions);
    }
}

// ---- イベント行 ----
// NOTE: カテゴリセルはrowSpanを使わず全行に個別配置。
// rowSpanはborder-collapse:separateでボーダー整合性の問題を起こすため削除。
// バッジは先頭行のみ表示し、CSSで非先頭行のborder-bottomを透明にして
// 視覚的なグループ化を実現。
// ---- イベント行の▲▼並べ替え ----
async function _moveRow(catName, parentTitle, direction, year) {
    const yearStr = String(year);
    if (!_rawConfig.titleOrders) _rawConfig.titleOrders = {};
    if (!_rawConfig.titleOrders[yearStr]) _rawConfig.titleOrders[yearStr] = {};

    // カテゴリ内の全タイトルを現在の順序で取得
    const catEvents = (_cachedGraphEvents || []).filter(e =>
        e.categories?.includes(catName) && e.startDate?.startsWith(yearStr)
    );
    const { order: currentOrder } = _groupByTitle(catEvents);

    // カスタム順序があれば使用、なければ現在順序
    let order = _rawConfig.titleOrders[yearStr][catName];
    if (!order || order.length === 0) order = [...currentOrder];

    const idx = order.indexOf(parentTitle);
    if (idx < 0) return;
    const newIdx = idx + direction;
    if (newIdx < 0 || newIdx >= order.length) return;

    // スワップ
    [order[idx], order[newIdx]] = [order[newIdx], order[idx]];
    _rawConfig.titleOrders[yearStr][catName] = order;

    // 再描画（バッジ位置も正しくなる）
    rerenderFromCache();

    // バックグラウンドで保存
    try {
        const token = await getAccessToken();
        await saveCategoryConfig(token, _rawConfig, _configEventId);
    } catch (err) {
        console.warn("[行並べ替え] 保存失敗:", err.message);
    }
}

function createEventRow(evGroup, cat, idx, totalInCat, totalCols, year, holidaySet, todayCol) {
    // evGroup: 同タイトルのイベント配列
    const primaryEv = evGroup[0];
    const tr = document.createElement("tr");
    tr.className = "gs-event-row";
    tr.dataset.category = cat.id;
    tr.dataset.categoryName = cat.name;
    tr.dataset.eventIds = evGroup.map(e => e.id).join(",");

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

    const parentTitle = _getParentTitle(primaryEv.title);
    const isSignedIn = !!getActiveAccount();

    // ▲▼ボタン（サインイン時 + 2行以上）またはスペーサー（テキスト揃え用）
    if (totalInCat > 1 && isSignedIn) {
        const btnWrap = document.createElement("span");
        btnWrap.style.cssText = "display:inline-flex;flex-direction:column;gap:0;margin-right:2px;width:14px;";

        const upBtn = document.createElement("button");
        upBtn.textContent = "▲";
        upBtn.className = "gs-row-move-btn";
        upBtn.style.cssText = "font-size:8px;color:#94a3b8;background:none;border:none;cursor:pointer;padding:0 2px;line-height:1;";
        if (idx === 0) { upBtn.disabled = true; upBtn.style.opacity = "0.3"; }
        upBtn.addEventListener("click", (e) => { e.stopPropagation(); _moveRow(cat.name, parentTitle, -1, year); });

        const downBtn = document.createElement("button");
        downBtn.textContent = "▼";
        downBtn.className = "gs-row-move-btn";
        downBtn.style.cssText = "font-size:8px;color:#94a3b8;background:none;border:none;cursor:pointer;padding:0 2px;line-height:1;";
        if (idx === totalInCat - 1) { downBtn.disabled = true; downBtn.style.opacity = "0.3"; }
        downBtn.addEventListener("click", (e) => { e.stopPropagation(); _moveRow(cat.name, parentTitle, 1, year); });

        btnWrap.appendChild(upBtn);
        btnWrap.appendChild(downBtn);
        evtInner.appendChild(btnWrap);
    } else {
        // スペーサー（▲▼ボタンと同じ幅でテキスト位置を揃える）
        const spacer = document.createElement("span");
        spacer.style.cssText = "display:inline-block;width:14px;flex-shrink:0;";
        evtInner.appendChild(spacer);
    }

    const evtLabel = document.createElement("span");
    evtLabel.className = "gs-evt-label";
    evtLabel.textContent = parentTitle;
    evtLabel.title = parentTitle;

    // サインイン時のみインライン編集を有効化
    if (isSignedIn) {
        evtLabel.setAttribute("role", "button");
        evtLabel.setAttribute("tabindex", "0");
        evtLabel.setAttribute("aria-label", `${parentTitle}の名称を変更`);
        evtLabel.style.cursor = "pointer";
    }

    // クリックでインライン編集（サインイン時のみ）
    if (isSignedIn) evtLabel.addEventListener("click", (e) => {
        e.stopPropagation();
        if (evtLabel.querySelector("input")) return; // 既に編集中

        const input = document.createElement("input");
        input.type = "text";
        input.value = parentTitle;
        input.className = "gs-evt-rename-input";
        input.style.cssText = "width:100%;font:inherit;padding:2px 4px;border:1px solid #f59e0b;border-radius:4px;outline:none;background:#fff;color:#1e293b;";

        const oldText = evtLabel.textContent;
        evtLabel.textContent = "";
        evtLabel.appendChild(input);
        input.focus();
        input.select();

        const save = async () => {
            const newName = input.value.trim();
            if (!newName || newName === parentTitle) {
                evtLabel.textContent = oldText;
                return;
            }
            evtLabel.textContent = newName + "...";

            try {
                const token = await getAccessToken();
                // 同じ親タイトルの全イベントを更新
                for (const ev of evGroup) {
                    const subName = _getSubEventName(ev.title);
                    const newTitle = subName ? newName + "_" + subName : newName;
                    await updateCalendarEvent(token, ev.id, { title: newTitle, category: cat.name, startDate: ev.startDate, endDate: ev.endDate, notes: ev.bodyPreview || "" });
                    ev.title = newTitle;
                }
                rerenderFromCache();
                announceStatus(`「${parentTitle}」→「${newName}」に変更しました`);
                publishEventsToGitHub(_cachedGraphEvents, CATEGORIES, currentYear).catch(e => console.warn("[GitHub公開]", e.message));
            } catch (err) {
                console.error("Rename failed:", err);
                evtLabel.textContent = oldText;
                announceStatus("名称変更に失敗しました: " + err.message);
            }
        };

        input.addEventListener("blur", save);
        input.addEventListener("keydown", (ke) => {
            if (ke.key === "Enter") { ke.preventDefault(); input.blur(); }
            if (ke.key === "Escape") { evtLabel.textContent = oldText; }
        });
    });
    evtInner.appendChild(evtLabel);

    if (isSignedIn) {
        const addBtn = document.createElement("button");
        addBtn.className = "gs-row-add-btn";
        addBtn.textContent = "＋";
        addBtn.title = `${cat.name}にイベントを追加`;
        addBtn.setAttribute("aria-label", `${cat.name}にイベントを追加`);
        const rowParentTitle = _getParentTitle(evGroup[0].title);
        addBtn.addEventListener("click", (e) => { e.stopPropagation(); openEventModal(null, cat.name, null, null, rowParentTitle); });
        evtInner.appendChild(addBtn);
    }

    tdEvt.appendChild(evtInner);
    tr.appendChild(tdEvt);

    // 各イベントの startCol/endCol を事前計算
    const ranges = evGroup.map(ev => ({
        ev,
        startCol: absCol(ev.startDate),
        endCol:   absCol(ev.endDate),
    }));

    for (let c = 0; c < totalCols; c++) {
        const td = createTimelineCell(c, year, holidaySet, todayCol);
        // このカラムにバーがあるイベントを探す
        const hit = ranges.find(r => c >= r.startCol && c <= r.endCol);
        if (hit) {
            const ev = hit.ev;
            td.classList.add("gs-bar-cell");
            td.dataset.bar = "true";
            td.dataset.eventId = ev.id;
            td.style.setProperty("--bar-bg", cat.bg);
            td.style.setProperty("--bar-border", cat.border);
            td.style.setProperty("--bar-color", cat.color);
            if (c === hit.startCol) {
                td.dataset.barEdge = "start";
            }
            if (c === hit.endCol) {
                td.dataset.barEdge = (c === hit.startCol) ? "both" : "end";
            }
            // バー内部要素
            const barInner = document.createElement("div");
            barInner.className = "gs-bar-inner";
            // リサイズハンドル
            if (c === hit.startCol) {
                const handle = document.createElement("div");
                handle.className = "gs-resize-handle gs-resize-handle-left";
                barInner.appendChild(handle);
            }
            if (c === hit.endCol) {
                const handle = document.createElement("div");
                handle.className = "gs-resize-handle gs-resize-handle-right";
                barInner.appendChild(handle);
            }
            // サブイベントラベル: バーの開始セルのみに表示
            const subName = _getSubEventName(ev.title);
            if (subName && c === hit.startCol) {
                const label = document.createElement("span");
                label.className = "gs-bar-label";
                label.textContent = subName.slice(0, 20);
                barInner.appendChild(label);
            }
            td.appendChild(barInner);

            // ポップオーバー + 同一イベントのバーセルのみホバー
            ((barTd, evId) => {
                barTd.addEventListener("mouseenter", () => {
                    PopoverManager.scheduleShow(evId, barTd);
                    if (!DragManager.active) {
                        const row = barTd.closest("tr.gs-event-row");
                        if (row) row.querySelectorAll(`.gs-bar-cell[data-event-id="${evId}"]`).forEach(c => c.classList.add("gs-bar-hover"));
                    }
                });
                barTd.addEventListener("mouseleave", () => {
                    PopoverManager.scheduleHide();
                    const row = barTd.closest("tr.gs-event-row");
                    if (row) row.querySelectorAll(`.gs-bar-cell[data-event-id="${evId}"]`).forEach(c => c.classList.remove("gs-bar-hover"));
                });
            })(td, ev.id);

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

            // バーセルのクリックで編集モーダルを開く
            ((barTd, evId) => {
                barTd.addEventListener("click", (e) => {
                    if (DragManager.started) return;
                    e.stopPropagation();
                    PopoverManager.hide();
                    const evObj = _eventsById[evId];
                    if (evObj) {
                        const catName = barTd.closest("tr.gs-event-row")?.dataset.categoryName;
                        openEventModal(evObj, catName);
                    }
                });
            })(td, ev.id);

            td.dataset.col = c;
            td.dataset.startCol = hit.startCol;
            td.dataset.endCol = hit.endCol;
        } else {
            td.dataset.col = c;
            // 空セルクリックでイベント追加モーダルを開く
            td.style.cursor = "pointer";
            td.addEventListener("click", () => {
                const dateStr = colToStartDate(c, year);
                const eventTitle = evGroup.length > 0 ? evGroup[0].title : null;
                openEventModal(null, cat.name, dateStr, dateStr, eventTitle);
            });
        }
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
        const eventId = td.dataset.eventId || null;
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
            // 行のイベントタイトルを取得（イベント名セルのテキスト）
            const rowEventIds = (tr.dataset.eventIds || "").split(",").filter(Boolean);
            const rowTitle = rowEventIds.length > 0 && _eventsById[rowEventIds[0]]
                ? _eventsById[rowEventIds[0]].title : null;
            DragManager.begin(
                "create", tr, col, e.clientX, e.clientY,
                null, null, catName, -1, -1, year, rowTitle
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
        const eventId = cell.dataset.eventId || null;
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
            onDragCreate(result.categoryName, result.startDate, result.endDate, result.eventTitle);
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
            onDragCreate(result.categoryName, result.startDate, result.endDate, result.eventTitle);
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

