// 年間カレンダービュー描画エンジン
// 12ヶ月を 4列×3行 のグリッドで一覧表示

const CAL_WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];
const CAL_MAX_PILLS_YEAR = 2; // 年間ビューでセル内に表示するイベントの最大数

// ========================================
// メイン描画: 年間12ヶ月カレンダー
// ========================================
function renderCalendar(graphEvents, holidays, year, _month, holidaySet) {
    const container = document.getElementById("timeline-container");

    // バウンス防止
    const prevHeight = container.offsetHeight;
    if (prevHeight > 0) container.style.minHeight = prevHeight + "px";

    // 祝日の名前をマップ化 (dateStr -> title)
    const holidayNames = {};
    holidays.forEach(h => {
        if (h.startDate) holidayNames[h.startDate] = h.title;
    });

    const todayStr = formatDateYMD(new Date());
    const today = new Date();

    // 年間グリッドコンテナ
    const yearGrid = document.createElement("div");
    yearGrid.className = "cal-year-grid";
    yearGrid.setAttribute("role", "region");
    yearGrid.setAttribute("aria-label", `${year}年 年間カレンダー`);

    // 12ヶ月を生成
    for (let m = 0; m < 12; m++) {
        const card = _buildMonthCard(graphEvents, holidays, holidayNames, holidaySet, year, m, todayStr, today);
        yearGrid.appendChild(card);
    }

    container.replaceChildren(yearGrid);
    container.style.minHeight = "";
}

// ========================================
// 1ヶ月カード生成
// ========================================
function _buildMonthCard(graphEvents, holidays, holidayNames, holidaySet, year, month, todayStr, today) {
    const card = document.createElement("div");
    card.className = "cal-month-card";
    const isCurrent = (today.getFullYear() === year && today.getMonth() === month);
    if (isCurrent) card.classList.add("cal-month-current");

    // 月タイトル
    const title = document.createElement("div");
    title.className = "cal-month-title";
    title.textContent = `${month + 1}月`;
    card.appendChild(title);

    // 曜日ヘッダー
    const weekHeader = document.createElement("div");
    weekHeader.className = "cal-weekday-header";
    CAL_WEEKDAYS.forEach((day, i) => {
        const div = document.createElement("div");
        div.className = "cal-weekday";
        if (i === 6) div.classList.add("cal-weekday-sat");
        if (i === 0) div.classList.add("cal-weekday-sun");
        div.textContent = day;
        weekHeader.appendChild(div);
    });
    card.appendChild(weekHeader);

    // 日付→イベントインデックス
    const dateIndex = buildCalendarDateIndex(graphEvents, year, month);

    // グリッド計算
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    const daysInMonth = lastDay.getDate();
    const firstDayOfWeek = firstDay.getDay(); // 日曜始まり
    const totalCells = Math.ceil((firstDayOfWeek + daysInMonth) / 7) * 7;

    // 日グリッド
    const grid = document.createElement("div");
    grid.className = "cal-grid";

    for (let i = 0; i < totalCells; i++) {
        const dayNum = i - firstDayOfWeek + 1;
        const cell = document.createElement("div");
        cell.className = "cal-day";

        if (dayNum < 1 || dayNum > daysInMonth) {
            // 前月/翌月
            cell.classList.add("cal-other-month");
            const otherDate = new Date(year, month, dayNum);
            const numDiv = document.createElement("div");
            numDiv.className = "cal-day-number";
            numDiv.textContent = otherDate.getDate();
            cell.appendChild(numDiv);
        } else {
            // 当月
            const dateStr = `${year}-${String(month + 1).padStart(2, "0")}-${String(dayNum).padStart(2, "0")}`;

            if (dateStr === todayStr) cell.classList.add("cal-today");

            const isHoliday = holidaySet && holidaySet.has(dateStr);
            if (isHoliday) cell.classList.add("cal-holiday");

            const dow = (i % 7);
            if (dow === 6) cell.classList.add("cal-saturday");
            if (dow === 0) cell.classList.add("cal-sunday");

            // 日番号
            const numDiv = document.createElement("div");
            numDiv.className = "cal-day-number";
            numDiv.textContent = dayNum;
            cell.appendChild(numDiv);

            // イベントドット表示
            const events = dateIndex[dateStr] || [];
            if (events.length > 0) {
                const dotRow = document.createElement("div");
                dotRow.className = "cal-dot-row";

                const showCount = Math.min(events.length, CAL_MAX_PILLS_YEAR);
                for (let j = 0; j < showCount; j++) {
                    const ev = events[j];
                    const catName = (ev.categories && ev.categories[0]) || "";
                    const catDef = getCategoryDef(catName);
                    const color = catDef ? catDef.color : "#94a3b8";

                    const dot = document.createElement("span");
                    dot.className = "cal-event-dot";
                    dot.style.backgroundColor = color;
                    dot.title = ev.title;
                    dotRow.appendChild(dot);
                }
                if (events.length > CAL_MAX_PILLS_YEAR) {
                    const plus = document.createElement("span");
                    plus.className = "cal-dot-more";
                    plus.textContent = `+${events.length - CAL_MAX_PILLS_YEAR}`;
                    dotRow.appendChild(plus);
                }
                cell.appendChild(dotRow);
            }

            // クリック → 新規作成
            cell.addEventListener("click", () => {
                openEventModal(null, null, dateStr, dateStr);
            });
        }

        grid.appendChild(cell);
    }

    card.appendChild(grid);
    return card;
}

// ========================================
// 日付→イベント配列インデックスの構築
// ========================================
function buildCalendarDateIndex(events, year, month) {
    const index = {};
    const mStr = String(month + 1).padStart(2, "0");
    const lastDay = new Date(year, month + 1, 0).getDate();
    const monthStart = `${year}-${mStr}-01`;
    const monthEnd = `${year}-${mStr}-${String(lastDay).padStart(2, "0")}`;

    events.forEach(ev => {
        if (!ev.startDate || !ev.endDate) return;
        if (ev.endDate < monthStart || ev.startDate > monthEnd) return;

        const start = ev.startDate < monthStart ? monthStart : ev.startDate;
        const end = ev.endDate > monthEnd ? monthEnd : ev.endDate;

        const d = new Date(start + "T00:00:00");
        const endD = new Date(end + "T00:00:00");

        while (d <= endD) {
            const key = formatDateYMD(d);
            if (!index[key]) index[key] = [];
            index[key].push(ev);
            d.setDate(d.getDate() + 1);
        }
    });

    return index;
}
