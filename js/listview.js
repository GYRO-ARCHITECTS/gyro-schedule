// リストビュー描画エンジン
// イベントを月ごとにグループ化し、日付順に一覧表示

const LV_MONTH_NAMES = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"];

function renderListView(graphEvents, holidays, year, holidaySet) {
    const container = document.getElementById("timeline-container");

    // バウンス防止
    const prevHeight = container.offsetHeight;
    if (prevHeight > 0) container.style.minHeight = prevHeight + "px";

    // 全イベント + 祝日を統合してソート
    const yearStr = String(year);
    const allItems = [];

    graphEvents.forEach(ev => {
        if (ev.startDate && ev.startDate.startsWith(yearStr)) {
            allItems.push({ ...ev, _type: "event" });
        }
    });

    holidays.forEach(h => {
        if (h.startDate && h.startDate.startsWith(yearStr)) {
            allItems.push({ ...h, _type: "holiday" });
        }
    });

    allItems.sort((a, b) => a.startDate.localeCompare(b.startDate));

    // 月でグループ化
    const grouped = {};
    for (let m = 0; m < 12; m++) grouped[m] = [];

    allItems.forEach(item => {
        const month = parseInt(item.startDate.substring(5, 7), 10) - 1;
        if (month >= 0 && month < 12) grouped[month].push(item);
    });

    // DOM構築
    const wrapper = document.createElement("div");
    wrapper.className = "lv-container";

    let hasAny = false;

    for (let m = 0; m < 12; m++) {
        const items = grouped[m];
        if (items.length === 0) continue;
        hasAny = true;

        const group = document.createElement("div");
        group.className = "lv-month-group";

        const header = document.createElement("div");
        header.className = "lv-month-header";
        header.textContent = LV_MONTH_NAMES[m];
        group.appendChild(header);

        const table = document.createElement("table");
        table.className = "lv-table";
        table.setAttribute("aria-label", `${LV_MONTH_NAMES[m]}のイベント一覧`);
        const tbody = document.createElement("tbody");

        items.forEach(item => {
            const tr = document.createElement("tr");

            if (item._type === "holiday") {
                tr.className = "lv-holiday-row";

                const tdDate = document.createElement("td");
                tdDate.className = "lv-date";
                tdDate.textContent = formatDateRangeShort(item.startDate, item.endDate);
                tr.appendChild(tdDate);

                const tdCat = document.createElement("td");
                tdCat.className = "lv-category";
                tdCat.innerHTML = '<span class="lv-holiday-badge">祝日</span>';
                tr.appendChild(tdCat);

                const tdTitle = document.createElement("td");
                tdTitle.className = "lv-title";
                tdTitle.textContent = item.title;
                tr.appendChild(tdTitle);
            } else {
                tr.className = "lv-event-row";
                tr.dataset.eventId = item.id;

                // 土日・祝日クラスの付与
                if (holidaySet && holidaySet.has(item.startDate)) {
                    tr.classList.add("lv-on-holiday");
                } else {
                    const startD = new Date(item.startDate + "T00:00:00");
                    const dow = startD.getDay();
                    if (dow === 0) tr.classList.add("lv-on-sunday");
                    else if (dow === 6) tr.classList.add("lv-on-saturday");
                }

                const tdDate = document.createElement("td");
                tdDate.className = "lv-date";
                tdDate.textContent = formatDateRangeShort(item.startDate, item.endDate);
                tr.appendChild(tdDate);

                const tdCat = document.createElement("td");
                tdCat.className = "lv-category";
                const catName = (item.categories && item.categories[0]) || "";
                const catDef = getCategoryDef(catName);
                if (catDef) {
                    tdCat.innerHTML = `<span class="lv-cat-badge" style="background:${catDef.color}">${catDef.name}</span>`;
                } else if (catName) {
                    tdCat.innerHTML = `<span class="lv-cat-badge" style="background:#94a3b8">${catName}</span>`;
                }
                tr.appendChild(tdCat);

                const tdTitle = document.createElement("td");
                tdTitle.className = "lv-title";
                tdTitle.textContent = item.title;
                tr.appendChild(tdTitle);

                // クリックで編集モーダル
                tr.setAttribute("tabindex", "0");
                tr.setAttribute("role", "button");
                tr.setAttribute("aria-label", `${item.title}を編集（${formatDateRangeShort(item.startDate, item.endDate)}）`);
                tr.addEventListener("click", () => {
                    const evObj = _cachedGraphEvents.find(e => e.id === item.id);
                    if (evObj) openEventModal(evObj, catName);
                });
                tr.addEventListener("keydown", (e) => {
                    if (e.key === "Enter" || e.key === " ") {
                        e.preventDefault();
                        const evObj = _cachedGraphEvents.find(ev => ev.id === item.id);
                        if (evObj) openEventModal(evObj, catName);
                    }
                });
            }

            tbody.appendChild(tr);
        });

        table.appendChild(tbody);
        group.appendChild(table);
        wrapper.appendChild(group);
    }

    if (!hasAny) {
        const empty = document.createElement("div");
        empty.className = "lv-empty-month";
        empty.textContent = `${year}年のイベントはありません`;
        wrapper.appendChild(empty);
    }

    container.replaceChildren(wrapper);
    container.style.minHeight = "";
}

// ヘルパー関数は utils.js に統合済み
