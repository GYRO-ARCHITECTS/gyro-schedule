// 共通ユーティリティ関数

// Date → "YYYY-MM-DD" 文字列
function formatDateYMD(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
}

// "YYYY-MM-DD" → Date
function parseDateStr(str) {
    const [y, m, d] = str.split("-").map(Number);
    return new Date(y, m - 1, d);
}

// 指定月の日数を返す
function daysInMonth(year, month) {
    return new Date(year, month + 1, 0).getDate();
}

// 日付文字列に日数を加算
function addDaysToDateStr(dateStr, days) {
    const d = new Date(dateStr + "T00:00:00");
    d.setDate(d.getDate() + days);
    return formatDateYMD(d);
}

// 日付範囲を "M/D" or "M/D - M/D" にフォーマット
function formatDateRangeShort(startDate, endDate) {
    const s = startDate.split("-");
    const startStr = `${parseInt(s[1])}/${parseInt(s[2])}`;
    if (startDate === endDate) return startStr;
    const e = endDate.split("-");
    return `${startStr} - ${parseInt(e[1])}/${parseInt(e[2])}`;
}

// HTMLエスケープ
function escapeHtml(text) {
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
}
