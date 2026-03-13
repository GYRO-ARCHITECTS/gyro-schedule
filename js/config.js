// Azure AD 設定
const msalConfig = {
    auth: {
        clientId: "86199793-1bca-4357-b750-9d9b40d437ef",
        authority: "https://login.microsoftonline.com/ec5acb07-3045-44e7-9454-fcf1d00198d5",
        redirectUri: window.location.hostname === "localhost"
            ? window.location.origin + "/auth-redirect.html"
            : window.location.origin + "/gyro-schedule/auth-redirect.html",
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false,
    },
    system: { loggerOptions: { logLevel: 3 } },
};

const graphConfig = {
    // ReadWrite に変更（双方向同期のため）
    scopes: ["Calendars.ReadWrite", "Group.ReadWrite.All"],
    calendarType: "group",
    calendarOwner: "msteams_c346c9@gyroarchitects01.onmicrosoft.com",
    year: new Date().getFullYear(),
};

// ========================================
// カラーユーティリティ
// ========================================
// メインカラーから bg（バー背景）と border（バー枠線）を自動導出
// bg: 原色を白で25%だけ薄めた鮮やかな色
// border: 原色そのまま（左端アクセント用）
function deriveColors(hex) {
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    // 白方向に25%ブレンド → 鮮やかだが少し柔らかい色
    const mix = (c) => Math.round(c + (255 - c) * 0.25);
    const bgHex = `#${[mix(r),mix(g),mix(b)].map(v=>v.toString(16).padStart(2,'0')).join('')}`;
    return {
        color: hex,
        bg: bgHex,
        border: hex,
    };
}

// カラープリセット（管理UI用: 20色）
const COLOR_PRESETS = [
    "#ef4444", "#f97316", "#eab308", "#22c55e", "#14b8a6",
    "#3b82f6", "#6366f1", "#a855f7", "#ec4899", "#64748b",
    "#dc2626", "#ea580c", "#ca8a04", "#16a34a", "#0d9488",
    "#2563eb", "#4f46e5", "#9333ea", "#db2777", "#475569",
];

// ========================================
// カテゴリ定義（動的管理対応）
// ========================================
// ハードコードされたデフォルト（初回用・フォールバック用）
const DEFAULT_CATEGORIES = [
    { id: "holiday",         name: "休日",          color: "#ef4444" },
    { id: "morning_meeting", name: "朝会",          color: "#3b82f6" },
];

// 動的カテゴリ（loadCalendar時にOutlookから読み込まれて差し替わる）
let CATEGORIES = DEFAULT_CATEGORIES.map(c => ({ ...c, ...deriveColors(c.color) }));

// カテゴリ名 → 定義の逆引き
function getCategoryDef(name) {
    return CATEGORIES.find(c => c.name === name) || CATEGORIES[CATEGORIES.length - 1];
}
