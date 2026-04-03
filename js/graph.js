// Microsoft Graph API モジュール（読み取り＋書き込み対応）

// キャッシュ: グループID
let _cachedGroupId = null;

async function resolveGroupId(accessToken, groupEmail) {
    if (_cachedGroupId) return _cachedGroupId;

    const url = `https://graph.microsoft.com/v1.0/groups?$filter=mail eq '${encodeURIComponent(groupEmail)}'&$select=id,displayName`;

    const response = await fetch(url, {
        headers: { Authorization: `Bearer ${accessToken}` },
    });

    if (!response.ok) {
        const errorBody = await response.json().catch(() => ({}));
        const msg = errorBody?.error?.message || response.statusText;
        if (response.status === 403) {
            throw new Error("グループ情報へのアクセス権限がありません。Azure ADで Group.Read.All の管理者同意が必要です。");
        }
        throw new Error(`グループ検索エラー (${response.status}): ${msg}`);
    }

    const data = await response.json();
    if (!data.value || data.value.length === 0) {
        throw new Error(`グループが見つかりません: ${groupEmail}`);
    }

    _cachedGroupId = data.value[0].id;
    return _cachedGroupId;
}

// ========================================
// ベースURL構築ヘルパー（共通化）
// ========================================
// endpoint: "calendarView" | "calendar/events" など
async function getCalendarBaseUrl(accessToken, endpoint) {
    const calType = graphConfig.calendarType || "me";
    const owner = graphConfig.calendarOwner;

    if (calType === "group") {
        const groupId = await resolveGroupId(accessToken, owner);
        return `https://graph.microsoft.com/v1.0/groups/${groupId}/${endpoint}`;
    } else if (calType === "me") {
        return `https://graph.microsoft.com/v1.0/me/${endpoint}`;
    } else {
        return `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(owner)}/${endpoint}`;
    }
}

// ---- 読み取り ----
async function fetchCalendarEvents(accessToken, year) {
    const startDateTime = `${year}-01-01T00:00:00`;
    const endDateTime = `${year}-12-31T23:59:59`;

    const baseUrl = await getCalendarBaseUrl(accessToken, "calendarView");

    const params = new URLSearchParams({
        startDateTime,
        endDateTime,
        $top: "250",
        $select: "id,type,seriesMasterId,subject,start,end,categories,isAllDay,showAs,bodyPreview,location",
        $orderby: "start/dateTime",
    });

    let allEvents = [];
    let nextUrl = `${baseUrl}?${params.toString()}`;

    while (nextUrl) {
        const response = await fetch(nextUrl, {
            headers: {
                Authorization: `Bearer ${accessToken}`,
                Prefer: 'outlook.timezone="Asia/Tokyo"',
            },
        });

        if (!response.ok) {
            const errorBody = await response.json().catch(() => ({}));
            const msg = errorBody?.error?.message || response.statusText;
            if (response.status === 403) {
                throw new Error("カレンダーへのアクセス権限がありません。管理者にお問い合わせください。");
            }
            throw new Error(`Graph API エラー (${response.status}): ${msg}`);
        }

        const data = await response.json();
        allEvents = allEvents.concat(data.value || []);
        nextUrl = data["@odata.nextLink"] || null;
    }

    const normalized = normalizeEvents(allEvents);
    // 生データ（normalizeEvents前）のcategoriesを保持（移行処理用）
    normalized._rawGraphEvents = allEvents;
    return normalized;
}

// ========================================
// Config Event: カテゴリ設定の読み書き
// ========================================
const CONFIG_EVENT_SUBJECT = "__GYRO_SCHEDULE_CONFIG__";
const CONFIG_EVENT_CATEGORY = "__GYRO_CONFIG__";

// ---- Config Event 取得 ----
async function fetchCategoryConfig(accessToken) {
    const baseUrl = await getCalendarBaseUrl(accessToken, "calendar/events");

    const params = new URLSearchParams({
        $filter: `subject eq '${CONFIG_EVENT_SUBJECT}'`,
        $select: "id,subject,body",
        $top: "1",
    });

    let configEventId = null;
    try {
        const response = await fetch(`${baseUrl}?${params.toString()}`, {
            headers: { Authorization: `Bearer ${accessToken}` },
        });

        if (!response.ok) return null;

        const data = await response.json();
        if (!data.value || data.value.length === 0) return null;

        const configEvent = data.value[0];
        configEventId = configEvent.id; // パース前にIDを保存
        const bodyContent = configEvent.body?.content || "";

        // HTMLタグ除去 + HTMLエンティティをデコード（Graph APIはtext/htmlで返す場合がある）
        const tagStripped = bodyContent.replace(/<[^>]*>/g, "").trim();
        const textarea = document.createElement("textarea");
        textarea.innerHTML = tagStripped;
        const cleanContent = textarea.value.trim();
        if (!cleanContent) {
            return { configEventId, rawConfig: { version: 3, yearCategories: {} } };
        }

        const parsed = JSON.parse(cleanContent);

        // v1 → v3 マイグレーション
        if (!parsed.version || parsed.version === 1) {
            if (!parsed.categories || !Array.isArray(parsed.categories)) {
                return { configEventId, rawConfig: { version: 3, yearCategories: {} } };
            }
            const thisYear = String(new Date().getFullYear());
            return {
                configEventId,
                rawConfig: {
                    version: 3,
                    yearCategories: { [thisYear]: parsed.categories },
                },
            };
        }

        // v2 → v3 マイグレーション
        if (parsed.version === 2) {
            const base = parsed.baseCategories || [];
            const overrides = parsed.yearOverrides || {};
            const yearCategories = {};

            for (const [yr, ov] of Object.entries(overrides)) {
                const hidden = new Set(ov.hidden || []);
                const filtered = base.filter(c => !hidden.has(c.id));
                const additions = ov.additions || [];
                yearCategories[yr] = [...filtered, ...additions].map(c => ({
                    id: c.id, name: c.name, color: c.color,
                }));
            }

            const thisYear = String(new Date().getFullYear());
            if (!yearCategories[thisYear]) {
                yearCategories[thisYear] = base.map(c => ({
                    id: c.id, name: c.name, color: c.color,
                }));
            }

            return { configEventId, rawConfig: { version: 3, yearCategories } };
        }

        // v3: そのまま返す
        if (parsed.version === 3 && parsed.yearCategories) {
            return { configEventId, rawConfig: parsed };
        }

        return { configEventId, rawConfig: { version: 3, yearCategories: {} } };
    } catch (e) {
        console.warn("Config Event parse error:", e);
        // パースエラーでもIDを返す（既存を上書きできるように）
        if (configEventId) {
            return { configEventId, rawConfig: { version: 3, yearCategories: {} } };
        }
        return null;
    }
}

// ---- Config Event 保存（作成 or 更新） ----
async function saveCategoryConfig(accessToken, rawConfig, configEventId) {
    const baseUrl = await getCalendarBaseUrl(accessToken, "calendar/events");
    const configJson = JSON.stringify(rawConfig);

    if (configEventId) {
        // 既存を更新（PATCH）
        const response = await fetch(`${baseUrl}/${configEventId}`, {
            method: "PATCH",
            headers: {
                Authorization: `Bearer ${accessToken}`,
                "Content-Type": "application/json",
            },
            body: JSON.stringify({
                body: { contentType: "text", content: configJson },
            }),
        });

        if (!response.ok) {
            const errorBody = await response.json().catch(() => ({}));
            throw new Error(`Config保存エラー (${response.status}): ${errorBody?.error?.message || response.statusText}`);
        }

        return await response.json();
    } else {
        // 新規作成（POST）
        const response = await fetch(baseUrl, {
            method: "POST",
            headers: {
                Authorization: `Bearer ${accessToken}`,
                "Content-Type": "application/json",
                Prefer: 'outlook.timezone="Asia/Tokyo"',
            },
            body: JSON.stringify({
                subject: CONFIG_EVENT_SUBJECT,
                isAllDay: true,
                start: { dateTime: "2000-01-01T00:00:00", timeZone: "Asia/Tokyo" },
                end: { dateTime: "2000-01-02T00:00:00", timeZone: "Asia/Tokyo" },
                showAs: "free",
                categories: [CONFIG_EVENT_CATEGORY],
                body: { contentType: "text", content: configJson },
            }),
        });

        if (!response.ok) {
            const errorBody = await response.json().catch(() => ({}));
            throw new Error(`Config作成エラー (${response.status}): ${errorBody?.error?.message || response.statusText}`);
        }

        return await response.json();
    }
}

// ---- 通常イベントの正規化 ----
function normalizeEvents(graphEvents) {
    // Config Event を通常イベントから除外
    graphEvents = graphEvents.filter(ev =>
        !(ev.categories || []).includes(CONFIG_EVENT_CATEGORY)
    );
    return graphEvents.map((event) => {
        let startDate = event.start.dateTime.split("T")[0];
        let endDate = event.end.dateTime.split("T")[0];

        if (event.isAllDay) {
            const end = new Date(endDate + "T00:00:00");
            end.setDate(end.getDate() - 1);
            endDate = formatDateYMD(end);
        }

        return {
            id: event.id,
            graphType: event.type || "singleInstance", // "singleInstance" | "occurrence" | "seriesMaster"
            seriesMasterId: event.seriesMasterId || null,
            title: event.subject || "(無題)",
            startDate,
            endDate,
            isAllDay: event.isAllDay,
            categories: (function() {
                const rawCats = event.categories || [];
                // Outlook色カテゴリを除外してアプリカテゴリのみ返す
                const appCats = rawCats.filter(c => !OUTLOOK_COLOR_CATS.has(c));
                if (appCats.length > 0) return appCats;
                // フォールバック: Outlook色カテゴリのみの場合（旧データ）
                if (rawCats.length > 0) {
                    const colorCat = rawCats[0];
                    if (colorCat === "Blue category") return ["朝会"];
                    if (colorCat === "Purple category") return ["GYRO休み"];
                    if (OUTLOOK_COLOR_CATS.has(colorCat)) return []; // 未分類として扱う
                }
                return rawCats;
            })(),
            showAs: event.showAs,
            bodyPreview: event.bodyPreview || "",
            location: event.location?.displayName || "",
            type: "event",
        };
    });
}

// ---- 書き込み: イベント作成 ----
async function createCalendarEvent(accessToken, eventData) {
    const url = await getCalendarBaseUrl(accessToken, "calendar/events");
    const body = buildEventBody(eventData);

    const response = await fetch(url, {
        method: "POST",
        headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
            Prefer: 'outlook.timezone="Asia/Tokyo"',
        },
        body: JSON.stringify(body),
    });

    if (!response.ok) {
        const errorBody = await response.json().catch(() => ({}));
        const msg = errorBody?.error?.message || response.statusText;
        throw new Error(`イベント作成エラー (${response.status}): ${msg}`);
    }

    return await response.json();
}

// ---- 書き込み: イベント更新 ----
async function updateCalendarEvent(accessToken, eventId, eventData) {
    const body = buildEventBody(eventData);
    const headers = {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
        Prefer: 'outlook.timezone="Asia/Tokyo"',
    };

    // calendar/events を最優先（グループカレンダー + calendarView IDと互換性高い）
    const calEventsUrl = await getCalendarBaseUrl(accessToken, "calendar/events");
    let response = await fetch(`${calEventsUrl}/${eventId}`, {
        method: "PATCH", headers, body: JSON.stringify(body),
    });

    // 404 → events エンドポイントでフォールバック
    if (response.status === 404) {
        const eventsUrl = await getCalendarBaseUrl(accessToken, "events");
        response = await fetch(`${eventsUrl}/${eventId}`, {
            method: "PATCH", headers, body: JSON.stringify(body),
        });
    }

    if (!response.ok) {
        const errorBody = await response.json().catch(() => ({}));
        const msg = errorBody?.error?.message || response.statusText;
        throw new Error(`イベント更新エラー (${response.status}): ${msg}`);
    }

    return await response.json();
}

// ---- 書き込み: イベント削除 ----
// グループカレンダーでは calendar/events が正しいパス
// フォールバック: calendar/events → events
async function deleteCalendarEvent(accessToken, eventId) {
    console.log("[DELETE] eventId:", eventId);
    const headers = { Authorization: `Bearer ${accessToken}` };

    // 1. calendar/events を最優先（グループカレンダー + calendarView IDと互換性高い）
    const calEventsUrl = await getCalendarBaseUrl(accessToken, "calendar/events");
    let response = await fetch(`${calEventsUrl}/${eventId}`, { method: "DELETE", headers });
    console.log("[DELETE] calendar/events:", response.status);

    // 2. 404 → events エンドポイントでフォールバック
    if (response.status === 404) {
        const eventsUrl = await getCalendarBaseUrl(accessToken, "events");
        response = await fetch(`${eventsUrl}/${eventId}`, { method: "DELETE", headers });
        console.log("[DELETE] events fallback:", response.status);
    }

    if (!response.ok && response.status !== 204) {
        const errorBody = await response.json().catch(() => ({}));
        const msg = errorBody?.error?.message || response.statusText;
        throw new Error(`イベント削除エラー (${response.status}): ${msg}`);
    }

    return true;
}

// ---- デュアルカテゴリ方式 ----
// Outlook色カテゴリ（定義済み）の一覧
const OUTLOOK_COLOR_CATS = new Set([
    "Blue category", "Green category", "Orange category",
    "Purple category", "Red category", "Yellow category",
    "None",
]);

// アプリカテゴリ名 → Outlook色カテゴリを返す
function _getOutlookColorCategory(appCatName) {
    if (appCatName === "朝会") return "Blue category";
    if (appCatName === "GYRO休み") return "Purple category";
    return "Red category";
}

// ---- ヘルパー: Graph API用のイベントボディ構築 ----
function buildEventBody(eventData) {
    const body = {
        subject: eventData.title,
        isAllDay: true,
        start: {
            dateTime: eventData.startDate + "T00:00:00",
            timeZone: "Asia/Tokyo",
        },
        end: {
            // Graph APIは終日イベントの場合、endは翌日の00:00:00
            dateTime: addDaysToDateStr(eventData.endDate, 1) + "T00:00:00",
            timeZone: "Asia/Tokyo",
        },
        showAs: "free",
    };

    // デュアルカテゴリ: [Outlook色カテゴリ, アプリカテゴリ名]
    if (eventData.category) {
        const colorCat = _getOutlookColorCategory(eventData.category);
        body.categories = [colorCat, eventData.category];
    }

    // メモ
    if (eventData.notes) {
        body.body = {
            contentType: "text",
            content: eventData.notes,
        };
    }

    return body;
}

// ========================================
// Outlookカテゴリ色同期
// ========================================

// Outlookプリセット色とHex値のマッピング
const OUTLOOK_PRESETS = [
    { name: "preset0",  hex: "#e7a1a2" }, // Red
    { name: "preset1",  hex: "#f9ba89" }, // Orange
    { name: "preset2",  hex: "#f7dd8f" }, // Brown/Peach
    { name: "preset3",  hex: "#fcfa90" }, // Yellow
    { name: "preset4",  hex: "#78d168" }, // Green
    { name: "preset5",  hex: "#9fdcc9" }, // Teal
    { name: "preset6",  hex: "#c6d2b0" }, // Olive
    { name: "preset7",  hex: "#9db7e8" }, // Blue
    { name: "preset8",  hex: "#b5a1e2" }, // Purple
    { name: "preset9",  hex: "#daaec2" }, // Cranberry
    { name: "preset10", hex: "#dad9dc" }, // Steel
    { name: "preset11", hex: "#6b7994" }, // DarkSteel
    { name: "preset12", hex: "#bfbfbf" }, // Gray
    { name: "preset13", hex: "#6f6f6f" }, // DarkGray
    { name: "preset14", hex: "#4f4f4f" }, // Black
    { name: "preset15", hex: "#c11a25" }, // DarkRed
    { name: "preset16", hex: "#e2620d" }, // DarkOrange
    { name: "preset17", hex: "#c79930" }, // DarkBrown
    { name: "preset18", hex: "#b9b300" }, // DarkYellow
    { name: "preset19", hex: "#368f20" }, // DarkGreen
    { name: "preset20", hex: "#329b7a" }, // DarkTeal
    { name: "preset21", hex: "#778b45" }, // DarkOlive
    { name: "preset22", hex: "#2858a5" }, // DarkBlue
    { name: "preset23", hex: "#5c3fa3" }, // DarkPurple
    { name: "preset24", hex: "#93446b" }, // DarkCranberry
];

function _hexToRgb(hex) {
    const h = hex.replace("#", "");
    return [parseInt(h.substring(0, 2), 16), parseInt(h.substring(2, 4), 16), parseInt(h.substring(4, 6), 16)];
}

function _closestPreset(hexColor) {
    const [r, g, b] = _hexToRgb(hexColor);
    let best = "preset7";
    let bestDist = Infinity;
    for (const p of OUTLOOK_PRESETS) {
        const [pr, pg, pb] = _hexToRgb(p.hex);
        const dist = (r - pr) ** 2 + (g - pg) ** 2 + (b - pb) ** 2;
        if (dist < bestDist) { bestDist = dist; best = p.name; }
    }
    return best;
}

// Outlookカテゴリ色の明示的マッピング
function _getOutlookPreset(catName) {
    if (catName === "朝会") return "preset7";        // Blue
    if (catName === "GYRO休み") return "preset23";    // DarkPurple
    return "preset0";                                  // Red（それ以外すべて）
}

// Outlookのマスターカテゴリリストにブラウザ側の色を同期
async function syncOutlookCategoryColors(accessToken, categories) {
    const baseUrl = "https://graph.microsoft.com/v1.0/me/outlook/masterCategories";

    // 1. 既存のマスターカテゴリを取得
    let existing = [];
    try {
        const res = await fetch(baseUrl, {
            headers: { Authorization: `Bearer ${accessToken}` },
        });
        if (res.ok) {
            const data = await res.json();
            existing = data.value || [];
        }
    } catch (e) {
        console.warn("[カテゴリ色同期] マスターカテゴリ取得失敗:", e.message);
        return;
    }

    const existingMap = new Map(existing.map(c => [c.displayName, c]));

    // 2. 各カテゴリを同期（明示的な色マッピング）
    for (const cat of categories) {
        const presetColor = _getOutlookPreset(cat.name);
        const existingCat = existingMap.get(cat.name);

        try {
            if (existingCat) {
                // 既存 → 色が違えば更新
                if (existingCat.color !== presetColor) {
                    console.log(`[カテゴリ色同期] "${cat.name}" 更新: ${existingCat.color} → ${presetColor}`);
                    const res = await fetch(`${baseUrl}/${encodeURIComponent(existingCat.id)}`, {
                        method: "PATCH",
                        headers: {
                            Authorization: `Bearer ${accessToken}`,
                            "Content-Type": "application/json",
                        },
                        body: JSON.stringify({ color: presetColor }),
                    });
                    if (!res.ok) {
                        const err = await res.json().catch(() => ({}));
                        console.warn(`[カテゴリ色同期] "${cat.name}" 更新失敗 (${res.status}):`, err?.error?.message);
                    } else {
                        console.log(`[カテゴリ色同期] "${cat.name}" 更新成功`);
                    }
                } else {
                    console.log(`[カテゴリ色同期] "${cat.name}" 変更なし (${presetColor})`);
                }
            } else {
                // 新規 → 作成
                console.log(`[カテゴリ色同期] "${cat.name}" 新規作成: ${presetColor}`);
                const res = await fetch(baseUrl, {
                    method: "POST",
                    headers: {
                        Authorization: `Bearer ${accessToken}`,
                        "Content-Type": "application/json",
                    },
                    body: JSON.stringify({ displayName: cat.name, color: presetColor }),
                });
                if (!res.ok) {
                    const err = await res.json().catch(() => ({}));
                    console.warn(`[カテゴリ色同期] "${cat.name}" 作成失敗 (${res.status}):`, err?.error?.message);
                }
            }
        } catch (e) {
            console.warn(`[カテゴリ色同期] "${cat.name}" の同期失敗:`, e.message);
        }
    }

    console.log(`[カテゴリ色同期] ${categories.length}件のカテゴリ色を同期しました`);
}

// ========================================
// GitHub自動公開: data/events.json を更新
// ========================================
let _cachedGitHubSha = null;
let _cachedGitHubData = null;

async function _fetchGitHubFile(token) {
    const { owner, repo, branch, path } = githubConfig;
    const url = `https://api.github.com/repos/${owner}/${repo}/contents/${path}?ref=${branch}&_t=${Date.now()}`;
    const res = await fetch(url, {
        headers: { Authorization: `token ${token}`, "If-None-Match": "" },
        cache: "no-store",
    });
    if (!res.ok) return { sha: null, data: { lastUpdated: null, years: {} } };
    const fileData = await res.json();
    const data = JSON.parse(atob(fileData.content));
    return { sha: fileData.sha, data };
}

async function publishEventsToGitHub(events, categories, year) {
    const { owner, repo, branch, path, token } = githubConfig;
    if (!token) {
        console.log("[GitHub公開] トークン未設定。スキップ。");
        return;
    }

    const maxRetries = 3;
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        let sha, existingData;

        // キャッシュSHAがあり初回試行 → GETスキップ（409回避）
        if (_cachedGitHubSha && attempt === 1) {
            sha = _cachedGitHubSha;
            existingData = _cachedGitHubData ? JSON.parse(JSON.stringify(_cachedGitHubData)) : { lastUpdated: null, years: {} };
        } else {
            // キャッシュなし or リトライ → APIから再取得
            try {
                const file = await _fetchGitHubFile(token);
                sha = file.sha;
                existingData = file.data;
            } catch (e) {
                console.warn("[GitHub公開] ファイル取得失敗:", e.message);
                existingData = { lastUpdated: null, years: {} };
                sha = null;
            }
        }

        // 現在年度のデータを更新
        existingData.years[String(year)] = {
            events: events.map(e => ({
                id: e.id,
                title: e.title,
                startDate: e.startDate,
                endDate: e.endDate,
                categories: e.categories || [],
                bodyPreview: e.bodyPreview || "",
            })),
            categories: categories.map(c => ({
                id: c.id,
                name: c.name,
                color: c.color,
            })),
        };
        existingData.lastUpdated = new Date().toISOString();

        // GitHub Contents APIでコミット
        const putUrl = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
        const content = btoa(unescape(encodeURIComponent(JSON.stringify(existingData, null, 2))));
        const body = {
            message: `auto: update events.json (${new Date().toLocaleDateString("ja-JP")})`,
            content,
            branch,
        };
        if (sha) body.sha = sha;

        const putRes = await fetch(putUrl, {
            method: "PUT",
            headers: {
                Authorization: `token ${token}`,
                "Content-Type": "application/json",
            },
            body: JSON.stringify(body),
        });

        if (putRes.ok) {
            const result = await putRes.json();
            _cachedGitHubSha = result.content?.sha || null;
            _cachedGitHubData = existingData;
            console.log("[GitHub公開] events.json を更新しました");
            return;
        }

        if (putRes.status === 409 && attempt < maxRetries) {
            _cachedGitHubSha = null; // キャッシュ破棄して次回GETで再取得
            console.log(`[GitHub公開] SHA競合、リトライ (${attempt}/${maxRetries})`);
            await new Promise(r => setTimeout(r, 2000));
            continue;
        }

        const err = await putRes.json().catch(() => ({}));
        console.warn("[GitHub公開] 更新失敗:", putRes.status, err.message || "");
        _cachedGitHubSha = null;
        return;
    }
}
