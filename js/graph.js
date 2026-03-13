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
        $select: "subject,start,end,categories,isAllDay,showAs,bodyPreview,location",
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

    return normalizeEvents(allEvents);
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

    try {
        const response = await fetch(`${baseUrl}?${params.toString()}`, {
            headers: { Authorization: `Bearer ${accessToken}` },
        });

        if (!response.ok) return null;

        const data = await response.json();
        if (!data.value || data.value.length === 0) return null;

        const configEvent = data.value[0];
        const bodyContent = configEvent.body?.content || "";

        // HTMLタグ除去（Graph APIはtext/htmlで返す場合がある）
        const cleanContent = bodyContent.replace(/<[^>]*>/g, "").trim();
        if (!cleanContent) return null;

        const parsed = JSON.parse(cleanContent);

        // v1 → v3 マイグレーション
        if (!parsed.version || parsed.version === 1) {
            if (!parsed.categories || !Array.isArray(parsed.categories)) return null;
            const thisYear = String(new Date().getFullYear());
            return {
                configEventId: configEvent.id,
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

            // yearOverrides に記録されている年度を変換
            for (const [yr, ov] of Object.entries(overrides)) {
                const hidden = new Set(ov.hidden || []);
                const filtered = base.filter(c => !hidden.has(c.id));
                const additions = ov.additions || [];
                yearCategories[yr] = [...filtered, ...additions].map(c => ({
                    id: c.id, name: c.name, color: c.color,
                }));
            }

            // 現在の年度が無ければ base をそのままコピー
            const thisYear = String(new Date().getFullYear());
            if (!yearCategories[thisYear]) {
                yearCategories[thisYear] = base.map(c => ({
                    id: c.id, name: c.name, color: c.color,
                }));
            }

            return {
                configEventId: configEvent.id,
                rawConfig: { version: 3, yearCategories },
            };
        }

        // v3: そのまま返す
        if (parsed.version === 3 && parsed.yearCategories) {
            return {
                configEventId: configEvent.id,
                rawConfig: parsed,
            };
        }

        return null;
    } catch (e) {
        console.warn("Config Event parse error:", e);
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
            title: event.subject || "(無題)",
            startDate,
            endDate,
            isAllDay: event.isAllDay,
            categories: event.categories || [],
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
    const baseUrl = await getCalendarBaseUrl(accessToken, "calendar/events");
    const body = buildEventBody(eventData);

    const response = await fetch(`${baseUrl}/${eventId}`, {
        method: "PATCH",
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
        throw new Error(`イベント更新エラー (${response.status}): ${msg}`);
    }

    return await response.json();
}

// ---- 書き込み: イベント削除 ----
async function deleteCalendarEvent(accessToken, eventId) {
    const baseUrl = await getCalendarBaseUrl(accessToken, "calendar/events");

    const response = await fetch(`${baseUrl}/${eventId}`, {
        method: "DELETE",
        headers: { Authorization: `Bearer ${accessToken}` },
    });

    if (!response.ok && response.status !== 204) {
        const errorBody = await response.json().catch(() => ({}));
        const msg = errorBody?.error?.message || response.statusText;
        throw new Error(`イベント削除エラー (${response.status}): ${msg}`);
    }

    return true;
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

    // カテゴリ
    if (eventData.category) {
        body.categories = [eventData.category];
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
