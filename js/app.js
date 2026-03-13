// アプリケーション制御（SPA楽観的更新版 + マルチビュー対応）
let currentYear = graphConfig.year;
let _editingEvent = null;
let _lastFocusBeforeModal = null; // モーダル前のフォーカス保持

// カテゴリ設定（Config Event）
let _configEventId = null; // Outlook上のConfig EventのID
let _rawConfig = null;     // Config Event の生データ（v2形式）

// ========================================
// ビュー状態管理
// ========================================
let _currentView = "timeline";  // "timeline" | "calendar" | "list"
let _currentMonth = new Date().getMonth(); // 0-11 (カレンダービュー用)

// ========================================
// ローカルイベントキャッシュ（SPA の核）
// ========================================
let _cachedGraphEvents = [];   // Graph API から取得したイベント
let _cachedHolidays = [];      // 祝日データ
let _cachedHolidaySet = null;  // 祝日判定用Set

// ---- アクセシビリティ: ステータス通知 ----
function announceStatus(message) {
    const el = document.getElementById("a11y-status");
    if (el) {
        el.textContent = "";
        requestAnimationFrame(() => { el.textContent = message; });
    }
}

// キャッシュからスクロール位置を保持したまま再描画
// initialLoad=true の場合、タイムラインは今月へ自動スクロール
function rerenderFromCache(initialLoad) {
    const wrapper = document.getElementById("timeline-wrapper");
    const scrollLeft = wrapper ? wrapper.scrollLeft : 0;
    const scrollTop = wrapper ? wrapper.scrollTop : 0;

    switch (_currentView) {
        case "timeline": {
            // モードのオフセットテーブルが年と一致しなければ再構築
            if (_tlYear !== currentYear) setTimelineMode(_tlMode, currentYear);
            const allEvents = [..._cachedHolidays, ..._cachedGraphEvents];
            if (initialLoad) {
                renderTimeline(allEvents, currentYear, _cachedHolidaySet);
            } else {
                renderTimeline(allEvents, currentYear, _cachedHolidaySet, { skipScroll: true });
                if (wrapper) {
                    wrapper.scrollLeft = scrollLeft;
                    wrapper.scrollTop = scrollTop;
                }
            }
            break;
        }
        case "calendar":
            renderCalendar(_cachedGraphEvents, _cachedHolidays, currentYear, _currentMonth, _cachedHolidaySet);
            break;
        case "list":
            renderListView(_cachedGraphEvents, _cachedHolidays, currentYear, _cachedHolidaySet);
            break;
    }
}

// ========================================
// 初期化
// ========================================
document.addEventListener("DOMContentLoaded", async () => {
    setupUI();
    setupModal();
    setupCategoryManager();

    try {
        await initAuth();
    } catch (error) {
        console.error("Auth init error:", error);
        // 認証初期化に失敗しても公開ビューは表示
        loadPublicCalendar();
        return;
    }

    const account = getActiveAccount();
    if (account) {
        await loadCalendar();
    } else {
        // 未サインインでも公開ビュー（祝日+デフォルトカテゴリ）を表示
        loadPublicCalendar();
    }
});

// ---- UI初期設定 ----
function setupUI() {
    updateYearLabel();
    updateMonthLabel();

    document.getElementById("sign-in-btn").addEventListener("click", handleSignIn);
    document.getElementById("header-sign-in-btn").addEventListener("click", handleSignIn);
    document.getElementById("sign-out-btn").addEventListener("click", handleSignOut);
    document.getElementById("prev-year-btn").addEventListener("click", () => changeYear(-1));
    document.getElementById("next-year-btn").addEventListener("click", () => changeYear(1));

    // 今日ボタン（ビューごとに挙動分岐）
    document.getElementById("today-btn").addEventListener("click", () => {
        const thisYear = new Date().getFullYear();
        const thisMonth = new Date().getMonth();

        if (currentYear !== thisYear) {
            currentYear = thisYear;
            graphConfig.year = thisYear;
            _currentMonth = thisMonth;
            updateYearLabel();
            updateMonthLabel();
            if (getActiveAccount()) loadCalendar(); else loadPublicCalendar();
        } else if (_currentView === "calendar") {
            _currentMonth = thisMonth;
            updateMonthLabel();
            rerenderFromCache();
        } else if (_currentView === "timeline") {
            scrollToToday();
        }
        announceStatus("今日の位置へ移動しました");
    });

    // ＋追加ボタン
    document.getElementById("add-event-btn").addEventListener("click", () => {
        openEventModal(null, null);
    });

    // 設定ボタン
    document.getElementById("settings-btn").addEventListener("click", () => {
        openCategoryManager();
    });

    // 公開データ書き出しボタン
    document.getElementById("export-btn").addEventListener("click", handleExportPublicData);

    // ビュー切替
    setupViewSwitcher();

    // タイムラインモード切替
    setupTimelineModeSwitcher();

    // 月ナビ
    setupMonthNav();

    // fitモード用リサイズハンドラ
    let _resizeTimer = 0;
    window.addEventListener("resize", () => {
        if (_tlMode !== "fit" || _currentView !== "timeline") return;
        clearTimeout(_resizeTimer);
        _resizeTimer = setTimeout(() => rerenderFromCache(true), 200);
    });

    // モバイル初期ビュー: スマホではリストビューをデフォルトにする
    if (window.innerWidth <= 480 && _currentView === "timeline") {
        const listBtn = document.querySelector('.view-btn[data-view="list"]');
        if (listBtn) listBtn.click();
    }

    // orientationchange: 画面回転時に再描画
    window.addEventListener("orientationchange", () => {
        setTimeout(() => {
            if (_cachedGraphEvents.length > 0 || _cachedHolidays.length > 0) {
                rerenderFromCache(true);
            }
        }, 300);
    });
}

// ---- ビュー切替 ----
function setupViewSwitcher() {
    const viewBtns = document.querySelectorAll(".view-btn");

    viewBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            const view = btn.dataset.view;
            if (view === _currentView) return;

            // ボタンのアクティブ状態 + ARIA
            viewBtns.forEach(b => {
                b.classList.remove("active");
                b.setAttribute("aria-selected", "false");
            });
            btn.classList.add("active");
            btn.setAttribute("aria-selected", "true");

            _currentView = view;

            // 月ナビの表示/非表示（年間カレンダーでは不要）
            document.getElementById("month-nav").style.display = "none";

            // タイムラインモードバーの表示/非表示
            document.getElementById("timeline-mode-bar").style.display =
                (view === "timeline") ? "flex" : "none";

            // データがあれば再描画（ビュー切替は初回表示扱い）
            if (_cachedGraphEvents.length > 0 || _cachedHolidays.length > 0) {
                rerenderFromCache(true);
            }

            const viewNames = { timeline: "タイムライン", calendar: "カレンダー", list: "リスト" };
            announceStatus(`${viewNames[view]}ビューに切り替えました`);
        });

        // キーボード: 矢印キーでタブ切替
        btn.addEventListener("keydown", (e) => {
            const btns = Array.from(viewBtns);
            const idx = btns.indexOf(btn);
            let next = -1;

            if (e.key === "ArrowRight" || e.key === "ArrowDown") {
                next = (idx + 1) % btns.length;
            } else if (e.key === "ArrowLeft" || e.key === "ArrowUp") {
                next = (idx - 1 + btns.length) % btns.length;
            } else if (e.key === "Home") {
                next = 0;
            } else if (e.key === "End") {
                next = btns.length - 1;
            }

            if (next >= 0) {
                e.preventDefault();
                btns[next].focus();
                btns[next].click();
            }
        });
    });

    // roving tabindex
    viewBtns.forEach((btn, i) => {
        btn.setAttribute("tabindex", i === 0 ? "0" : "-1");
    });
}

// ---- タイムラインモード切替 ----
function setupTimelineModeSwitcher() {
    const modeBtns = document.querySelectorAll(".tl-mode-btn");

    modeBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            const mode = btn.dataset.mode;
            if (mode === _tlMode) return;

            modeBtns.forEach(b => {
                b.classList.remove("active");
                b.setAttribute("aria-checked", "false");
                b.setAttribute("tabindex", "-1");
            });
            btn.classList.add("active");
            btn.setAttribute("aria-checked", "true");
            btn.setAttribute("tabindex", "0");

            setTimelineMode(mode, currentYear);

            if (_cachedGraphEvents.length > 0 || _cachedHolidays.length > 0) {
                rerenderFromCache(true);
            }

            const modeNames = { week: "週", fiveday: "5日", day: "1日", fit: "全体" };
            announceStatus(`表示間隔を${modeNames[mode]}に切り替えました`);
        });

        btn.addEventListener("keydown", (e) => {
            const btns = Array.from(modeBtns);
            const idx = btns.indexOf(btn);
            let next = -1;

            if (e.key === "ArrowRight" || e.key === "ArrowDown") {
                next = (idx + 1) % btns.length;
            } else if (e.key === "ArrowLeft" || e.key === "ArrowUp") {
                next = (idx - 1 + btns.length) % btns.length;
            } else if (e.key === "Home") {
                next = 0;
            } else if (e.key === "End") {
                next = btns.length - 1;
            }

            if (next >= 0) {
                e.preventDefault();
                btns[next].focus();
                btns[next].click();
            }
        });
    });

    // roving tabindex: set tabindex=0 on the active (default) button
    modeBtns.forEach((btn) => {
        btn.setAttribute("tabindex", btn.classList.contains("active") ? "0" : "-1");
    });
}

// ---- 月ナビ ----
function setupMonthNav() {
    document.getElementById("prev-month-btn").addEventListener("click", () => {
        _currentMonth--;
        if (_currentMonth < 0) {
            _currentMonth = 11;
            currentYear--;
            graphConfig.year = currentYear;
            updateYearLabel();
            updateMonthLabel();
            if (getActiveAccount()) loadCalendar(); else loadPublicCalendar();
            return;
        }
        updateMonthLabel();
        rerenderFromCache();
    });

    document.getElementById("next-month-btn").addEventListener("click", () => {
        _currentMonth++;
        if (_currentMonth > 11) {
            _currentMonth = 0;
            currentYear++;
            graphConfig.year = currentYear;
            updateYearLabel();
            updateMonthLabel();
            if (getActiveAccount()) loadCalendar(); else loadPublicCalendar();
            return;
        }
        updateMonthLabel();
        rerenderFromCache();
    });
}

function updateMonthLabel() {
    const el = document.getElementById("month-label");
    if (el) el.textContent = `${currentYear}年 ${_currentMonth + 1}月`;
}

// ---- カテゴリselect動的生成 ----
function populateCategorySelect() {
    const sel = document.getElementById("evt-category");
    sel.innerHTML = "";
    CATEGORIES.forEach(cat => {
        const opt = document.createElement("option");
        opt.value = cat.name;
        opt.textContent = cat.name;
        sel.appendChild(opt);
    });
}

// ---- 年度ごとの有効カテゴリ算出（v3形式） ----
function _buildCategoriesForYear(rawConfig, year) {
    const cats = rawConfig.yearCategories?.[String(year)] || [];
    return cats.map(c => ({ ...c, ...deriveColors(c.color) }));
}

// ---- モーダル設定 ----
function setupModal() {
    const catSelect = document.getElementById("evt-category");

    populateCategorySelect();

    // カテゴリ変更 → アクセントバー + プレビュードット連動
    catSelect.addEventListener("change", () => {
        _updateCategoryPreview(catSelect.value);
    });

    // 開始日 → 終了日min属性連動
    const startInput = document.getElementById("evt-start");
    const endInput = document.getElementById("evt-end");
    startInput.addEventListener("change", () => {
        if (startInput.value) {
            endInput.min = startInput.value;
            if (endInput.value && endInput.value < startInput.value) {
                endInput.value = startInput.value;
            }
        }
    });

    // 入力時にエラー解除
    document.querySelectorAll("#event-modal .form-group[data-field] input, #event-modal .form-group[data-field] select").forEach(el => {
        const handler = () => {
            const fg = el.closest(".form-group");
            if (fg && fg.classList.contains("has-error")) {
                fg.classList.remove("has-error");
                const errSpan = fg.querySelector(".form-error");
                if (errSpan) errSpan.textContent = "";
            }
        };
        el.addEventListener("input", handler);
        el.addEventListener("change", handler);
    });

    document.getElementById("modal-close-btn").addEventListener("click", closeModal);
    document.getElementById("modal-cancel-btn").addEventListener("click", closeModal);

    document.getElementById("event-modal").addEventListener("click", (e) => {
        if (e.target.id === "event-modal") closeModal();
    });

    // モーダル内キーボード操作
    document.getElementById("event-modal").addEventListener("keydown", (e) => {
        if (e.key === "Escape") {
            e.preventDefault();
            closeModal();
            return;
        }

        // フォーカストラップ
        if (e.key === "Tab") {
            trapFocusInModal(e);
        }
    });

    document.getElementById("modal-save-btn").addEventListener("click", handleSaveEvent);
    document.getElementById("modal-save-next-btn").addEventListener("click", () => handleSaveEvent(true));
    document.getElementById("modal-delete-btn").addEventListener("click", _showDeleteConfirm);

    // インライン削除確認ボタン
    document.getElementById("modal-delete-yes").addEventListener("click", handleDeleteEvent);
    document.getElementById("modal-delete-no").addEventListener("click", _hideDeleteConfirm);
}

// ---- カテゴリプレビュー更新 ----
function _updateCategoryPreview(catName) {
    const def = getCategoryDef(catName);
    const accent = document.getElementById("modal-accent-bar");
    const dot = document.getElementById("cat-preview-dot");
    if (accent) accent.style.background = def.color;
    if (dot) dot.style.background = def.color;
}

// ---- インライン削除確認 ----
function _showDeleteConfirm() {
    document.getElementById("modal-footer-actions").style.display = "none";
    document.querySelector(".modal-footer-left").style.display = "none";
    const confirm = document.getElementById("modal-delete-confirm");
    confirm.style.display = "flex";
    document.getElementById("modal-delete-no").focus();
}

function _hideDeleteConfirm() {
    document.getElementById("modal-delete-confirm").style.display = "none";
    document.getElementById("modal-footer-actions").style.display = "flex";
    document.querySelector(".modal-footer-left").style.display = "flex";
    document.getElementById("modal-delete-btn").focus();
}

// ---- フォーカストラップ ----
function trapFocusInModal(e) {
    const modal = document.querySelector(".modal-content");
    if (!modal) return;

    const focusable = modal.querySelectorAll(
        'button:not([disabled]):not([style*="display:none"]):not([style*="display: none"]), ' +
        'input:not([disabled]), select:not([disabled]), textarea:not([disabled]), ' +
        '[tabindex]:not([tabindex="-1"])'
    );

    if (focusable.length === 0) return;

    const first = focusable[0];
    const last = focusable[focusable.length - 1];

    if (e.shiftKey) {
        if (document.activeElement === first) {
            e.preventDefault();
            last.focus();
        }
    } else {
        if (document.activeElement === last) {
            e.preventDefault();
            first.focus();
        }
    }
}

// ---- モーダル操作 ----
function openEventModal(event, categoryName, startDate, endDate) {
    _editingEvent = event;
    _lastFocusBeforeModal = document.activeElement;

    const modal = document.getElementById("event-modal");
    const titleEl = document.getElementById("modal-title");
    const deleteBtn = document.getElementById("modal-delete-btn");

    // インラインエラーをクリア
    _clearFormErrors();

    // 削除確認パネルをリセット
    _hideDeleteConfirm();

    const saveNextBtn = document.getElementById("modal-save-next-btn");

    if (event) {
        titleEl.textContent = "イベント編集";
        document.getElementById("evt-title").value = event.title || "";
        const catVal = (event.categories && event.categories[0]) || categoryName || CATEGORIES[0].name;
        document.getElementById("evt-category").value = catVal;
        document.getElementById("evt-start").value = event.startDate || "";
        document.getElementById("evt-end").value = event.endDate || "";
        document.getElementById("evt-notes").value = event.bodyPreview || "";
        deleteBtn.style.display = "inline-block";
        saveNextBtn.style.display = "none";
        _updateCategoryPreview(catVal);
    } else {
        titleEl.textContent = "イベント追加";
        document.getElementById("evt-title").value = "";
        const catVal = categoryName || CATEGORIES[0].name;
        document.getElementById("evt-category").value = catVal;
        document.getElementById("evt-start").value = startDate || "";
        document.getElementById("evt-end").value = endDate || "";
        document.getElementById("evt-notes").value = "";
        deleteBtn.style.display = "none";
        saveNextBtn.style.display = "inline-block";
        _updateCategoryPreview(catVal);
    }

    // 開始日min連動
    const startInput = document.getElementById("evt-start");
    const endInput = document.getElementById("evt-end");
    if (startInput.value) {
        endInput.min = startInput.value;
    } else {
        endInput.min = "";
    }

    modal.classList.remove("closing");
    modal.classList.add("active");
    modal.setAttribute("aria-hidden", "false");
    document.body.style.overflow = "hidden";
    document.getElementById("evt-title").focus();
}

function closeModal() {
    const modal = document.getElementById("event-modal");
    if (!modal.classList.contains("active")) return;

    // 閉じアニメーション
    modal.classList.add("closing");
    const onEnd = () => {
        modal.classList.remove("active", "closing");
        modal.setAttribute("aria-hidden", "true");
        document.body.style.overflow = "";
        _editingEvent = null;

        // フォーカス復元
        if (_lastFocusBeforeModal && _lastFocusBeforeModal.focus) {
            _lastFocusBeforeModal.focus();
            _lastFocusBeforeModal = null;
        }
        modal.removeEventListener("animationend", onEnd);
    };

    // reduced-motion の場合アニメーションが無効なので即座に閉じる
    const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    if (reducedMotion) {
        onEnd();
    } else {
        modal.addEventListener("animationend", onEnd, { once: true });
        // フォールバック: 500ms以内にanimationendが来なければ強制閉じ
        setTimeout(() => {
            if (modal.classList.contains("closing")) onEnd();
        }, 500);
    }
}

// ========================================
// 保存処理（モーダル経由）
// ========================================
// ---- インラインバリデーション ----
function _setFieldError(fieldName, message) {
    const fg = document.querySelector(`.form-group[data-field="${fieldName}"]`);
    if (!fg) return;
    fg.classList.add("has-error");
    const errSpan = fg.querySelector(".form-error");
    if (errSpan) errSpan.textContent = message;
}

function _clearFormErrors() {
    document.querySelectorAll("#event-modal .form-group.has-error").forEach(fg => {
        fg.classList.remove("has-error");
        const errSpan = fg.querySelector(".form-error");
        if (errSpan) errSpan.textContent = "";
    });
}

async function handleSaveEvent(continueAdding) {
    const title = document.getElementById("evt-title").value.trim();
    const category = document.getElementById("evt-category").value;
    const startDate = document.getElementById("evt-start").value;
    const endDate = document.getElementById("evt-end").value;
    const notes = document.getElementById("evt-notes").value.trim();

    // エラーをクリアしてからバリデーション
    _clearFormErrors();
    let hasError = false;
    let firstErrorField = null;

    if (!title) {
        _setFieldError("title", "イベント名を入力してください");
        if (!firstErrorField) firstErrorField = "evt-title";
        hasError = true;
    }
    if (!startDate) {
        _setFieldError("start", "開始日を入力してください");
        if (!firstErrorField) firstErrorField = "evt-start";
        hasError = true;
    }
    if (!endDate) {
        _setFieldError("end", "終了日を入力してください");
        if (!firstErrorField) firstErrorField = "evt-end";
        hasError = true;
    }
    if (startDate && endDate && startDate > endDate) {
        _setFieldError("end", "終了日は開始日以降にしてください");
        if (!firstErrorField) firstErrorField = "evt-end";
        hasError = true;
    }

    if (hasError) {
        if (firstErrorField) document.getElementById(firstErrorField).focus();
        return;
    }

    const saveBtn = document.getElementById("modal-save-btn");
    const saveNextBtn = document.getElementById("modal-save-next-btn");
    saveBtn.disabled = true;
    saveNextBtn.disabled = true;
    saveBtn.textContent = "保存中...";

    try {
        const token = await getAccessToken();
        const eventData = { title, category, startDate, endDate, notes };

        if (_editingEvent && _editingEvent.id) {
            // --- 更新：楽観的にローカルキャッシュを変更 ---
            await updateCalendarEvent(token, _editingEvent.id, eventData);
            const cached = _cachedGraphEvents.find(e => e.id === _editingEvent.id);
            if (cached) {
                cached.title = title;
                cached.categories = [category];
                cached.startDate = startDate;
                cached.endDate = endDate;
                cached.bodyPreview = notes;
            }
            closeModal();
            rerenderFromCache();
            announceStatus(`「${title}」を更新しました`);
        } else {
            // --- 新規：API の戻り値からイベントIDを取得する必要がある ---
            const created = await createCalendarEvent(token, eventData);
            // ローカルキャッシュに追加
            _cachedGraphEvents.push({
                id: created?.id || "temp-" + Date.now(),
                title: title,
                categories: [category],
                startDate: startDate,
                endDate: endDate,
                bodyPreview: notes,
            });

            if (continueAdding === true) {
                // フォームをリセットしてカテゴリ・日付は保持
                document.getElementById("evt-title").value = "";
                document.getElementById("evt-notes").value = "";
                _clearFormErrors();
                _editingEvent = null;
                rerenderFromCache();
                announceStatus(`「${title}」を追加しました — 続けて入力できます`);
                document.getElementById("evt-title").focus();
            } else {
                closeModal();
                rerenderFromCache();
                announceStatus(`「${title}」を追加しました`);
            }
        }
    } catch (error) {
        console.error("Save event failed:", error);
        _setFieldError("title", "保存に失敗しました: " + error.message);
    } finally {
        saveBtn.disabled = false;
        saveNextBtn.disabled = false;
        saveBtn.textContent = "保存";
    }
}

// ========================================
// 削除処理
// ========================================
async function handleDeleteEvent() {
    if (!_editingEvent || !_editingEvent.id) return;

    const yesBtn = document.getElementById("modal-delete-yes");
    yesBtn.disabled = true;
    yesBtn.textContent = "削除中...";

    try {
        const token = await getAccessToken();
        const deletedTitle = _editingEvent.title;
        await deleteCalendarEvent(token, _editingEvent.id);

        // ローカルキャッシュから削除
        _cachedGraphEvents = _cachedGraphEvents.filter(e => e.id !== _editingEvent.id);
        closeModal();
        rerenderFromCache();
        announceStatus(`「${deletedTitle}」を削除しました`);
    } catch (error) {
        console.error("Delete event failed:", error);
        _hideDeleteConfirm();
        _setFieldError("title", "削除に失敗しました: " + error.message);
    } finally {
        yesBtn.disabled = false;
        yesBtn.textContent = "削除する";
    }
}

// ========================================
// ドラッグ操作コールバック（SPA楽観的更新）
// ========================================

// 範囲選択で新規作成 → モーダルを開く（まだAPIは叩かない）
function onDragCreate(categoryName, startDate, endDate) {
    openEventModal(null, categoryName, startDate, endDate);
}

// バー移動 or リサイズ → 即座にローカル反映、API はバックグラウンド
async function onDragMoveOrResize(eventId, eventObj, newStartDate, newEndDate) {
    if (!eventId || !eventObj) return;
    if (eventObj.startDate === newStartDate && eventObj.endDate === newEndDate) return;

    // 旧日付を退避（ロールバック用）
    const oldStartDate = eventObj.startDate;
    const oldEndDate = eventObj.endDate;

    // ① ローカルキャッシュを即時更新
    const cached = _cachedGraphEvents.find(e => e.id === eventId);
    if (cached) {
        cached.startDate = newStartDate;
        cached.endDate = newEndDate;
    }

    // ② スクロール位置を保ったまま即時再描画（API待ちなし）
    rerenderFromCache();
    announceStatus(`「${eventObj.title}」の日程を変更しました`);

    // ③ バックグラウンドで API に送信
    try {
        const token = await getAccessToken();
        const category = (eventObj.categories && eventObj.categories[0]) || null;
        await updateCalendarEvent(token, eventId, {
            title: eventObj.title,
            category: category,
            startDate: newStartDate,
            endDate: newEndDate,
            notes: eventObj.bodyPreview || "",
        });
        // 成功 — 何もしない（ローカルは既に最新）
    } catch (error) {
        console.error("Drag update failed:", error);
        // ④ 失敗 → ローカルを巻き戻して再描画
        if (cached) {
            cached.startDate = oldStartDate;
            cached.endDate = oldEndDate;
        }
        rerenderFromCache();
        announceStatus("更新に失敗しました。元に戻しました。");
        alert("更新に失敗しました: " + error.message);
    }
}

// ========================================
// 認証
// ========================================
async function handleSignIn() {
    try {
        await signIn();
        document.getElementById("header-sign-in-btn").style.display = "none";
        _rawConfig = null; // サインイン後にOutlookから再取得
        await loadCalendar();
    } catch (error) {
        if (error.message === "popup_blocked") {
            showStatus("error", "ポップアップがブロックされました。ブラウザの設定でポップアップを許可してください。");
        } else if (error.message === "user_cancelled") {
            // キャンセル
        } else {
            showStatus("error", "サインインに失敗しました: " + error.message);
        }
    }
}

async function handleSignOut() {
    try { await signOut(); } catch (e) { /* ignore */ }
    _cachedGraphEvents = [];
    _cachedHolidays = [];
    _cachedHolidaySet = null;
    _rawConfig = null;
    _configEventId = null;
    document.getElementById("timeline-container").innerHTML = "";
    document.getElementById("user-name").textContent = "";
    document.getElementById("sign-out-btn").style.display = "none";
    document.getElementById("add-event-btn").style.display = "none";
    document.getElementById("settings-btn").style.display = "none";
    document.getElementById("export-btn").style.display = "none";
    document.getElementById("legend").innerHTML = "";
    document.getElementById("month-nav").style.display = "none";
    document.getElementById("timeline-mode-bar").style.display = "none";
    // サインアウト後は公開ビューに戻る
    loadPublicCalendar();
    announceStatus("サインアウトしました");
}

// ========================================
// 公開ビュー（未サインイン時）
// ========================================
async function loadPublicCalendar() {
    const year = currentYear;

    // デフォルトカテゴリを使用
    if (!_rawConfig) {
        _rawConfig = {
            version: 3,
            yearCategories: {},
        };
    }
    if (!_rawConfig.yearCategories[String(year)]) {
        _rawConfig.yearCategories[String(year)] = DEFAULT_CATEGORIES.map(c => ({ id: c.id, name: c.name, color: c.color }));
    }
    CATEGORIES = _buildCategoriesForYear(_rawConfig, year);
    populateCategorySelect();

    // 祝日（ローカルデータ）
    const holidays = getJapaneseHolidays(year);
    const holidaySet = buildHolidaySet(holidays);

    // 公開イベントデータを読み込み（localStorage → data/events.json の優先順）
    let publicEvents = [];
    try {
        // まずlocalStorageからキャッシュを試行
        const cached = localStorage.getItem(`gyro_events_${year}`);
        if (cached) {
            const parsed = JSON.parse(cached);
            publicEvents = parsed.events || [];
            if (parsed.categories && parsed.categories.length > 0) {
                _rawConfig.yearCategories[String(year)] = parsed.categories;
                CATEGORIES = _buildCategoriesForYear(_rawConfig, year);
                populateCategorySelect();
            }
        }
    } catch (e) {
        console.warn("localStorageキャッシュ読み込み失敗:", e);
    }

    // localStorageに無ければ data/events.json を試行
    if (publicEvents.length === 0) {
        try {
            const res = await fetch("data/events.json");
            if (res.ok) {
                const data = await res.json();
                const yearData = data.years?.[String(year)];
                if (yearData) {
                    publicEvents = yearData.events || [];
                    if (yearData.categories && yearData.categories.length > 0) {
                        _rawConfig.yearCategories[String(year)] = yearData.categories;
                        CATEGORIES = _buildCategoriesForYear(_rawConfig, year);
                        populateCategorySelect();
                    }
                }
            }
        } catch (e) {
            console.warn("公開イベントデータの読み込みに失敗:", e);
        }
    }

    // イベントからカテゴリを自動検出してマージ
    _mergeEventCategories(publicEvents);

    // キャッシュ
    _cachedGraphEvents = publicEvents;
    _cachedHolidays = holidays;
    _cachedHolidaySet = holidaySet;

    hideStatus();
    renderLegend();

    // タイムラインモードバー表示
    document.getElementById("timeline-mode-bar").style.display =
        (_currentView === "timeline") ? "flex" : "none";

    rerenderFromCache(true);

    // ヘッダーにサインインボタンを表示
    document.getElementById("header-sign-in-btn").style.display = "inline-block";

    const eventCount = publicEvents.length;
    const suffix = eventCount > 0 ? `（${eventCount}件）` : "（閲覧モード）";
    announceStatus(`${year}年のスケジュールを表示しています${suffix}`);
}

// ========================================
// カレンダー読み込み（初回 & 年変更時のみAPI通信）
// ========================================
async function loadCalendar() {
    showStatus("loading");

    try {
        const token = await getAccessToken();
        const year = currentYear;

        // [DEBUG] _rawConfig の状態を確認
        if (_rawConfig && _rawConfig.yearCategories) {
            const summary = Object.entries(_rawConfig.yearCategories).map(([y, cats]) => `${y}:${cats.length}件`).join(", ");
            console.log(`[loadCalendar] year=${year}, _rawConfig exists: {${summary}}`);
        } else {
            console.log(`[loadCalendar] year=${year}, _rawConfig=${_rawConfig}`);
        }

        // ① カテゴリ設定をOutlookから読み込む（v3形式）
        // 初回のみAPIから取得。以降はメモリ上の_rawConfigを使用
        // （年変更のたびに再取得すると、保存直後のデータが反映されない場合がある）
        if (!_rawConfig) {
            try {
                const config = await fetchCategoryConfig(token);
                if (config && config.rawConfig) {
                    _rawConfig = config.rawConfig;
                    _configEventId = config.configEventId;
                } else {
                    // Config Event が存在しない → デフォルトで作成（バックグラウンド）
                    _rawConfig = {
                        version: 3,
                        yearCategories: {},
                    };
                    _createDefaultConfigEvent(token);
                }
            } catch (configErr) {
                console.warn("カテゴリ設定の読み込みに失敗（デフォルトを使用）:", configErr);
                _rawConfig = {
                    version: 3,
                    yearCategories: {},
                };
            }
        }
        // 当年のカテゴリが無ければデフォルトをセット
        if (!_rawConfig.yearCategories[String(year)]) {
            _rawConfig.yearCategories[String(year)] = DEFAULT_CATEGORIES.map(c => ({ id: c.id, name: c.name, color: c.color }));
        }
        CATEGORIES = _buildCategoriesForYear(_rawConfig, year);
        // カテゴリselectを更新
        populateCategorySelect();

        // ② カレンダーイベント取得
        const holidays = getJapaneseHolidays(year);
        const holidaySet = buildHolidaySet(holidays);
        const graphEvents = await fetchCalendarEvents(token, year);

        // イベントからカテゴリを自動検出してマージ
        _mergeEventCategories(graphEvents);
        populateCategorySelect();

        // キャッシュを更新
        _cachedGraphEvents = graphEvents;
        _cachedHolidays = holidays;
        _cachedHolidaySet = holidaySet;

        // localStorageに公開用キャッシュを保存（未サインイン時に使用）
        _saveEventsToLocalStorage(year, graphEvents, CATEGORIES);

        hideStatus();
        renderLegend();

        // 凡例に管理ボタンを追加
        _appendCategoryManageButton();

        // タイムラインモードバーを表示（タイムラインビュー時のみ）
        document.getElementById("timeline-mode-bar").style.display =
            (_currentView === "timeline") ? "flex" : "none";

        // 現在のビューに応じて描画（初回ロードフラグ付き）
        rerenderFromCache(true);

        const account = getActiveAccount();
        if (account) {
            document.getElementById("user-name").textContent = account.name || account.username;
            document.getElementById("sign-out-btn").style.display = "inline-block";
            document.getElementById("add-event-btn").style.display = "flex";
            document.getElementById("settings-btn").style.display = "flex";
            document.getElementById("export-btn").style.display = "flex";
        }

        announceStatus(`${year}年のスケジュールを読み込みました（${graphEvents.length}件）`);
    } catch (error) {
        console.error("Calendar load failed:", error);
        if (error.message?.includes("consent") || error.message?.includes("interaction")) {
            showStatus("error", "アクセス許可が必要です。管理者にお問い合わせください。");
        } else {
            showStatus("error", error.message || "カレンダーの取得に失敗しました。");
        }
    }
}

// ---- デフォルトConfig Eventをバックグラウンドで作成 ----
async function _createDefaultConfigEvent(token) {
    try {
        const result = await saveCategoryConfig(token, _rawConfig, null);
        _configEventId = result?.id || null;
        console.log("デフォルトConfig Eventを作成しました（v3形式）");
    } catch (e) {
        console.warn("Config Event作成失敗:", e);
    }
}

// ---- 凡例に管理ボタン追加 ----
function _appendCategoryManageButton() {
    const legend = document.getElementById("legend");
    // 既存の管理ボタンを削除（年変更時の重複防止）
    const existing = legend.querySelector(".legend-manage-btn");
    if (existing) existing.remove();

    const btn = document.createElement("button");
    btn.className = "legend-manage-btn";
    btn.textContent = "\u2699 管理";
    btn.title = "カテゴリ管理";
    btn.setAttribute("aria-label", "カテゴリを管理");
    btn.addEventListener("click", openCategoryManager);
    legend.appendChild(btn);
}

// ---- 年変更 ----
function changeYear(delta) {
    currentYear += delta;
    graphConfig.year = currentYear;
    updateYearLabel();
    updateMonthLabel();
    if (getActiveAccount()) loadCalendar(); else loadPublicCalendar();
}

function updateYearLabel() {
    const el = document.getElementById("year-label");
    if (el) el.textContent = currentYear + "年";
    document.getElementById("today-btn").textContent = currentYear + "年";
}

// ---- ステータス表示 ----
function showStatus(type, message) {
    const area = document.getElementById("status-area");
    const spinner = document.getElementById("loading-spinner");
    const errorEl = document.getElementById("error-message");
    const signInBtn = document.getElementById("sign-in-btn");

    spinner.style.display = "none";
    errorEl.style.display = "none";
    signInBtn.style.display = "none";
    area.style.display = "flex";

    switch (type) {
        case "loading":
            spinner.style.display = "flex";
            if (message) spinner.querySelector("span").textContent = message;
            break;
        case "error":
            errorEl.style.display = "block";
            errorEl.textContent = message;
            break;
        case "auth":
            signInBtn.style.display = "inline-block";
            break;
    }
}

function hideStatus() {
    document.getElementById("status-area").style.display = "none";
}

// ========================================
// localStorageキャッシュ（サインイン時に自動保存）
// ========================================
function _saveEventsToLocalStorage(year, events, categories) {
    try {
        const data = {
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
            savedAt: new Date().toISOString(),
        };
        localStorage.setItem(`gyro_events_${year}`, JSON.stringify(data));
    } catch (e) {
        console.warn("localStorageへの保存失敗:", e);
    }
}

// ========================================
// 公開データ書き出し
// ========================================
async function handleExportPublicData() {
    if (_cachedGraphEvents.length === 0) {
        alert("書き出すイベントがありません。");
        return;
    }

    const exportBtn = document.getElementById("export-btn");
    exportBtn.disabled = true;

    try {
        // 既存のevents.jsonを読み込んで他年度のデータを保持
        let existingData = { lastUpdated: null, years: {} };
        try {
            const res = await fetch("data/events.json");
            if (res.ok) existingData = await res.json();
        } catch (e) { /* 無視 */ }

        // 現在の年度のイベント+カテゴリを書き出し
        const year = String(currentYear);
        existingData.years[year] = {
            events: _cachedGraphEvents.map(e => ({
                id: e.id,
                title: e.title,
                startDate: e.startDate,
                endDate: e.endDate,
                categories: e.categories || [],
                bodyPreview: e.bodyPreview || "",
            })),
            categories: CATEGORIES.map(c => ({
                id: c.id,
                name: c.name,
                color: c.color,
            })),
        };
        existingData.lastUpdated = new Date().toISOString();

        // JSONファイルをダウンロード
        const json = JSON.stringify(existingData, null, 2);
        const blob = new Blob([json], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "events.json";
        a.click();
        URL.revokeObjectURL(url);

        announceStatus("公開用データをダウンロードしました。data/events.json として配置してください。");
        alert("events.json をダウンロードしました。\nリポジトリの data/events.json に配置してコミットすると、サインインなしでもイベントが表示されます。");
    } catch (error) {
        console.error("Export failed:", error);
        alert("書き出しに失敗しました: " + error.message);
    } finally {
        exportBtn.disabled = false;
    }
}

// カテゴリ管理モーダルは category-manager.js に分割済み
