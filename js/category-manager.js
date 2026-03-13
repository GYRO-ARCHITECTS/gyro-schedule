// カテゴリ管理モーダル
let _catManagerFocusBefore = null;
let _colorPickerTarget = null; // 現在カラーピッカーが開いている行
let _catManagerYear = null;    // モーダル内で表示中の年度

function setupCategoryManager() {
    document.getElementById("cat-manager-close").addEventListener("click", closeCategoryManager);
    document.getElementById("cat-manager-cancel").addEventListener("click", closeCategoryManager);
    document.getElementById("cat-manager-save").addEventListener("click", _saveCategoryChanges);
    document.getElementById("cat-add-btn").addEventListener("click", () => _addCategoryRow());

    // 年ナビ
    document.getElementById("cat-prev-year").addEventListener("click", () => _switchCatManagerYear(-1));
    document.getElementById("cat-next-year").addEventListener("click", () => _switchCatManagerYear(1));

    // 複製パネル
    document.getElementById("cat-duplicate-btn").addEventListener("click", _openDuplicatePanel);
    document.getElementById("cat-duplicate-cancel").addEventListener("click", _closeDuplicatePanel);
    document.getElementById("cat-duplicate-exec").addEventListener("click", _executeDuplicate);

    // モーダル外クリックで閉じる
    document.getElementById("cat-manager-modal").addEventListener("click", (e) => {
        if (e.target.id === "cat-manager-modal") closeCategoryManager();
    });

    // Escape キー
    document.getElementById("cat-manager-modal").addEventListener("keydown", (e) => {
        if (e.key === "Escape") {
            e.preventDefault();
            _hideColorPicker();
            _closeDuplicatePanel();
            closeCategoryManager();
        }
    });

    // カラーピッカー: 外クリックで閉じる
    document.addEventListener("click", (e) => {
        const picker = document.getElementById("color-picker-popover");
        if (picker.getAttribute("aria-hidden") === "true") return;
        if (!picker.contains(e.target) && !e.target.classList.contains("cat-manager-color-btn")) {
            _hideColorPicker();
        }
    });

    // カラーピッカーのスウォッチを生成
    _initColorPicker();
}

function _initColorPicker() {
    const picker = document.getElementById("color-picker-popover");
    picker.innerHTML = "";
    COLOR_PRESETS.forEach(hex => {
        const swatch = document.createElement("button");
        swatch.type = "button";
        swatch.className = "color-swatch";
        swatch.style.background = hex;
        swatch.dataset.color = hex;
        swatch.setAttribute("aria-label", hex);
        swatch.addEventListener("click", () => {
            if (_colorPickerTarget) {
                _colorPickerTarget.style.background = hex;
                _colorPickerTarget.dataset.color = hex;
                picker.querySelectorAll(".color-swatch").forEach(s => s.classList.remove("selected"));
                swatch.classList.add("selected");
            }
            _hideColorPicker();
        });
        picker.appendChild(swatch);
    });
}

function openCategoryManager() {
    _catManagerFocusBefore = document.activeElement;
    _catManagerYear = currentYear;
    const modal = document.getElementById("cat-manager-modal");

    _updateCatManagerYearLabel();
    _renderCategoryList();
    _closeDuplicatePanel();

    modal.classList.remove("closing");
    modal.classList.add("active");
    modal.setAttribute("aria-hidden", "false");
    document.body.style.overflow = "hidden";

    const firstInput = modal.querySelector(".cat-manager-name");
    if (firstInput) firstInput.focus();
}

function closeCategoryManager() {
    const modal = document.getElementById("cat-manager-modal");
    if (!modal.classList.contains("active")) return;

    _hideColorPicker();
    _closeDuplicatePanel();

    modal.classList.add("closing");
    const onEnd = () => {
        modal.classList.remove("active", "closing");
        modal.setAttribute("aria-hidden", "true");
        document.body.style.overflow = "";

        if (_catManagerFocusBefore && _catManagerFocusBefore.focus) {
            _catManagerFocusBefore.focus();
            _catManagerFocusBefore = null;
        }
        modal.removeEventListener("animationend", onEnd);
    };

    const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    if (reducedMotion) {
        onEnd();
    } else {
        modal.addEventListener("animationend", onEnd, { once: true });
        setTimeout(() => { if (modal.classList.contains("closing")) onEnd(); }, 500);
    }
}

// ---- 年度ナビ ----
function _updateCatManagerYearLabel() {
    const label = document.getElementById("cat-year-label");
    if (label) label.textContent = `${_catManagerYear}年`;
}

function _switchCatManagerYear(delta) {
    _catManagerYear += delta;
    _updateCatManagerYearLabel();
    _renderCategoryList();
    _closeDuplicatePanel();
}

// ---- 複製パネル ----
function _openDuplicatePanel() {
    const panel = document.getElementById("cat-duplicate-panel");
    const select = document.getElementById("cat-duplicate-source");
    select.innerHTML = "";

    if (!_rawConfig || !_rawConfig.yearCategories) {
        panel.style.display = "none";
        return;
    }

    // 複製元の候補: 現在モーダル表示年以外でカテゴリが存在する年度
    const years = Object.keys(_rawConfig.yearCategories)
        .filter(y => y !== String(_catManagerYear) && _rawConfig.yearCategories[y].length > 0)
        .sort((a, b) => Number(b) - Number(a));

    if (years.length === 0) {
        // 複製元がない
        const opt = document.createElement("option");
        opt.textContent = "複製元の年度がありません";
        opt.disabled = true;
        select.appendChild(opt);
        document.getElementById("cat-duplicate-exec").disabled = true;
    } else {
        years.forEach(y => {
            const opt = document.createElement("option");
            opt.value = y;
            opt.textContent = `${y}年（${_rawConfig.yearCategories[y].length}件）`;
            select.appendChild(opt);
        });
        document.getElementById("cat-duplicate-exec").disabled = false;
    }

    panel.style.display = "flex";
}

function _closeDuplicatePanel() {
    document.getElementById("cat-duplicate-panel").style.display = "none";
}

function _executeDuplicate() {
    const sourceYear = document.getElementById("cat-duplicate-source").value;
    if (!sourceYear || !_rawConfig?.yearCategories?.[sourceYear]) return;

    const sourceCats = _rawConfig.yearCategories[sourceYear];
    const list = document.getElementById("cat-list");
    list.innerHTML = "";

    // 複製元のカテゴリを新しいIDでコピーしてUIに追加
    sourceCats.forEach(cat => {
        const newId = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
        const row = _createCategoryRow(newId, cat.name, cat.color);
        list.appendChild(row);
    });

    _closeDuplicatePanel();
    announceStatus(`${sourceYear}年のカテゴリを複製しました`);

    const firstInput = list.querySelector(".cat-manager-name");
    if (firstInput) firstInput.focus();
}

// ---- カテゴリリスト描画（v3: 単一リスト） ----
function _renderCategoryList() {
    const list = document.getElementById("cat-list");
    list.innerHTML = "";

    if (!_rawConfig) return;

    const year = String(_catManagerYear);
    const cats = _rawConfig.yearCategories?.[year] || [];

    if (cats.length === 0) {
        // 空の年度: ヒントを表示
        const hint = document.createElement("div");
        hint.className = "cat-empty-hint";
        hint.textContent = "この年度にはカテゴリがありません。追加するか、他の年から複製してください。";
        list.appendChild(hint);
        return;
    }

    cats.forEach(cat => {
        const row = _createCategoryRow(cat.id, cat.name, cat.color);
        list.appendChild(row);
    });
}

function _createCategoryRow(id, name, color) {
    const row = document.createElement("div");
    row.className = "cat-manager-row";
    row.dataset.catId = id || Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

    const colorBtn = document.createElement("div");
    colorBtn.className = "cat-manager-color-btn";
    colorBtn.style.background = color;
    colorBtn.dataset.color = color;
    colorBtn.tabIndex = 0;
    colorBtn.setAttribute("role", "button");
    colorBtn.setAttribute("aria-label", "色を選択");
    colorBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        _showColorPicker(colorBtn);
    });
    colorBtn.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            e.stopPropagation();
            _showColorPicker(colorBtn);
        }
    });

    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.className = "cat-manager-name";
    nameInput.value = name;
    nameInput.placeholder = "カテゴリ名を入力";
    nameInput.addEventListener("input", () => {
        if (row.classList.contains("has-error")) {
            row.classList.remove("has-error");
            const errSpan = row.querySelector(".form-error");
            if (errSpan) errSpan.remove();
        }
    });

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "cat-manager-remove";
    removeBtn.setAttribute("aria-label", "削除");
    removeBtn.title = "削除";
    removeBtn.textContent = "\u2715";
    removeBtn.addEventListener("click", () => { row.remove(); });

    row.appendChild(colorBtn);
    row.appendChild(nameInput);
    row.appendChild(removeBtn);
    return row;
}

function _addCategoryRow() {
    const list = document.getElementById("cat-list");

    // 空ヒントがあれば除去
    const hint = list.querySelector(".cat-empty-hint");
    if (hint) hint.remove();

    const usedColors = _collectUsedColors();
    const nextColor = COLOR_PRESETS.find(c => !usedColors.has(c)) || COLOR_PRESETS[0];

    const row = _createCategoryRow(null, "", nextColor);
    list.appendChild(row);
    row.scrollIntoView({ behavior: "smooth", block: "nearest" });
    row.querySelector(".cat-manager-name").focus();
}

function _collectUsedColors() {
    const usedColors = new Set();
    document.querySelectorAll("#cat-list .cat-manager-color-btn").forEach(btn => {
        usedColors.add(btn.dataset.color);
    });
    return usedColors;
}

function _showColorPicker(colorBtn) {
    const picker = document.getElementById("color-picker-popover");
    _colorPickerTarget = colorBtn;

    const currentColor = colorBtn.dataset.color;
    picker.querySelectorAll(".color-swatch").forEach(s => {
        s.classList.toggle("selected", s.dataset.color === currentColor);
    });

    const rect = colorBtn.getBoundingClientRect();
    picker.style.top = (rect.bottom + 6) + "px";
    picker.style.left = rect.left + "px";
    picker.setAttribute("aria-hidden", "false");

    requestAnimationFrame(() => {
        const pickerRect = picker.getBoundingClientRect();
        if (pickerRect.right > window.innerWidth - 8) {
            picker.style.left = Math.max(8, window.innerWidth - pickerRect.width - 8) + "px";
        }
        if (pickerRect.bottom > window.innerHeight - 8) {
            picker.style.top = (rect.top - pickerRect.height - 6) + "px";
        }
    });
}

function _hideColorPicker() {
    const picker = document.getElementById("color-picker-popover");
    picker.setAttribute("aria-hidden", "true");
    _colorPickerTarget = null;
}

async function _saveCategoryChanges() {
    const list = document.getElementById("cat-list");
    const rows = list.querySelectorAll(".cat-manager-row");

    // バリデーション
    let hasError = false;
    const newCategories = [];
    const namesSeen = new Set();

    rows.forEach(row => {
        row.classList.remove("has-error");
        const existingErr = row.querySelector(".form-error");
        if (existingErr) existingErr.remove();

        const name = row.querySelector(".cat-manager-name").value.trim();
        const color = row.querySelector(".cat-manager-color-btn").dataset.color;
        const id = row.dataset.catId;

        if (!name) {
            row.classList.add("has-error");
            const errSpan = document.createElement("span");
            errSpan.className = "form-error";
            errSpan.textContent = "カテゴリ名を入力してください";
            row.appendChild(errSpan);
            hasError = true;
            return;
        }

        if (namesSeen.has(name)) {
            row.classList.add("has-error");
            const errSpan = document.createElement("span");
            errSpan.className = "form-error";
            errSpan.textContent = "カテゴリ名が重複しています";
            row.appendChild(errSpan);
            hasError = true;
            return;
        }

        namesSeen.add(name);
        newCategories.push({ id, name, color });
    });

    if (hasError) {
        const firstErr = list.querySelector(".has-error .cat-manager-name");
        if (firstErr) firstErr.focus();
        return;
    }

    // 空でもOK（その年のカテゴリを全削除可能）

    // 保存処理
    const saveBtn = document.getElementById("cat-manager-save");
    saveBtn.disabled = true;
    saveBtn.textContent = "保存中...";

    try {
        const token = await getAccessToken();
        const savingYear = String(_catManagerYear);

        // _rawConfig を更新
        if (!_rawConfig.yearCategories) _rawConfig.yearCategories = {};
        _rawConfig.yearCategories[savingYear] = newCategories;

        // [DEBUG]
        const summary = Object.entries(_rawConfig.yearCategories).map(([y, cats]) => `${y}:${cats.length}件`).join(", ");
        console.log(`[saveCategoryChanges] saving year=${savingYear}, _rawConfig: {${summary}}`);

        const result = await saveCategoryConfig(token, _rawConfig, _configEventId);

        if (!_configEventId && result?.id) {
            _configEventId = result.id;
        }

        // 常に現在の表示年度のカテゴリを再構築して UI を更新
        CATEGORIES = _buildCategoriesForYear(_rawConfig, currentYear);
        populateCategorySelect();
        renderLegend();
        _appendCategoryManageButton();
        rerenderFromCache();

        closeCategoryManager();
        announceStatus(`${savingYear}年のカテゴリ設定を保存しました`);
    } catch (error) {
        console.error("カテゴリ保存失敗:", error);
        const desc = document.querySelector(".cat-manager-desc");
        if (desc) {
            desc.textContent = "保存に失敗しました: " + error.message;
            desc.style.color = "#ef4444";
            setTimeout(() => {
                desc.textContent = "年度ごとにカテゴリを管理できます。変更は全ユーザーに反映されます。";
                desc.style.color = "";
            }, 4000);
        }
    } finally {
        saveBtn.disabled = false;
        saveBtn.textContent = "保存";
    }
}
