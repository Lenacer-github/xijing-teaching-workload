(() => {
    "use strict";

    const config = window.APP_CONFIG;
    const catalog = config.catalog;
    const state = {
        dept: config.departments[0] || "",
        name: "",
        title: "",
        term: config.defaultTerm,
        rows: [],
    };

    const $ = (selector, root = document) => root.querySelector(selector);
    const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];
    const deptInput = $("#department");
    const nameInput = $("#teacher-name");
    const titleInput = $("#teacher-title");
    const termInput = $("#academic-term");
    const recordList = $("#record-list");
    const emptyState = $("#empty-state");
    const bottomAddWrap = $("#bottom-add-wrap");
    const totalWorkload = $("#total-workload");
    const restoreFile = $("#restore-file");
    const recordTemplate = $("#record-template");
    let exportInFlight = false;

    const makeId = () =>
        globalThis.crypto?.randomUUID?.()
        || `row-${Date.now()}-${Math.random().toString(16).slice(2)}`;

    function formatNumber(value) {
        const number = Number(value);
        if (!Number.isFinite(number)) return "0";
        return Number.isInteger(number)
            ? String(number)
            : number.toFixed(2).replace(/0+$/, "").replace(/\.$/, "");
    }

    function showToast(message, isError = false) {
        const toastElement = $("#app-toast");
        toastElement.classList.toggle("toast-error", isError);
        $(".toast-body", toastElement).textContent = message;
        bootstrap.Toast.getOrCreateInstance(toastElement, { delay: 2800 }).show();
    }

    function populateSelect(select, values) {
        select.replaceChildren(...values.map((value) => {
            const option = document.createElement("option");
            option.value = value;
            option.textContent = value;
            return option;
        }));
    }

    function syncBasicForm() {
        deptInput.value = state.dept;
        nameInput.value = state.name;
        titleInput.value = state.title;
        termInput.value = state.term;
    }

    function payload() {
        return {
            dept: state.dept,
            name: state.name.trim(),
            title: state.title.trim(),
            term: state.term,
            rows: state.rows.map(({ id, ...row }) => ({ ...row })),
        };
    }

    function addRecord() {
        const row = {
            id: makeId(),
            category: "",
            standard: "",
            workload: "",
            time: "",
            remark: "",
            reviewer: "",
        };
        state.rows.push(row);
        renderRecords();
        requestAnimationFrame(() => {
            const card = document.querySelector(`[data-row-id="${row.id}"]`);
            card?.scrollIntoView({ behavior: "smooth", block: "center" });
            $(".category-input", card)?.focus({ preventScroll: true });
        });
    }

    function deleteRecord(id) {
        state.rows = state.rows.filter((row) => row.id !== id);
        renderRecords();
        showToast("已删除该条记录");
    }

    function standardOptions(row) {
        return catalog[row.category] || [];
    }

    function selectedStandard(row) {
        return standardOptions(row).find((item) => item.value === row.standard);
    }

    function renderStandardMenu(card, row, keyword = "") {
        const menu = $(".standard-menu", card);
        const normalized = keyword.trim().toLowerCase();
        const options = standardOptions(row).filter((item) =>
            item.text.toLowerCase().includes(normalized)
        );
        menu.replaceChildren();

        if (!row.category) {
            menu.innerHTML = '<div class="standard-empty">请先选择项目类别</div>';
        } else if (!options.length) {
            menu.innerHTML = '<div class="standard-empty">未找到匹配结果</div>';
        } else {
            options.forEach((item) => {
                const button = document.createElement("button");
                button.type = "button";
                button.className = "standard-option";
                const text = document.createElement("span");
                text.textContent = item.text;
                button.append(text);
                button.addEventListener("mousedown", (event) => {
                    event.preventDefault();
                    row.standard = item.value;
                    row.reviewer = item.owner;
                    $(".standard-input", card).value = item.text;
                    updateReviewer(card, row.reviewer);
                    menu.classList.remove("show");
                    clearInvalid($(".standard-input", card));
                    markDuplicates();
                });
                menu.append(button);
            });
        }
        menu.classList.add("show");
    }

    function updateReviewer(card, reviewer) {
        const pill = $(".reviewer-pill", card);
        pill.textContent = reviewer || "待选择";
        pill.classList.toggle("long", (reviewer || "").length > 18);
        pill.title = reviewer || "";
    }

    function bindRecord(card, row, index) {
        card.dataset.rowId = row.id;
        $(".record-number", card).textContent = `记录 ${index + 1}`;
        updateReviewer(card, row.reviewer);

        const category = $(".category-input", card);
        populateSelect(category, ["", ...Object.keys(catalog)]);
        category.options[0].textContent = "请选择项目类别";
        category.value = row.category;
        category.addEventListener("change", () => {
            row.category = category.value;
            row.standard = "";
            row.reviewer = "";
            const standard = $(".standard-input", card);
            standard.disabled = !row.category;
            standard.placeholder = row.category ? "输入关键词检索计算标准" : "请先选择项目类别";
            standard.value = "";
            updateReviewer(card, "");
            clearInvalid(category);
            markDuplicates();
        });

        const standard = $(".standard-input", card);
        standard.disabled = !row.category;
        standard.placeholder = row.category ? "输入关键词检索计算标准" : "请先选择项目类别";
        standard.value = selectedStandard(row)?.text || "";
        standard.addEventListener("focus", () => renderStandardMenu(card, row, standard.value));
        standard.addEventListener("input", () => {
            if (selectedStandard(row)?.text !== standard.value) {
                row.standard = "";
                row.reviewer = "";
                updateReviewer(card, "");
            }
            renderStandardMenu(card, row, standard.value);
        });
        standard.addEventListener("keydown", (event) => {
            if (event.key === "Escape") $(".standard-menu", card).classList.remove("show");
        });

        const workload = $(".workload-input", card);
        workload.value = row.workload;
        workload.addEventListener("input", () => {
            row.workload = workload.value;
            clearInvalid(workload);
            updateTotal();
        });

        const date = $(".date-input", card);
        date.value = row.time;
        date.addEventListener("input", () => {
            row.time = date.value;
            clearInvalid(date);
        });

        const remark = $(".remark-input", card);
        remark.value = row.remark;
        remark.addEventListener("input", () => {
            row.remark = remark.value;
            clearInvalid(remark);
        });
        $(".delete-record", card).addEventListener("click", () => deleteRecord(row.id));
    }

    function renderRecords() {
        recordList.replaceChildren();
        state.rows.forEach((row, index) => {
            const card = recordTemplate.content.firstElementChild.cloneNode(true);
            bindRecord(card, row, index);
            recordList.append(card);
        });
        emptyState.classList.toggle("d-none", state.rows.length > 0);
        bottomAddWrap.classList.toggle("d-none", state.rows.length === 0);
        updateTotal();
        markDuplicates();
    }

    function updateTotal() {
        const total = state.rows.reduce((sum, row) => {
            const number = Number(row.workload);
            return sum + (Number.isFinite(number) ? number : 0);
        }, 0);
        totalWorkload.textContent = formatNumber(total);
    }

    function fieldContainer(element) {
        return element.closest(
            ".field-category,.field-standard,.field-workload,.field-date,.record-body > .mt-3"
        );
    }

    function setInvalid(element, message) {
        if (!element) return;
        element.classList.add("is-invalid");
        const feedback = $(".invalid-feedback", fieldContainer(element));
        if (feedback) feedback.textContent = message;
    }

    function clearInvalid(element) {
        element?.classList.remove("is-invalid");
    }

    function duplicateMap() {
        const firstByStandard = new Map();
        const duplicates = new Map();
        state.rows.forEach((row, index) => {
            if (!row.standard) return;
            if (firstByStandard.has(row.standard)) {
                duplicates.set(index, firstByStandard.get(row.standard));
            } else {
                firstByStandard.set(row.standard, index);
            }
        });
        return duplicates;
    }

    function markDuplicates() {
        const duplicates = duplicateMap();
        state.rows.forEach((row, index) => {
            const card = document.querySelector(`[data-row-id="${row.id}"]`);
            if (!card) return;
            const input = $(".standard-input", card);
            if (duplicates.has(index)) {
                setInvalid(input, `该成果类别与记录 ${duplicates.get(index) + 1} 重复`);
            } else if (row.standard) {
                clearInvalid(input);
            }
        });
    }

    function validate() {
        $$(".is-invalid").forEach(clearInvalid);
        let firstInvalid = null;
        const invalidate = (element, message) => {
            setInvalid(element, message);
            firstInvalid ||= element;
        };

        if (!state.dept) invalidate(deptInput, "请选择所属系部");
        if (!state.name.trim()) invalidate(nameInput, "请输入姓名");
        if (!state.title.trim()) invalidate(titleInput, "请输入职称");
        if (!state.term) invalidate(termInput, "请选择学期");
        if (!state.rows.length) {
            showToast("请至少添加一条教学工作量记录", true);
            return { valid: false, firstInvalid: $("[data-action='add-record']") };
        }

        const duplicates = duplicateMap();
        state.rows.forEach((row, index) => {
            const card = document.querySelector(`[data-row-id="${row.id}"]`);
            if (!row.category) invalidate($(".category-input", card), "请选择项目类别");
            if (!row.standard) invalidate($(".standard-input", card), "请选择计算标准");
            const workload = Number(row.workload);
            if (row.workload === "" || !Number.isFinite(workload) || workload < 0) {
                invalidate($(".workload-input", card), "请输入有效工作量");
            }
            if (!row.time) invalidate($(".date-input", card), "请选择实际取得时间");
            if (!row.remark.trim()) {
                invalidate($(".remark-input", card), "请填写具体成果或业绩内容");
            }
            if (duplicates.has(index)) {
                invalidate(
                    $(".standard-input", card),
                    `该成果类别与记录 ${duplicates.get(index) + 1} 重复`
                );
            }
        });
        return { valid: !firstInvalid, firstInvalid };
    }

    function focusInvalid(element) {
        element?.scrollIntoView({ behavior: "smooth", block: "center" });
        setTimeout(() => element?.focus({ preventScroll: true }), 300);
    }

    function downloadBlob(blob, filename) {
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = filename;
        document.body.append(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    async function exportExcel() {
        if (exportInFlight) {
            showToast("报表正在生成，请稍候");
            return;
        }
        const result = validate();
        if (!result.valid) {
            focusInvalid(result.firstInvalid);
            return;
        }

        exportInFlight = true;
        const button = $("#export-button");
        const originalContent = button.innerHTML;
        button.innerHTML =
            '<span class="spinner-border spinner-border-sm me-2" aria-hidden="true"></span>正在生成报表';
        try {
            const response = await fetch("/api/export", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload()),
            });
            if (!response.ok) {
                const error = await response.json().catch(() => ({}));
                throw new Error(error.message || "生成报表失败");
            }
            const filename =
                `${state.dept}-${state.name.trim()}-${state.title.trim()}-教师工作量填报表.xlsx`;
            downloadBlob(await response.blob(), filename);
            showToast("Excel 报表已生成并开始下载");
        } catch (error) {
            showToast(error.message, true);
        } finally {
            exportInFlight = false;
            button.innerHTML = originalContent;
        }
    }

    function saveProgress() {
        const content = {
            schema_version: 2,
            saved_at: new Date().toISOString(),
            ...payload(),
        };
        const blob = new Blob([JSON.stringify(content, null, 2)], {
            type: "application/json;charset=utf-8",
        });
        downloadBlob(
            blob,
            `${state.dept}-${state.name.trim() || "未命名"}-教学工作量填报进度.json`
        );
        showToast("填报进度已下载");
    }

    async function restoreProgress(file) {
        if (!file) return;
        if (file.size > 2 * 1024 * 1024) {
            showToast("暂存文件不能超过 2MB", true);
            return;
        }
        try {
            const content = JSON.parse(await file.text());
            const response = await fetch("/api/restore", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(content),
            });
            const result = await response.json();
            if (!response.ok) throw new Error(result.message || "恢复失败");
            const restored = result.data;
            state.dept = restored.dept;
            state.name = restored.name;
            state.title = restored.title;
            state.term = restored.term;
            state.rows = restored.rows.map((row) => ({ id: makeId(), ...row }));
            syncBasicForm();
            renderRecords();
            showToast(`已恢复 ${state.rows.length} 条记录`);
            $("#detail-heading").scrollIntoView({ behavior: "smooth", block: "start" });
        } catch (error) {
            showToast(`无法恢复暂存文件：${error.message}`, true);
        } finally {
            restoreFile.value = "";
        }
    }

    populateSelect(deptInput, config.departments);
    populateSelect(termInput, config.terms);
    $("#app-version").textContent = config.version;
    syncBasicForm();

    deptInput.addEventListener("change", () => {
        state.dept = deptInput.value;
        clearInvalid(deptInput);
    });
    nameInput.addEventListener("input", () => {
        state.name = nameInput.value;
        clearInvalid(nameInput);
    });
    titleInput.addEventListener("input", () => {
        state.title = titleInput.value;
        clearInvalid(titleInput);
    });
    termInput.addEventListener("change", () => {
        state.term = termInput.value;
        clearInvalid(termInput);
    });
    $$("[data-action='add-record']").forEach((button) =>
        button.addEventListener("click", addRecord)
    );
    $("#export-button").addEventListener("click", exportExcel);
    $("#save-progress").addEventListener("click", saveProgress);
    $("#restore-progress").addEventListener("click", () => restoreFile.click());
    restoreFile.addEventListener("change", () => restoreProgress(restoreFile.files[0]));
    document.addEventListener("mousedown", (event) => {
        $$(".standard-menu.show").forEach((menu) => {
            if (!menu.parentElement.contains(event.target)) menu.classList.remove("show");
        });
    });

    renderRecords();
})();
