/* TabEscape — Увлекательное фэнтези-приключение, замаскированное под оригинальный Excel Web (Office 365).
 * Создание книг и листов Excel с нуля, загрузка любых своих .xlsx файлов и экспорт со всеми стилями и цветами. */
(() => {
    "use strict";

    // ------------------------------------------------------------------ утилиты
    const colName = (n) => {
        let t = "";
        while (n > 0) {
            const m = (n - 1) % 26;
            t = String.fromCharCode(65 + m) + t;
            n = Math.floor((n - 1) / 26);
        }
        return t;
    };
    const LETTERS = Array.from({ length: 140 }, (_, i) => colName(i + 1));
    let NCOLS = 22;
    const $ = (id) => document.getElementById(id);
    const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
    const key = (r, c) => r + ":" + c;
    const refOf = (r, c) => LETTERS[c - 1] + r;

    const store = {
        get(k, d) { try { const v = localStorage.getItem(k); return v === null ? d : v; } catch { return d; } },
        set(k, v) { try { localStorage.setItem(k, v); } catch { /* ignore */ } },
    };
    const session = {
        get(k, d) { try { const v = sessionStorage.getItem(k); return v === null ? d : v; } catch { return d; } },
        set(k, v) { try { sessionStorage.setItem(k, v); } catch { /* ignore */ } },
    };
    const nf = (n, o) => n.toLocaleString("ru-RU", o);

    function seeded(seed) {
        let s = seed;
        return () => ((s = (s * 1664525 + 1013904223) % 4294967296) / 4294967296);
    }

    function arrayBufferToBase64(buf) {
        const bytes = new Uint8Array(buf);
        let binary = "";
        const chunk = 0x8000;
        for (let i = 0; i < bytes.length; i += chunk) {
            binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
        }
        return btoa(binary);
    }

    function createBlankWorkbookObj(name = "Книга1.xlsx") {
        return {
            name,
            fromScratch: true,
            sheets: [
                {
                    name: "Лист1",
                    cells: [],
                    sparse: {},
                    merges: [],
                    cols: Array.from({ length: 26 }, () => 80),
                    rowH: {},
                    nrows: 60,
                    ncols: 22,
                },
            ],
        };
    }

    // ------------------------------------------------------------------ состояние
    const START_STATS = { ROI: 100, KPI: 100, Anxiety: 10 };
    const DEFAULT_WB_NAME = "Книга1.xlsx";

    const S = {
        stats: { ...START_STATS }, options: [], turn: 1, over: false, won: false,
        narrative: "", auditLog: "", sceneId: null, seen: [], inventory: [], log: [],
        history: [{ ...START_STATS }], loading: null, custom: "",
        role: store.get("te_role", "Начинающий искатель"),
        boss: store.get("te_boss", "Мастер Гильдии Морган"),
        apiKey: session.get("te_key", ""),
        imported: null,
        sourceB64: null,
        impIdx: 0,
        bossMode: false,
        stealthTabs: false,
        sheet: "adv",
        prevSheet: "adv",
        tab: "home",
        ribbonCollapsed: false,
        theme: store.get("te_theme", "light"),
        zoom: 100,
        anchor: { r: 2, c: 2 },
        cur: { r: 2, c: 2 },
        drag: false,
        edits: {},
        styles: {},
        fillColor: "#dce6f2",
        fontColor: "#ff0000",
        best: parseInt(store.get("te_best", "0"), 10) || 0,
    };

    const clockMinutes = () => 9 * 60 + 30 * (S.turn - 1);
    const clockStr = () => {
        const m = Math.min(clockMinutes(), 18 * 60);
        return String(Math.floor(m / 60)).padStart(2, "0") + ":" + String(m % 60).padStart(2, "0");
    };
    const score = () => Math.max(0, S.stats.ROI + S.stats.KPI + (100 - S.stats.Anxiety) + S.inventory.length * 10);

    const ITEM_DESC = {
        "Письмо в Гильдию": "Рекомендательное письмо с печатью: вызывает доверие и снимает подозрения.",
        "Карта долины": "Подробная карта старого следопыта, помогающая легко ориентироваться в пути.",
        "Тёплый пирожок": "Ароматная выпечка от хозяйки Марты (+25 Здоровья и Бодрости).",
        "Значок Гильдии": "Бронзовый значок искателя приключений Оукхейвена.",
        "Справочник трав": "Карманная книжка травницы с рецептами целебных настоев.",
        "Старинный компас": "Латунный компас часового мастера, всегда указывающий верный путь.",
        "Кошачий амулет": "Блестящий медальон от говорящего кота из башни мага.",
        "Зелье бодрости": "Яблочно-мятный тоник чародея Корнелиуса, возвращающий силы.",
        "Серебряный ключ": "Подарок мельника за помощь с озорными енотами.",
        "Линза рассвета": "Солнечный кристалл из Эльфийской Обсерватории.",
        "Чешуйка удачи": "Золотистая чешуйка доброй драконицы Аурелии, приносящая удачу.",
    };

    // ------------------------------------------------------------------ формулы
    const FN = {
        SUM: (...a) => a.flat(Infinity).reduce((s, v) => s + (typeof v === "number" ? v : 0), 0),
        AVERAGE: (...a) => {
            const n = a.flat(Infinity).filter((v) => typeof v === "number");
            if (!n.length) throw new Error("#ДЕЛ/0!");
            return n.reduce((s, v) => s + v, 0) / n.length;
        },
        MIN: (...a) => Math.min(...a.flat(Infinity).filter((v) => typeof v === "number")),
        MAX: (...a) => Math.max(...a.flat(Infinity).filter((v) => typeof v === "number")),
        COUNT: (...a) => a.flat(Infinity).filter((v) => typeof v === "number").length,
        ROUND: (x, n = 0) => { const k = 10 ** n; return Math.round(x * k) / k; },
        IF: (c, a, b = false) => (c ? a : b),
        ABS: (x) => Math.abs(x),
    };
    const FN_ALIAS = {
        "СУММ": "SUM", "СРЗНАЧ": "AVERAGE", "МИН": "MIN", "МАКС": "MAX", "СЧЁТ": "COUNT", "СЧЕТ": "COUNT",
        "ОКРУГЛ": "ROUND", "ЕСЛИ": "IF", "ABS": "ABS", "SUM": "SUM", "AVERAGE": "AVERAGE", "MIN": "MIN",
        "MAX": "MAX", "COUNT": "COUNT", "ROUND": "ROUND", "IF": "IF",
    };

    function colIndex(letters) {
        return letters.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    }

    function resolveSheetIdByName(name) {
        const clean = String(name).replace(/^'|'$/g, "").trim();
        if (S.imported && S.imported.sheets) {
            const idx = S.imported.sheets.findIndex((sh) => sh.name === clean);
            if (idx >= 0) return "imp" + idx;
        }
        return null;
    }

    function evalFormula(sheetId, formula, depth) {
        if (depth > 25) return "#ЦИКЛ!";
        try {
            const orig = formula.slice(1);
            const strs = [];
            let s = orig.replace(/"([^"]*)"/g, (m, a) => {
                strs.push(a);
                return `\u0001${strs.length - 1}\u0001`;
            });
            s = s.replace(/(?:'([^']+)'|([A-Za-zА-Яа-яЁё0-9_]+))!\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?/g, (m, qName, bName, colL, rowN) => {
                const targetSheet = resolveSheetIdByName(qName || bName) || sheetId;
                return `VS("${targetSheet}","${colL.toUpperCase()}${rowN}")`;
            });
            s = s.toUpperCase();
            s = s.replace(/([A-ZА-ЯЁ_][A-ZА-ЯЁ_0-9.]*)\s*\(/g, (m, name) => {
                if (name === "VS") return "VS(";
                const fn = FN_ALIAS[name];
                if (!fn) throw new Error("#ИМЯ?");
                return `F["${fn}"](`;
            });
            s = s.replace(/\$?([A-Z]{1,3})\$?(\d{1,5}):\$?([A-Z]{1,3})\$?(\d{1,5})/g, (m, a, b, c, d) => `R("${a}${b}","${c}${d}")`);
            s = s.replace(/(?<![A-Z"\w])\$?([A-Z]{1,3})\$?(\d{1,5})(?![\w(])/g, (m, a, b) => `V("${a}${b}")`);
            s = s.replace(/;/g, ",").replace(/<>/g, "!=").replace(/(?<![<>!=])=(?!=)/g, "==").replace(/\^/g, "**")
                .replace(/(\d+(?:\.\d+)?)%/g, "($1/100)").replace(/&/g, "+");
            const rest = s.replace(/F\["[A-Z]+"\]/g, "")
                .replace(/VS\("[^"]+","[A-Z0-9]+"\)/g, "")
                .replace(/[RV]\("[A-Z0-9]+"(?:,"[A-Z0-9]+")?\)/g, "")
                .replace(/\u0001\d+\u0001/g, "");
            if (/[A-Za-zА-Яа-яЁё_]/.test(rest)) throw new Error("#ИМЯ?");
            s = s.replace(/\u0001(\d+)\u0001/g, (m, i) => JSON.stringify(strs[+i]));

            const V = (ref) => {
                const v = cellValue(sheetId, ref, depth + 1);
                if (typeof v === "string" && v.startsWith("#")) throw new Error(v);
                return v === null ? 0 : v;
            };
            const VS = (shId, ref) => {
                const v = cellValue(shId.toLowerCase(), ref, depth + 1);
                if (typeof v === "string" && v.startsWith("#")) throw new Error(v);
                return v === null ? 0 : v;
            };
            const R = (a, b) => {
                const ma = /([A-Z]+)(\d+)/.exec(a), mb = /([A-Z]+)(\d+)/.exec(b), out = [];
                const c1 = colIndex(ma[1]), c2 = colIndex(mb[1]), r1 = +ma[2], r2 = +mb[2];
                for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
                    for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
                        const v = cellValue(sheetId, refOf(r, c), depth + 1);
                        if (typeof v === "string" && v.startsWith("#")) throw new Error(v);
                        out.push(v);
                    }
                }
                return out;
            };
            // eslint-disable-next-line no-new-func
            const res = new Function("F", "R", "V", "VS", `"use strict"; return (${s});`)(FN, R, V, VS);
            if (typeof res === "number" && !Number.isFinite(res)) return "#ДЕЛ/0!";
            return res;
        } catch (e) {
            return /^#/.test(e.message) ? e.message : "#ЗНАЧ!";
        }
    }

    function rawAt(sheet, r, c) {
        const ed = S.edits[sheet.id] && S.edits[sheet.id][key(r, c)];
        if (ed !== undefined) return ed;
        const d = sheet.cells[key(r, c)];
        return d ? d.raw : undefined;
    }

    function cellValue(sheetId, ref, depth = 0) {
        const m = /^([A-Z]+)(\d+)$/.exec(ref);
        if (!m) return null;
        const sheet = getSheet(sheetId);
        const r = +m[2], c = colIndex(m[1]);
        const ed = S.edits[sheet.id] && S.edits[sheet.id][key(r, c)];
        const d = sheet.cells[key(r, c)];
        if (ed === undefined && d && d.val !== undefined && d.val !== null && !(typeof d.raw === "string" && d.raw.startsWith("="))) {
            return d.val;
        }
        const raw = rawAt(sheet, r, c);
        if (raw === undefined || raw === "") return null;
        if (!(d && d.text) && typeof raw === "string" && raw[0] === "=") {
            const ev = evalFormula(sheetId, raw, depth);
            if (typeof ev === "string" && ev.startsWith("#") && d && d.val !== undefined && d.val !== null) {
                return d.val;
            }
            return ev;
        }
        return parseInput(raw);
    }

    function parseInput(raw) {
        if (typeof raw === "number") return raw;
        const t = String(raw).trim().replace(/[\s\u00a0]/g, "");
        if (/^-?\d+([.,]\d+)?$/.test(t)) return parseFloat(t.replace(",", "."));
        if (/^-?\d+([.,]\d+)?%$/.test(t)) return parseFloat(t.replace(",", ".")) / 100;
        return String(raw);
    }

    function formatValue(v, fmt) {
        if (v === null || v === undefined) return "";
        if (typeof v === "boolean") return v ? "ИСТИНА" : "ЛОЖЬ";
        if (typeof v !== "number") return String(v);
        switch (fmt) {
            case "money": return nf(Math.round(v), { maximumFractionDigits: 0 }) + " ₽";
            case "int": return nf(Math.round(v), { maximumFractionDigits: 0 });
            case "pct": return nf(v * 100, { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + "%";
            case "dec": return nf(v, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            default: return nf(v, { maximumFractionDigits: 6 });
        }
    }

    // ------------------------------------------------------------------ листы
    const SHEETS = {};
    function getSheet(id) {
        if (id && id.startsWith("imp") && S.imported) return importedSheet(+id.slice(3));
        return (SHEETS[id] || SHEETS.adv)();
    }
    const isBossSheet = (id) => id === "boss" || (id && id.startsWith("imp"));

    const DECOY_ITEMS = [
        ["GL-401", "Амортизация и износ основных средств (график, РСБУ)"],
        ["GL-402", "Облачная инфраструктура и лицензии SaaS"],
        ["GL-403", "Региональные коммерческие и управленческие расходы"],
        ["GL-404", "Отложенные налоговые обязательства"],
        ["GL-405", "Резерв по оплате труда и неиспользованным отпускам"],
        ["GL-406", "Обслуживание кредитной линии оборотного капитала"],
        ["GL-407", "Резерв под реструктуризацию и условные обязательства"],
        ["GL-408", "Курсовые разницы и операции хеджирования"],
        ["GL-409", "Внутригрупповые взаимозачёты и элиминация"],
        ["GL-410", "Расходы на аудит и консалтинг"],
        ["GL-411", "Маркетинг и представительские расходы"],
        ["GL-412", "Командировочные и транспортные расходы"],
        ["GL-413", "Прочие операционные расходы"],
    ];

    SHEETS.adv = () => {
        const cells = {}, rowCls = {};
        const put = (r, c, o) => { cells[key(r, c)] = o; };
        const rnd = seeded(26);
        put(1, 1, { raw: "ID", cls: "hdr" });
        put(1, 2, { raw: "ДНЕВНИК ПРИКЛЮЧЕНИЙ ОУКХЕЙВЕНА (МАСКИРОВКА: ОТЧЁТ Q4)", cls: "hdr" });
        put(1, 3, { raw: "Бюджет", cls: "hdr c" });
        put(1, 4, { raw: "Параметр героя", cls: "hdr" });
        put(1, 5, { raw: "Значение", cls: "hdr c" });
        put(1, 6, { raw: "Состояние", cls: "hdr" });
        put(1, 7, { raw: "ФГ 2026", cls: "hdr c" });
        put(1, 8, { raw: "", cls: "hdr" });
        rowCls[1] = "tbl-head";

        const money = [1240500, 410200, 88900, 323000, 512400];
        put(2, 1, { raw: "СЦЕНА", cls: "code", locked: true });
        put(2, 2, { cls: "story", locked: true, html: () => S.loading ? esc(S.loading) : esc(S.narrative), fx: () => (S.loading ? "" : S.auditLog) });
        put(3, 1, { raw: "СВОЙ ХОД", cls: "code", locked: true });
        put(3, 2, { cls: "prompt", html: () => S.custom ? esc(S.custom) : "✎ Свой ход: напишите действие в строке формул (B3) и нажмите Enter", editableAction: true });
        for (let i = 0; i < 3; i++) {
            put(4 + i, 1, { raw: `[ВЫБОР ${i + 1}]`, cls: "code", locked: true });
            put(4 + i, 2, {
                cls: "optcell", locked: true,
                html: () => {
                    if (S.loading) return "";
                    if (S.over) return ["▶ Начать новое приключение", S.won ? `🏆 Победа! Очки: ${score()} · Рекорд: ${S.best}` : `💫 Приключение окончено! Очки: ${score()} · Рекорд: ${S.best}`, "⤓ Скачать историю похода (.txt)"][i];
                    return `<span class="optkey">${i + 1}</span>${esc(S.options[i] || "…")}`;
                },
                fx: () => (S.over ? "" : S.options[i] || ""),
            });
            rowCls[4 + i] = "opt" + (S.loading ? " disabled" : "");
        }
        rowCls[2] = "story";
        for (let r = 2; r <= 6; r++) put(r, 3, { raw: money[r - 2], fmt: "money", locked: true });

        const statCell = (r, name, k, tiers, fxName) => {
            put(r, 4, { raw: name, locked: true });
            put(r, 5, {
                locked: true, fx: "=" + fxName,
                html: () => { const v = S.stats[k]; const lvl = tiers(v); return `<div class="databar bar-${lvl}"><i style="width:${v}%"></i><span>${v}%</span></div>`; },
            });
            put(r, 6, {
                locked: true, fx: "=ЕСЛИ(" + fxName + ">50;\"НОРМА\";\"ВНИМАНИЕ\")",
                html: () => {
                    const v = S.stats[k];
                    const lvl = tiers(v);
                    const txt = {
                        ROI: ["ПОЛОН СИЛ", "УШИБЫ", "НУЖЕН ЛЕКАРЬ"],
                        KPI: ["БОДР", "УТОМЛЁН", "ВАЛИТСЯ С НОГ"],
                        Anxiety: ["СПОКОЙНО", "ЗАМЕЧЕН", "ИЩЕТ СТРАЖА"],
                    }[k][{ ok: 0, warn: 1, bad: 2 }[lvl]];
                    return `<span class="badge ${lvl}">${txt}</span>`;
                },
            });
        };
        statCell(2, "Здоровье / ОЗ", "ROI", (v) => (v >= 60 ? "ok" : v >= 35 ? "warn" : "bad"), "ЗДОРОВЬЕ");
        statCell(3, "Бодрость / Мана", "KPI", (v) => (v >= 55 ? "ok" : v >= 30 ? "warn" : "bad"), "БОДРОСТЬ");
        statCell(4, "Неприятности", "Anxiety", (v) => (v <= 40 ? "ok" : v <= 70 ? "warn" : "bad"), "РОЗЫСК");
        put(5, 4, { raw: "Время дня", locked: true });
        put(5, 5, { locked: true, cls: "r", fx: "=ТЕКСТ(ВРЕМЯ();\"чч:мм\")", html: () => `<b>${clockStr()}</b>` });
        put(5, 6, { locked: true, html: () => { const left = Math.max(0, 18 * 60 - clockMinutes()); return `<span class="badge ${left > 180 ? "ok" : "warn"}">ДО ВЕЧЕРА — ${Math.floor(left / 60)} Ч ${left % 60} МИН</span>`; }, fx: "=ВРЕМЯ(18;0;0)-ВРЕМЯ()" });
        put(6, 4, { raw: "Находки", locked: true });
        put(6, 5, { locked: true, cls: "r", fx: "=СЧЁТ(Инвентарь!B:B)", html: () => `<b>${S.inventory.length}</b>` });
        put(6, 6, { locked: true, html: () => `<span class="badge ${S.inventory.length ? "ok" : "warn"}">${S.inventory.length ? "В СУМКЕ" : "ПУСТО"}</span>` });

        const hdr = ["Код", "Статья баланса", "1 кв.", "2 кв.", "3 кв.", "4 кв.", "Итого", "Δ 4/3 кв."];
        hdr.forEach((h, i) => put(8, i + 1, { raw: h, cls: "hdr" + (i > 1 ? " c" : "") }));
        rowCls[8] = "tbl-head";
        DECOY_ITEMS.forEach(([code, name], i) => {
            const r = 9 + i;
            put(r, 1, { raw: code, cls: "code" });
            put(r, 2, { raw: name });
            const base = 200000 + Math.floor(rnd() * 2400000);
            for (let q = 0; q < 4; q++) put(r, 3 + q, { raw: Math.round(base * (0.85 + rnd() * 0.35) / 100) * 100, fmt: "money" });
            put(r, 7, { raw: `=СУММ(C${r}:F${r})`, fmt: "money" });
            put(r, 8, { raw: `=F${r}/E${r}-1`, fmt: "pct", signed: true });
            if (i % 2) rowCls[r] = "zebra";
        });
        const tr = 9 + DECOY_ITEMS.length, last = tr - 1;
        put(tr, 2, { raw: "ИТОГО по статьям", cls: "bold" });
        for (let c = 3; c <= 7; c++) put(tr, c, { raw: `=СУММ(${LETTERS[c - 1]}9:${LETTERS[c - 1]}${last})`, fmt: "money" });
        put(tr, 8, { raw: `=F${tr}/E${tr}-1`, fmt: "pct", signed: true });
        rowCls[tr] = "total";
        put(tr + 2, 2, { raw: "Средний итог по статье", cls: "ital" });
        put(tr + 2, 7, { raw: `=СРЗНАЧ(G9:G${last})`, fmt: "money" });
        put(tr + 3, 2, { raw: "Максимальный итог", cls: "ital" });
        put(tr + 3, 7, { raw: `=МАКС(G9:G${last})`, fmt: "money" });
        return { id: "adv", name: "Приключение", cells, rowCls, cols: [88, 570, 120, 170, 140, 190, 130, 110, 90, 90, 90, 90], ncols: 12, rows: 60, lockedFrom: 7 };
    };

    SHEETS.boss = () => {
        const cells = {}, rowCls = {}, rnd = seeded(7);
        const put = (r, c, o) => { cells[key(r, c)] = o; };
        ["№", "Консолидированный прогноз Q4 2026 (признание выручки по МСФО 15)", "Факт Q3", "План Q4", "Прогноз Q4", "Откл., ₽", "Откл., %", "Статус"]
            .forEach((h, i) => put(1, i + 1, { raw: h, cls: "hdr" + (i > 1 ? " c" : "") }));
        rowCls[1] = "tbl-head";
        put(2, 2, { raw: "Все суммы указаны в рублях без НДС. Отклонения в пределах допустимого порога ±1,4%.", cls: "ital muted" });
        const lines = [
            "Выручка — Москва", "Выручка — Санкт-Петербург", "Выручка — Екатеринбург", "Выручка — Новосибирск", "Выручка — Казань",
            "Себестоимость реализации", "Коммерческие расходы", "Управленческие расходы", "Амортизация НМА",
            "Прочие операционные доходы", "Процентные расходы", "Налог на прибыль (текущий)", "Отложенный налог",
        ];
        lines.forEach((name, i) => {
            const r = 3 + i;
            put(r, 1, { raw: String(401 + i), cls: "code" });
            put(r, 2, { raw: name });
            const plan = Math.round((300000 + rnd() * 4200000) / 100) * 100;
            put(r, 3, { raw: Math.round(plan * (0.9 + rnd() * 0.15) / 100) * 100, fmt: "money" });
            put(r, 4, { raw: plan, fmt: "money" });
            put(r, 5, { raw: Math.round(plan * (0.97 + rnd() * 0.06) / 100) * 100, fmt: "money" });
            put(r, 6, { raw: `=E${r}-D${r}`, fmt: "money" });
            put(r, 7, { raw: `=E${r}/D${r}-1`, fmt: "pct", signed: true });
            put(r, 8, { raw: ["Сверено", "Утверждено", "На проверке", "Сверено"][i % 4], cls: "muted" });
            if (i % 2) rowCls[r] = "zebra";
        });
        const tr = 3 + lines.length, last = tr - 1;
        put(tr, 2, { raw: "ИТОГО", cls: "bold" });
        ["C", "D", "E", "F"].forEach((L, i) => put(tr, 3 + i, { raw: `=СУММ(${L}3:${L}${last})`, fmt: "money" }));
        put(tr, 7, { raw: `=E${tr}/D${tr}-1`, fmt: "pct", signed: true });
        rowCls[tr] = "total";
        return { id: "boss", name: "Q4_Прогноз", cells, rowCls, cols: [60, 520, 130, 130, 130, 130, 100, 130, 90, 90, 90, 90], ncols: 12, rows: 60, lockedFrom: 99 };
    };

    SHEETS.log = () => {
        const cells = {}, rowCls = {};
        const put = (r, c, o) => { cells[key(r, c)] = Object.assign({ locked: true }, o); };
        ["№", "Время", "Событие приключения", "Здоровье", "Бодрость", "Неприятности", "Запись в дневнике"].forEach((h, i) => put(1, i + 1, { raw: h, cls: "hdr" + (i === 3 || i === 4 || i === 5 ? " c" : "") }));
        rowCls[1] = "tbl-head";
        S.log.forEach((e, i) => {
            const r = 2 + i;
            put(r, 1, { raw: e.turn, cls: "code" });
            put(r, 2, { raw: e.time });
            put(r, 3, { raw: e.action });
            put(r, 4, { raw: e.roi / 100, fmt: "pct" });
            put(r, 5, { raw: e.kpi / 100, fmt: "pct" });
            put(r, 6, { raw: e.anx / 100, fmt: "pct" });
            put(r, 7, { raw: e.formula, cls: "code", text: true });
            if (i % 2) rowCls[r] = "zebra";
        });
        return { id: "log", name: "Журнал", cells, rowCls, cols: [50, 70, 440, 90, 90, 90, 340, 90, 90, 90, 90, 90], ncols: 12, rows: Math.max(40, S.log.length + 12), lockedFrom: 999, chart: true };
    };

    SHEETS.inv = () => {
        const cells = {}, rowCls = {};
        const put = (r, c, o) => { cells[key(r, c)] = Object.assign({ locked: true }, o); };
        ["№", "Предмет", "Описание", "Кол-во", "Применение"].forEach((h, i) => put(1, i + 1, { raw: h, cls: "hdr" + (i > 2 ? " c" : "") }));
        rowCls[1] = "tbl-head";
        const grouped = {};
        S.inventory.forEach((n) => { grouped[n] = (grouped[n] || 0) + 1; });
        const names = Object.keys(grouped);
        if (!names.length) put(2, 2, { raw: "Сумка пока пуста — гуляйте по городу и общайтесь с жителями.", cls: "muted" });
        names.forEach((n, i) => {
            const r = 2 + i;
            put(r, 1, { raw: i + 1, cls: "code" });
            put(r, 2, { raw: n, cls: "bold" });
            put(r, 3, { raw: ITEM_DESC[n] || "Полезная находка из Оукхейвена (+25 Здоровья и Бодрости)." });
            put(r, 4, { raw: grouped[n], cls: "c" });
            put(r, 5, { cls: "c", html: () => `<button class="use-btn" data-use="${esc(n)}">Использовать</button>` });
            if (i % 2) rowCls[r] = "zebra";
        });
        return { id: "inv", name: "Инвентарь", cells, rowCls, cols: [50, 260, 420, 90, 150, 90, 90, 90, 90, 90, 90, 90], ncols: 12, rows: 40, lockedFrom: 999 };
    };

    function importedSheet(i) {
        const src = S.imported.sheets[i] || S.imported.sheets[0];
        const cells = {}, merges = {}, skip = {};
        src.cells.forEach(([r, c, rawVal, dispVal, stObj, cachedVal]) => {
            cells[key(r, c)] = {
                raw: rawVal,
                disp: dispVal,
                val: cachedVal !== undefined ? cachedVal : rawVal,
                st: stObj || null,
                text: typeof rawVal === "string" && rawVal[0] !== "=",
            };
        });
        src.merges.forEach(([r1, c1, r2, c2]) => {
            merges[key(r1, c1)] = { rs: r2 - r1 + 1, cs: c2 - c1 + 1, r2, c2 };
            const anchorKey = key(r1, c1);
            const anchorCell = cells[anchorKey] || (cells[anchorKey] = { raw: "", disp: null, val: null, st: {} });
            const aSt = Object.assign({}, anchorCell.st || {});
            const trCell = cells[key(r1, c2)];
            const blCell = cells[key(r2, c1)];
            const brCell = cells[key(r2, c2)];
            if (trCell && trCell.st && trCell.st.br) aSt.br = trCell.st.br;
            else if (brCell && brCell.st && brCell.st.br) aSt.br = brCell.st.br;
            if (blCell && blCell.st && blCell.st.bb) aSt.bb = blCell.st.bb;
            else if (brCell && brCell.st && brCell.st.bb) aSt.bb = brCell.st.bb;
            anchorCell.st = aSt;

            for (let r = r1; r <= r2; r++) {
                for (let c = c1; c <= c2; c++) {
                    if (r !== r1 || c !== c1) skip[key(r, c)] = anchorKey;
                }
            }
        });
        const nc = Math.max(24, Math.min(130, src.ncols + 4));
        const cols = Array.from({ length: nc }, (_, k) => src.cols[k] || 76);
        return {
            id: "imp" + i,
            name: src.name,
            cells,
            sparse: src.sparse || {},
            rowCls: {},
            cols,
            ncols: nc,
            rows: Math.max(60, src.nrows + 18),
            lockedFrom: 1,
            merges,
            skip,
            rowH: src.rowH || {},
        };
    }

    // ------------------------------------------------------------------ Парсер .xlsx со всеми стилями, цветами темы и рамками
    const INDEXED_PALETTE = {
        0: "#000000", 1: "#FFFFFF", 2: "#FF0000", 3: "#00FF00", 4: "#0000FF", 5: "#FFFF00", 6: "#FF00FF", 7: "#00FFFF",
        8: "#000000", 9: "#FFFFFF", 10: "#FF0000", 11: "#00FF00", 12: "#0000FF", 13: "#FFFF00", 14: "#FF00FF", 15: "#00FFFF",
        16: "#800000", 17: "#008000", 18: "#000080", 19: "#808000", 20: "#800080", 21: "#008080", 22: "#C0C0C0", 23: "#808080",
        24: "#9999FF", 25: "#993366", 26: "#FFFFCC", 27: "#CCFFFF", 28: "#660066", 29: "#FF8080", 30: "#0066CC", 31: "#CCCCFF",
        32: "#000080", 33: "#FF00FF", 34: "#FFFF00", 35: "#00FFFF", 36: "#800080", 37: "#800000", 38: "#008080", 39: "#0000FF",
        40: "#00CCFF", 41: "#CCFFFF", 42: "#CCFFCC", 43: "#FFFF99", 44: "#99CCFF", 45: "#FF99CC", 46: "#CC99FF", 47: "#FFCC99",
        48: "#3366FF", 49: "#33CCCC", 50: "#99CC00", 51: "#FFCC00", 52: "#FF9900", 53: "#FF6600", 54: "#666699", 55: "#969696",
        56: "#003366", 57: "#339966", 58: "#003300", 59: "#333300", 60: "#993300", 61: "#993366", 62: "#333399", 63: "#333333",
        64: "#000000", 81: "#000000",
    };

    function rgbToHls(r, g, b) {
        r /= 255; g /= 255; b /= 255;
        const max = Math.max(r, g, b), min = Math.min(r, g, b);
        let h = 0, s = 0;
        const l = (max + min) / 2;
        if (max !== min) {
            const d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            switch (max) {
                case r: h = (g - b) / d + (g < b ? 6 : 0); break;
                case g: h = (b - r) / d + 2; break;
                case b: h = (r - g) / d + 4; break;
            }
            h /= 6;
        }
        return [h, l, s];
    }

    function hlsToRgb(h, l, s) {
        let r, g, b;
        if (s === 0) {
            r = g = b = l;
        } else {
            const hue2rgb = (p, q, t) => {
                if (t < 0) t += 1;
                if (t > 1) t -= 1;
                if (t < 1 / 6) return p + (q - p) * 6 * t;
                if (t < 1 / 2) return q;
                if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
                return p;
            };
            const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            const p = 2 * l - q;
            r = hue2rgb(p, q, h + 1 / 3);
            g = hue2rgb(p, q, h);
            b = hue2rgb(p, q, h - 1 / 3);
        }
        const toHex = (x) => Math.max(0, Math.min(255, Math.round(x * 255))).toString(16).padStart(2, "0").toUpperCase();
        return "#" + toHex(r) + toHex(g) + toHex(b);
    }

    function applyExcelTint(hexColor, tint) {
        const clean = hexColor.replace(/^#/, "").slice(-6);
        if (!tint) return "#" + clean.toUpperCase();
        const r = parseInt(clean.slice(0, 2), 16);
        const g = parseInt(clean.slice(2, 4), 16);
        const b = parseInt(clean.slice(4, 6), 16);
        let [h, l, s] = rgbToHls(r, g, b);
        if (tint < 0) l = l * (1 + tint);
        else l = l * (1 - tint) + tint;
        l = Math.max(0, Math.min(1, l));
        return hlsToRgb(h, l, s);
    }

    function resolveExcelColor(colEl, themeList, isFill = false) {
        if (!colEl) return null;
        const rgb = colEl.getAttribute("rgb");
        if (rgb) {
            const hex = "#" + rgb.slice(-6).toUpperCase();
            const tint = parseFloat(colEl.getAttribute("tint") || "0");
            return tint ? applyExcelTint(hex, tint) : hex;
        }
        const themeAttr = colEl.getAttribute("theme");
        if (themeAttr !== null && themeAttr !== "") {
            const idx = parseInt(themeAttr, 10);
            const baseHex = themeList[idx] || (idx === 0 ? "FFFFFF" : "000000");
            const tint = parseFloat(colEl.getAttribute("tint") || "0");
            const resolved = applyExcelTint(baseHex, tint);
            if (isFill && idx === 0 && !tint && resolved === "#FFFFFF") return null;
            return resolved;
        }
        const indexedAttr = colEl.getAttribute("indexed");
        if (indexedAttr !== null && indexedAttr !== "") {
            const idx = parseInt(indexedAttr, 10);
            if (isFill && idx === 64) return null;
            return INDEXED_PALETTE[idx] || null;
        }
        return null;
    }

    function borderCss(bSideEl, themeList) {
        if (!bSideEl) return null;
        const st = bSideEl.getAttribute("style");
        if (!st || st === "none") return null;
        const colEl = bSideEl.getElementsByTagName("color")[0];
        const col = resolveExcelColor(colEl, themeList, false) || "#000000";
        switch (st) {
            case "medium":
            case "mediumDashed":
                return `1.5px ${st.includes("Dashed") ? "dashed" : "solid"} ${col}`;
            case "thick":
                return `2.5px solid ${col}`;
            case "double":
                return `3px double ${col}`;
            case "dashed":
            case "dotted":
            case "hair":
                return `1px ${st === "dotted" ? "dotted" : "dashed"} ${col}`;
            default:
                return `1px solid ${col}`;
        }
    }

    async function unzip(buf) {
        const u8 = new Uint8Array(buf), dv = new DataView(buf), dec = new TextDecoder();
        let e = u8.length - 22;
        while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
        if (e < 0) throw new Error("not zip");
        const n = dv.getUint16(e + 10, true), files = {};
        let p = dv.getUint32(e + 16, true);
        for (let i = 0; i < n; i++) {
            const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
            files[dec.decode(u8.subarray(p + 46, p + 46 + nl))] = {
                m: dv.getUint16(p + 10, true),
                cs: dv.getUint32(p + 20, true),
                lh: dv.getUint32(p + 42, true),
            };
            p += 46 + nl + el + cl;
        }
        return async (name) => {
            const f = files[name];
            if (!f) return null;
            const st = f.lh + 30 + dv.getUint16(f.lh + 26, true) + dv.getUint16(f.lh + 28, true);
            const data = u8.subarray(st, st + f.cs);
            if (f.m === 0) return dec.decode(data);
            return new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"))).text();
        };
    }

    function fmtExcel(v, code, id) {
        if (!code && id === 0) return null;
        const builtin = {
            1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%",
            14: "dd.mm.yyyy", 15: "dd mmm yy", 16: "dd mmm", 17: "mmm yy",
            22: "dd.mm.yyyy hh:mm", 45: "mm:ss", 46: "hh:mm:ss", 47: "mm:ss",
        };
        code = code || builtin[id];
        if (!code || /^general$/i.test(code) || code === "@") return null;
        const first = code.split(";")[0].replace(/\[[^\]]*\]/g, "");
        const bare = first.replace(/"[^"]*"/g, "");
        if (/[dmyhs]/i.test(bare) && !/[0#]/.test(bare)) {
            const dt = new Date(Math.round((v - 25569) * 864e5)), pad = (x) => String(x).padStart(2, "0");
            const M = ["янв", "фев", "мар", "апр", "мая", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"];
            const toks = bare.match(/yyyy|yy|mmmm|mmm|mm|m|dd|d|hh|h|ss|s|[^ymdhs]+/gi) || [];
            return toks.map((t, i) => {
                const l = t.toLowerCase();
                if (l === "yyyy") return dt.getUTCFullYear();
                if (l === "yy") return pad(dt.getUTCFullYear() % 100);
                if (l === "dd") return pad(dt.getUTCDate());
                if (l === "d") return dt.getUTCDate();
                if (l === "hh") return pad(dt.getUTCHours());
                if (l === "h") return dt.getUTCHours();
                if (l === "ss") return pad(dt.getUTCSeconds());
                if (l === "s") return dt.getUTCSeconds();
                if (l === "mmm" || l === "mmmm") return M[dt.getUTCMonth()];
                const prevH = toks.slice(0, i).reverse().find((x) => /[hdms]/i.test(x) && !/^[^a-z]+$/i.test(x));
                const nextS = toks.slice(i + 1).find((x) => /[hdms]/i.test(x) && !/^[^a-z]+$/i.test(x));
                const isMin = (prevH && /^h/i.test(prevH)) || (nextS && /^s/i.test(nextS));
                const val = isMin ? dt.getUTCMinutes() : dt.getUTCMonth() + 1;
                return l === "mm" ? pad(val) : val;
            }).join("");
        }
        const m = /\.([0#?]+)/.exec(bare), dec = m ? m[1].length : 0, pct = bare.includes("%");
        const lit = (first.match(/"([^"]*)"/g) || []).map((x) => x.slice(1, -1)).join("");
        const num = nf(pct ? v * 100 : v, { minimumFractionDigits: dec, maximumFractionDigits: dec, useGrouping: bare.includes(",") });
        const prefixLit = /^"[^"]*"/.test(first) ? lit : "";
        return (prefixLit ? prefixLit + " " : "") + num + (pct ? "%" : "") + (!prefixLit && lit ? " " + lit : "");
    }

    async function parseXlsx(buf) {
        const get = await unzip(buf), dp = new DOMParser();
        const xml = async (n) => { const t = await get(n); return t ? dp.parseFromString(t, "application/xml") : null; };
        const wbX = await xml("xl/workbook.xml"), relX = await xml("xl/_rels/workbook.xml.rels");
        if (!wbX) throw new Error("no workbook");
        const rels = {};
        Array.from(relX.getElementsByTagName("Relationship")).forEach((r) => {
            rels[r.getAttribute("Id")] = r.getAttribute("Target");
        });

        const themeList = ["FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646", "0000FF", "800080"];
        const thX = await xml("xl/theme/theme1.xml");
        if (thX) {
            const tags = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
            tags.forEach((tag, idx) => {
                const el = thX.getElementsByTagName("a:" + tag)[0] || thX.getElementsByTagName(tag)[0];
                if (el && el.firstElementChild) {
                    const ch = el.firstElementChild;
                    const hex = ch.getAttribute("lastClr") || ch.getAttribute("val");
                    if (hex) themeList[idx] = hex.toUpperCase();
                }
            });
        }

        const sst = [];
        const ssX = await xml("xl/sharedStrings.xml");
        if (ssX) {
            Array.from(ssX.getElementsByTagName("si")).forEach((si) => {
                sst.push(Array.from(si.getElementsByTagName("t")).filter((t) => t.parentNode.localName !== "rPh").map((t) => t.textContent).join(""));
            });
        }

        const codes = {}, fonts = [], fills = [], borders = [], xfs = [];
        const stX = await xml("xl/styles.xml");
        if (stX) {
            Array.from(stX.getElementsByTagName("numFmt")).forEach((n) => {
                codes[n.getAttribute("numFmtId")] = n.getAttribute("formatCode");
            });

            const fontsEl = stX.getElementsByTagName("fonts")[0];
            if (fontsEl) {
                Array.from(fontsEl.getElementsByTagName("font")).forEach((f) => {
                    const b = f.getElementsByTagName("b")[0];
                    const i = f.getElementsByTagName("i")[0];
                    const u = f.getElementsByTagName("u")[0];
                    const sz = f.getElementsByTagName("sz")[0];
                    const nm = f.getElementsByTagName("name")[0];
                    const col = f.getElementsByTagName("color")[0];
                    fonts.push({
                        b: Boolean(b && b.getAttribute("val") !== "0"),
                        i: Boolean(i && i.getAttribute("val") !== "0"),
                        u: Boolean(u && u.getAttribute("val") !== "none"),
                        sz: sz ? parseFloat(sz.getAttribute("val") || "11") : 11,
                        name: nm ? nm.getAttribute("val") : "Calibri",
                        color: resolveExcelColor(col, themeList, false),
                    });
                });
            }

            const fillsEl = stX.getElementsByTagName("fills")[0];
            if (fillsEl) {
                Array.from(fillsEl.getElementsByTagName("fill")).forEach((fl) => {
                    const pf = fl.getElementsByTagName("patternFill")[0];
                    if (pf && pf.getAttribute("patternType") === "solid") {
                        const fg = pf.getElementsByTagName("fgColor")[0];
                        fills.push(resolveExcelColor(fg, themeList, true));
                    } else {
                        fills.push(null);
                    }
                });
            }

            const bordersEl = stX.getElementsByTagName("borders")[0];
            if (bordersEl) {
                Array.from(bordersEl.getElementsByTagName("border")).forEach((bd) => {
                    borders.push({
                        bl: borderCss(bd.getElementsByTagName("left")[0], themeList),
                        br: borderCss(bd.getElementsByTagName("right")[0], themeList),
                        bt: borderCss(bd.getElementsByTagName("top")[0], themeList),
                        bb: borderCss(bd.getElementsByTagName("bottom")[0], themeList),
                    });
                });
            }

            const cx = stX.getElementsByTagName("cellXfs")[0];
            if (cx) {
                Array.from(cx.getElementsByTagName("xf")).forEach((x) => {
                    const numFmtId = +x.getAttribute("numFmtId") || 0;
                    const fontId = +x.getAttribute("fontId") || 0;
                    const fillId = +x.getAttribute("fillId") || 0;
                    const borderId = +x.getAttribute("borderId") || 0;
                    const al = x.getElementsByTagName("alignment")[0];
                    const fObj = fonts[fontId] || {};
                    const bdObj = borders[borderId] || {};
                    const st = { numFmtId };
                    if (fObj.b) st.b = true;
                    if (fObj.i) st.i = true;
                    if (fObj.u) st.u = true;
                    if (fObj.sz && fObj.sz !== 11) st.sz = fObj.sz;
                    if (fObj.name) st.font = fObj.name;
                    if (fObj.color && fObj.color !== "#000000") st.color = fObj.color;
                    if (fills[fillId]) st.fill = fills[fillId];
                    if (bdObj.bt) st.bt = bdObj.bt;
                    if (bdObj.br) st.br = bdObj.br;
                    if (bdObj.bb) st.bb = bdObj.bb;
                    if (bdObj.bl) st.bl = bdObj.bl;
                    if (al) {
                        const hz = al.getAttribute("horizontal");
                        const vt = al.getAttribute("vertical");
                        const wr = al.getAttribute("wrapText");
                        if (hz) st.align = hz;
                        if (vt) st.valign = vt;
                        if (wr === "1" || wr === "true") st.wrap = true;
                    }
                    xfs.push(st);
                });
            }
        }

        const out = [];
        for (const sh of Array.from(wbX.getElementsByTagName("sheet"))) {
            const rid = sh.getAttribute("r:id") || sh.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
            let tgt = rels[rid] || "";
            tgt = tgt.startsWith("/") ? tgt.slice(1) : "xl/" + tgt;
            const dx = await xml(tgt);
            const o = { name: sh.getAttribute("name"), cells: [], sparse: {}, merges: [], cols: [], rowH: {}, nrows: 0, ncols: 0 };

            if (dx) {
                Array.from(dx.getElementsByTagName("col")).forEach((c) => {
                    const w = c.getAttribute("hidden") === "1" ? 8 : Math.round((+c.getAttribute("width") || 8.43) * 7 + 5);
                    for (let i = +c.getAttribute("min"); i <= Math.min(+c.getAttribute("max"), 130); i++) o.cols[i - 1] = w;
                });

                const rawRows = Array.from(dx.getElementsByTagName("row"));
                const rawMerges = Array.from(dx.getElementsByTagName("mergeCell"));

                let maxContentRow = 14;
                rawRows.forEach((r) => {
                    const rn = +r.getAttribute("r");
                    if (rn <= 3000) maxContentRow = Math.max(maxContentRow, rn);
                });
                rawMerges.forEach((mc) => {
                    const parts = mc.getAttribute("ref").split(":");
                    parts.forEach((p) => {
                        const m = /([A-Z]+)(\d+)/.exec(p);
                        if (m && +m[2] <= 3000) maxContentRow = Math.max(maxContentRow, +m[2]);
                    });
                });

                const sharedFormulas = {};
                rawRows.forEach((r) => {
                    const rn = +r.getAttribute("r");
                    if (rn > maxContentRow + 10) {
                        Array.from(r.getElementsByTagName("c")).forEach((c) => {
                            const ref = c.getAttribute("r");
                            const t = c.getAttribute("t");
                            const vEl = c.getElementsByTagName("v")[0];
                            if (vEl && t === "s") o.sparse[ref] = sst[+vEl.textContent] || "";
                            else if (vEl) o.sparse[ref] = vEl.textContent;
                        });
                        return;
                    }
                    if (r.getAttribute("hidden") === "1") o.rowH[rn] = 2;
                    else if (r.getAttribute("ht")) o.rowH[rn] = Math.round(+r.getAttribute("ht") * 4 / 3);
                    o.nrows = Math.max(o.nrows, rn);

                    Array.from(r.getElementsByTagName("c")).forEach((c) => {
                        const m = /^([A-Z]+)(\d+)$/.exec(c.getAttribute("r"));
                        if (!m) return;
                        const ci = colIndex(m[1]);
                        if (ci > 130) return;

                        const sIdx = +(c.getAttribute("s") || 0);
                        const xfSt = xfs[sIdx] || {};
                        const t = c.getAttribute("t");
                        const vEl = c.getElementsByTagName("v")[0];
                        const fEl = c.getElementsByTagName("f")[0];

                        let cachedVal = null, disp = null, formulaStr = null;
                        if (fEl) {
                            const fText = (fEl.textContent || "").trim();
                            const si = fEl.getAttribute("si");
                            if (fText) {
                                formulaStr = "=" + fText;
                                if (si !== null) sharedFormulas[si] = fText;
                            } else if (si !== null && sharedFormulas[si]) {
                                formulaStr = "=" + sharedFormulas[si];
                            }
                        }

                        if (t === "inlineStr") {
                            cachedVal = Array.from(c.getElementsByTagName("t")).map((x) => x.textContent).join("");
                        } else if (vEl) {
                            const rawTxt = vEl.textContent;
                            if (t === "s") cachedVal = sst[+rawTxt];
                            else if (t === "str" || t === "e") cachedVal = rawTxt;
                            else if (t === "b") cachedVal = rawTxt === "1" ? "ИСТИНА" : "ЛОЖЬ";
                            else {
                                cachedVal = parseFloat(rawTxt);
                                const id = xfSt.numFmtId || 0;
                                const f = fmtExcel(cachedVal, codes[id], id);
                                if (f !== null) {
                                    if (/[dmyhs]/i.test((codes[id] || "").replace(/"[^"]*"|\[[^\]]*\]/g, "")) || (id >= 14 && id <= 22) || (id >= 45 && id <= 47)) {
                                        cachedVal = f;
                                    } else {
                                        disp = f;
                                    }
                                }
                            }
                        }

                        const hasStyle = Boolean(xfSt.fill || xfSt.bt || xfSt.br || xfSt.bb || xfSt.bl || xfSt.b || xfSt.i || xfSt.u || xfSt.color || xfSt.align);
                        if ((cachedVal === null || cachedVal === "" || cachedVal === undefined) && !formulaStr && !hasStyle) return;

                        const rawToStore = formulaStr || (cachedVal !== null && cachedVal !== undefined ? cachedVal : "");
                        const cleanSt = Object.keys(xfSt).length > 1 ? xfSt : null;
                        o.cells.push([+m[2], ci, rawToStore, disp !== null ? disp : (cachedVal !== null ? String(cachedVal) : null), cleanSt, cachedVal]);
                        o.ncols = Math.max(o.ncols, ci);
                    });
                });

                rawMerges.forEach((mc) => {
                    const [a, b] = mc.getAttribute("ref").split(":");
                    const ma = /([A-Z]+)(\d+)/.exec(a), mb = /([A-Z]+)(\d+)/.exec(b || a);
                    if (ma && mb && +mb[2] <= maxContentRow + 10) {
                        o.merges.push([+ma[2], colIndex(ma[1]), +mb[2], colIndex(mb[1])]);
                        o.ncols = Math.max(o.ncols, colIndex(mb[1]));
                    }
                });
            }
            out.push(o);
        }
        return out;
    }

    function parseCsv(text) {
        const first = text.split("\n")[0];
        const delim = (first.match(/;/g) || []).length >= (first.match(/,/g) || []).length ? ";" : ",";
        const rows = [];
        let row = [], cur = "", q = false;
        for (let i = 0; i < text.length; i++) {
            const ch = text[i];
            if (q) {
                if (ch === '"' && text[i + 1] === '"') { cur += '"'; i++; }
                else if (ch === '"') q = false;
                else cur += ch;
            } else if (ch === '"') q = true;
            else if (ch === delim) { row.push(cur); cur = ""; }
            else if (ch === "\n" || ch === "\r") {
                if (ch === "\r" && text[i + 1] === "\n") i++;
                row.push(cur); rows.push(row); row = []; cur = "";
            } else cur += ch;
        }
        if (cur || row.length) { row.push(cur); rows.push(row); }
        const cells = [];
        let nc = 0;
        rows.forEach((r, ri) => r.forEach((v, ci) => {
            if (v.trim() === "") return;
            const n = Number(v.replace(",", ".").replace(/\s/g, ""));
            const val = isNaN(n) ? v : n;
            cells.push([ri + 1, ci + 1, val, null, null, val]);
            nc = Math.max(nc, ci + 1);
        }));
        return [{ name: "Лист1", cells, sparse: {}, merges: [], cols: [], rowH: {}, nrows: rows.length, ncols: nc }];
    }

    // Создание новой пустой книги Excel с нуля
    function createNewExcelWorkbook() {
        S.imported = createBlankWorkbookObj("Книга1.xlsx");
        S.sourceB64 = null;
        S.impIdx = 0;
        S.edits = {};
        S.styles = {};
        workView = null;
        S.bossMode = true; S.stealthTabs = true;
        S.sheet = "imp0";
        render();
        select(1, 1);
        toast("Создана новая пустая книга «Книга1.xlsx» (Лист1)");
    }

    function renameWorkbook() {
        const curName = (S.imported && S.imported.name) || DEFAULT_WB_NAME;
        const next = window.prompt("Введите название книги Excel (.xlsx):", curName);
        if (next && next.trim()) {
            let clean = next.trim();
            if (!clean.toLowerCase().endsWith(".xlsx")) clean += ".xlsx";
            if (!S.imported) S.imported = createBlankWorkbookObj(clean);
            else S.imported.name = clean;
            render();
            toast("Книга переименована в «" + clean + "»");
        }
    }

    async function importFile(file) {
        if (!file) return;
        try {
            if (!/\.(xlsx|csv|tsv|txt)$/i.test(file.name)) throw new Error("Unsupported format");
            let sheets;
            if (/\.(csv|tsv|txt)$/i.test(file.name)) {
                sheets = parseCsv(await file.text());
                S.sourceB64 = null;
                S.imported = { name: file.name, fromScratch: true, sheets };
            } else {
                toast("Читаю файл Excel и стили…");
                const buf = await file.arrayBuffer();
                sheets = await parseXlsx(buf);
                if (!sheets.length) throw new Error("empty");
                S.sourceB64 = arrayBufferToBase64(buf);
                S.imported = { name: file.name, fromScratch: false, sheets };
            }
            if (!sheets.length) throw new Error("empty");
            S.impIdx = 0;
            S.edits = {};
            S.styles = {};
            workView = null;
            S.bossMode = true; S.stealthTabs = true;
            S.sheet = "imp0";
            render();
            select(1, 1);
            toast("Книга «" + file.name + "» загружена со всеми стилями и цветами");
        } catch (e) {
            console.error(e);
            toast("Не удалось прочитать файл. Поддерживаются .xlsx и .csv");
        }
    }

    function pickFile() { $("fileInput").click(); }

    // Объединение / разъединение выделенных ячеек при создании таблицы с нуля
    function toggleMergeCells() {
        if (!S.sheet.startsWith("imp") || !S.imported) {
            toast("Перейдите на Лист1 (или любой лист книги Excel), чтобы объединять ячейки");
            return;
        }
        const idx = +S.sheet.slice(3);
        const src = S.imported.sheets[idx];
        if (!src) return;
        const { r1, r2, c1, c2 } = rangeBounds();
        const existingIdx = src.merges.findIndex(([mr1, mc1, mr2, mc2]) => mr1 === r1 && mc1 === c1 && mr2 === r2 && mc2 === c2);
        if (existingIdx >= 0) {
            src.merges.splice(existingIdx, 1);
            render();
            toast("Объединение ячеек отменено");
            return;
        }
        if (r1 === r2 && c1 === c2) {
            toast("Выделите диапазон из нескольких ячеек мышью или с клавишей Shift");
            return;
        }
        // Удаляем пересекающиеся старые объединения
        src.merges = src.merges.filter(([mr1, mc1, mr2, mc2]) => mr2 < r1 || mr1 > r2 || mc2 < c1 || mc1 > c2);
        src.merges.push([r1, c1, r2, c2]);
        setStyleProp("align", "center");
        render();
        toast(`Ячейки ${refOf(r1, c1)}:${refOf(r2, c2)} объединены`);
    }

    // ------------------------------------------------------------------ Полноценный экспорт .xlsx со всеми стилями и цветами
    async function exportXlsx() {
        syncing();
        toast("Подготовка экспорта .xlsx со всеми стилями и цветами…");
        try {
            const filename = (S.imported && S.imported.name) || DEFAULT_WB_NAME;
            let payload;

            // Если пользователь загрузил свой .xlsx файл, патчим его бинарник, сохраняя 100% оригинальных тем и стилей
            if (S.imported && !S.imported.fromScratch && S.sourceB64) {
                payload = {
                    filename,
                    mode: "imported",
                    source_b64: S.sourceB64,
                    edits: S.edits,
                    styles: S.styles,
                };
            } else {
                // Если книга создана с нуля (Лист1, Лист2...) или экспортируется текущий лист
                const sheetsToExport = [];
                if (S.imported && S.imported.sheets && S.imported.sheets.length) {
                    S.imported.sheets.forEach((src, idx) => {
                        const sh = importedSheet(idx);
                        const cellsArr = [];
                        const allKeys = new Set([
                            ...Object.keys(sh.cells),
                            ...Object.keys(S.edits[sh.id] || {}),
                            ...Object.keys(S.styles[sh.id] || {}),
                        ]);
                        allKeys.forEach((k) => {
                            const [r, c] = k.split(":").map(Number);
                            const raw = rawAt(sh, r, c);
                            const disp = cellValue(sh.id, refOf(r, c));
                            const st = mergedCellStyle(sh, r, c);
                            if ((raw !== undefined && raw !== "") || st) {
                                cellsArr.push([r, c, raw !== undefined ? raw : "", disp, st]);
                            }
                        });
                        sheetsToExport.push({
                            name: src.name || `Лист${idx + 1}`,
                            cols: sh.cols,
                            rowH: sh.rowH || {},
                            merges: src.merges || [],
                            cells: cellsArr,
                        });
                    });
                } else {
                    const sh = getSheet(S.sheet);
                    const cellsArr = [];
                    Object.keys(sh.cells).forEach((k) => {
                        const [r, c] = k.split(":").map(Number);
                        const raw = rawAt(sh, r, c);
                        const disp = cellValue(sh.id, refOf(r, c));
                        const st = mergedCellStyle(sh, r, c);
                        if (raw !== undefined || st) {
                            cellsArr.push([r, c, raw !== undefined ? raw : "", disp, st]);
                        }
                    });
                    sheetsToExport.push({
                        name: sh.name || "Лист1",
                        cols: sh.cols,
                        rowH: sh.rowH || {},
                        merges: [],
                        cells: cellsArr,
                    });
                }
                payload = {
                    filename,
                    mode: "generated",
                    sheets: sheetsToExport,
                };
            }

            const res = await fetch("/api/export_xlsx", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload),
            });
            if (!res.ok) throw new Error("HTTP " + res.status);
            const blob = await res.blob();
            const a = document.createElement("a");
            a.href = URL.createObjectURL(blob);
            a.download = filename.endsWith(".xlsx") ? filename : filename + ".xlsx";
            a.click();
            URL.revokeObjectURL(a.href);
            toast("Книга «" + a.download + "» экспортирована со всеми цветами, рамками и формулами");
        } catch (e) {
            console.error(e);
            toast("Ошибка при экспорте .xlsx");
        }
    }

    // ------------------------------------------------------------------ рендер
    let cellEls = [];
    let curSheet = null;

    function styleOf(sheetId, r, c) {
        return (S.styles[sheetId] && S.styles[sheetId][key(r, c)]) || null;
    }

    function mergedCellStyle(sh, r, c) {
        const d = sh.cells[key(r, c)] || {};
        const baseSt = d.st || null;
        const userSt = styleOf(sh.id, r, c);
        if (!baseSt && !userSt) return null;
        return Object.assign({}, baseSt || {}, userSt || {});
    }

    function render() {
        if ($("guide")) document.body.append($("guide"));
        if (S.bossMode && workspaceActionPending) return;
        const sh = getSheet(S.sheet);
        curSheet = sh;
        const z = S.zoom / 100;
        document.documentElement.style.setProperty("--zoom", z);
        NCOLS = sh.ncols || 22;

        let h = "<colgroup><col style=\"width:" + Math.round(40 * z) + "px\">"
            + sh.cols.slice(0, NCOLS).map((w) => `<col style="width:${Math.round(w * z)}px">`).join("")
            + "</colgroup>";
        h += "<thead><tr><th class=\"corner\" title=\"Выделить все\"></th>"
            + LETTERS.slice(0, NCOLS).map((L, i) => `<th data-hc="${i + 1}">${L}</th>`).join("")
            + "</tr></thead><tbody>";

        const skip = sh.skip || {}, merges = sh.merges || {};
        for (let r = 1; r <= sh.rows; r++) {
            const rc = sh.rowCls[r] || "";
            const rh = sh.rowH && sh.rowH[r] ? ` style="--rh:${sh.rowH[r]}px"` : "";
            h += `<tr class="${rc}" ${rc.includes("opt") && !rc.includes("disabled") ? `data-opt="${r - 4}"` : ""}${rh}><th class="rn" data-hr="${r}">${r}</th>`;
            let coverEnd = 0;
            for (let c = 1; c <= NCOLS; c++) {
                if (skip[key(r, c)]) continue;
                const o = {}, mg = merges[key(r, c)];
                if (mg) {
                    o.cs = mg.cs;
                    o.rs = mg.rs;
                } else {
                    if (c < coverEnd) o.nb = true;
                    if (c > coverEnd) {
                        const d = sh.cells[key(r, c)] || {}, raw = rawAt(sh, r, c);
                        const st = mergedCellStyle(sh, r, c);
                        if (!d.html && !(st && st.wrap) && typeof raw === "string" && raw !== "" && !(raw[0] === "=" && !d.text) && !/hdr|story|optcell|prompt|\bc\b/.test(d.cls || "")) {
                            const need = raw.length * 7.4 + 14;
                            let w = sh.cols[c - 1], k = c;
                            while (need > w && k < NCOLS) {
                                const n = k + 1, nd = sh.cells[key(r, n)] || {}, nr = rawAt(sh, r, n);
                                if (nd.html || skip[key(r, n)] || merges[key(r, n)] || (nr !== undefined && nr !== "")) break;
                                k = n;
                                w += sh.cols[k - 1];
                            }
                            if (k > c) { coverEnd = k; o.spill = true; o.nb = true; }
                        }
                    }
                }
                h += tdHtml(sh, r, c, o);
            }
            h += "</tr>";
        }
        h += "</tbody>";
        $("grid").innerHTML = h;
        cellEls = [];
        $("grid").querySelectorAll("td").forEach((td) => {
            const r = +td.dataset.r, c = +td.dataset.c;
            (cellEls[r] = cellEls[r] || [])[c] = td;
        });
        renderChart(sh);
        renderTabs();
        updateSelection();
        renderRibbon();
        const ttl = (S.imported && S.imported.name) || DEFAULT_WB_NAME;
        document.title = ttl + " — Excel";
        $("docTitle").textContent = ttl;
        syncWorkspace();
    }

    function tdHtml(sh, r, c, o = {}) {
        const d = sh.cells[key(r, c)] || {};
        const st = mergedCellStyle(sh, r, c);
        const cls = [];
        let inner = "";
        if (d.cls) cls.push(d.cls);

        const hasUserEdit = Boolean(S.edits[sh.id] && S.edits[sh.id][key(r, c)] !== undefined);
        const raw = rawAt(sh, r, c);

        if (d.html) {
            inner = d.html();
        } else if (raw !== undefined && raw !== "") {
            const isF = !d.text && typeof raw === "string" && raw[0] === "=";
            const fmt = (st && st.fmt) || d.fmt;
            if (!isF && !hasUserEdit && d.disp !== undefined && d.disp !== null && !fmt) {
                inner = esc(d.disp);
                if (typeof d.val === "number") cls.push("num");
            } else {
                const v = isF ? cellValue(sh.id, refOf(r, c), 0) : parseInput(raw);
                inner = esc(formatValue(v, fmt));
                if (typeof v === "number") {
                    cls.push("num");
                    if (d.signed || fmt === "pct") {
                        if (d.signed && v < 0) cls.push("neg");
                        if (d.signed && v > 0) cls.push("pos");
                    }
                } else if (typeof v === "string" && v[0] === "#") {
                    cls.push("err");
                }
            }
        }

        const cssProps = [];
        if (st) {
            if (st.b) cls.push("bold");
            if (st.i) cls.push("ital");
            if (st.u) cls.push("und");
            if (st.wrap) cls.push("wrap");
            if (st.align === "center") cls.push("c");
            else if (st.align === "right") cls.push("r");
            else if (st.align === "left") cls.push("l");

            if (st.valign === "center") cls.push("vm");
            else if (st.valign === "top") cls.push("vt");
            else if (st.valign === "bottom") cls.push("vb");

            if (st.fill) cssProps.push(`background-color:${st.fill}`);
            if (st.color) cssProps.push(`color:${st.color}`);
            if (st.font) cssProps.push(`font-family:'${st.font.replace(/'/g, "")}',sans-serif`);
            if (st.sz) cssProps.push(`font-size:calc(${Math.round(st.sz * 1.18)}px * var(--zoom))`);
            if (st.bt) cssProps.push(`border-top:${st.bt}`);
            if (st.br) cssProps.push(`border-right:${st.br}`);
            if (st.bb) cssProps.push(`border-bottom:${st.bb}`);
            if (st.bl) cssProps.push(`border-left:${st.bl}`);
        }

        const styleAttr = cssProps.length ? ` style="${cssProps.join(";")}"` : "";
        const extra = (d.editableAction && S.custom ? " editing" : "")
            + (o.spill ? " spill" : "")
            + (o.nb ? " nb" : "")
            + (o.cs > 1 && !(st && st.align) ? " c" : "")
            + (o.rs > 1 && !(st && st.valign) ? " vm" : "");
        const span = (o.cs > 1 ? ` colspan="${o.cs}"` : "") + (o.rs > 1 ? ` rowspan="${o.rs}"` : "");
        return `<td data-r="${r}" data-c="${c}"${span} class="${cls.join(" ")}${extra}"${styleAttr}>${inner}</td>`;
    }

    function getTabsList() {
        const rpgTabs = S.stealthTabs ? [] : [["adv", "Приключение", true], ["log", "Журнал", true], ["inv", "Инвентарь", true]];
        const excelTabs = S.imported
            ? S.imported.sheets.map((x, i) => ["imp" + i, x.name.slice(0, 28), false])
            : [["boss", "Q4_Прогноз", false]];
        return rpgTabs.concat(excelTabs);
    }

    function renderTabs() {
        const tabs = getTabsList();
        $("sheetTabs").innerHTML = tabs.map(([id, n, isRpg]) =>
            `<button class="stab ${S.sheet === id ? "active" : ""} ${isRpg ? "rpg-tab" : ""}" data-sheet="${id}">${esc(n)}</button>`
        ).join("");
    }

    function scrollTabs(delta) {
        const box = $("sheetTabs");
        if (box) box.scrollBy({ left: delta, behavior: "smooth" });
    }

    function toggleSheetsMenu() {
        const pop = $("sheetsPopup");
        if (!pop.hidden) { pop.hidden = true; return; }
        const tabs = getTabsList();
        pop.innerHTML = tabs.map(([id, n]) =>
            `<button class="${S.sheet === id ? "active" : ""}" data-pop-sheet="${id}">${esc(n)}</button>`
        ).join("");
        pop.hidden = false;
    }

    function addSheet() {
        if (!S.imported) {
            S.imported = createBlankWorkbookObj(DEFAULT_WB_NAME);
        }
        const newIdx = S.imported.sheets.length;
        const newName = "Лист" + (newIdx + 1);
        S.imported.sheets.push({
            name: newName,
            cells: [],
            sparse: {},
            merges: [],
            cols: Array.from({ length: 26 }, () => 80),
            rowH: {},
            nrows: 60,
            ncols: 22,
        });
        gotoSheet("imp" + newIdx);
        toast("Добавлен новый пустой лист «" + newName + "»");
    }

    function renderChart(sh) {
        const box = $("chartObject");
        if (!sh.chart) { box.hidden = true; return; }
        box.hidden = false;
        const z = S.zoom / 100;
        box.style.left = Math.round(70 * z) + "px";
        box.style.top = Math.round((S.log.length + 3) * 24 * z + 30) + "px";
        const H = S.history, W = 450, HT = 210, pl = 28, pb = 18, pt = 8;
        const x = (i) => pl + (H.length < 2 ? 0 : (i * (W - pl - 8)) / (H.length - 1));
        const y = (v) => pt + (HT - pt - pb) * (1 - v / 100);
        const line = (k, col) => `<polyline fill="none" stroke="${col}" stroke-width="2.2" stroke-linejoin="round" points="${H.map((p, i) => x(i) + "," + y(p[k])).join(" ")}"/>`
            + H.map((p, i) => `<circle cx="${x(i)}" cy="${y(p[k])}" r="2.6" fill="${col}"/>`).join("");
        let grid = "";
        [0, 25, 50, 75, 100].forEach((v) => {
            grid += `<line x1="${pl}" x2="${W - 4}" y1="${y(v)}" y2="${y(v)}" stroke="var(--grid-line)"/><text x="${pl - 5}" y="${y(v) + 3.5}" text-anchor="end" font-size="10" fill="var(--text-dim)">${v}</text>`;
        });
        box.innerHTML = `<h4>Динамика параметров героя по ходам</h4><svg viewBox="0 0 ${W} ${HT}" preserveAspectRatio="none">${grid}${H.length > 1 ? line("ROI", "#107c41") + line("KPI", "#2f6fdb") + line("Anxiety", "#d13438") : `<text x="${W / 2}" y="${HT / 2}" text-anchor="middle" fill="var(--text-dim)" font-size="12">Сделайте первый ход — график оживёт</text>`}</svg>`
            + `<div class="legend"><span><i style="background:#107c41"></i>Здоровье</span><span><i style="background:#2f6fdb"></i>Бодрость</span><span><i style="background:#d13438"></i>Неприятности</span></div>`;
    }

    // ------------------------------------------------------------------ выделение
    function rangeBounds() {
        let r1 = Math.min(S.anchor.r, S.cur.r), r2 = Math.max(S.anchor.r, S.cur.r);
        let c1 = Math.min(S.anchor.c, S.cur.c), c2 = Math.max(S.anchor.c, S.cur.c);
        if (curSheet && curSheet.merges) {
            const mg = curSheet.merges[key(r1, c1)];
            if (mg && r1 === r2 && c1 === c2) {
                r2 = Math.max(r2, mg.r2);
                c2 = Math.max(c2, mg.c2);
            }
        }
        return { r1, r2, c1, c2 };
    }

    function updateSelection() {
        const { r1, r2, c1, c2 } = rangeBounds();
        document.querySelectorAll("#grid .sel, #grid .inrange, #grid .anchor, #grid .fill-handle").forEach((e) =>
            e.classList.remove("sel", "inrange", "anchor", "fill-handle")
        );
        document.querySelectorAll("#grid .hsel").forEach((e) => e.classList.remove("hsel"));

        const multi = r1 !== r2 || c1 !== c2;
        let lastTd = null;
        for (let r = r1; r <= r2; r++) {
            for (let c = c1; c <= c2; c++) {
                const td = cellEls[r] && cellEls[r][c];
                if (td) {
                    if (multi) td.classList.add("inrange");
                    lastTd = td;
                }
            }
        }
        const a = cellEls[S.anchor.r] && cellEls[S.anchor.r][S.anchor.c];
        if (a) {
            a.classList.add("sel", "anchor");
            if (!multi) a.classList.add("fill-handle");
        }
        if (multi && lastTd) {
            lastTd.classList.add("fill-handle");
        }

        for (let c = c1; c <= c2; c++) {
            const th = document.querySelector(`#grid th[data-hc="${c}"]`);
            if (th) th.classList.add("hsel");
        }
        for (let r = r1; r <= r2; r++) {
            const th = document.querySelector(`#grid th[data-hr="${r}"]`);
            if (th) th.classList.add("hsel");
        }
        $("nameBox").value = multi && !(curSheet && curSheet.merges && curSheet.merges[key(S.anchor.r, S.anchor.c)])
            ? `${refOf(r1, c1)}:${refOf(r2, c2)}`
            : refOf(S.anchor.r, S.anchor.c);
        syncFormulaBar();
        updateStatusStats();
    }

    function isLocked(r, c) {
        const sh = curSheet, d = sh.cells[key(r, c)];
        if (d && d.editableAction) return false;
        if (d && d.locked) return true;
        return r < sh.lockedFrom;
    }

    function syncFormulaBar() {
        const { r, c } = S.anchor, sh = curSheet, d = sh.cells[key(r, c)] || {};
        const fi = $("formulaInput");
        let v = "";
        if (d.editableAction) v = S.custom;
        else if (d.fx !== undefined) v = typeof d.fx === "function" ? d.fx() : d.fx;
        else {
            const raw = rawAt(sh, r, c);
            v = raw === undefined ? "" : String(raw);
        }
        fi.value = v;
        fi.readOnly = isLocked(r, c);
        fi.placeholder = d.editableAction ? "Введите своё действие и нажмите Enter…" : "";
    }

    function updateStatusStats() {
        const { r1, r2, c1, c2 } = rangeBounds();
        const nums = [];
        for (let r = r1; r <= r2; r++) {
            for (let c = c1; c <= c2; c++) {
                const d = curSheet.cells[key(r, c)] || {};
                if (d.html) continue;
                const v = cellValue(curSheet.id, refOf(r, c));
                if (typeof v === "number" && Number.isFinite(v)) nums.push(v);
            }
        }
        const el = $("sbStats");
        if (nums.length < 2) { el.innerHTML = ""; return; }
        const sum = nums.reduce((a, b) => a + b, 0);
        el.innerHTML = `<span>Среднее: <b>${nf(sum / nums.length, { maximumFractionDigits: 2 })}</b></span><span>Количество: <b>${nums.length}</b></span><span>Сумма: <b>${nf(sum, { maximumFractionDigits: 2 })}</b></span>`;
    }

    function select(r, c, extend) {
        r = Math.max(1, Math.min(curSheet.rows, r));
        c = Math.max(1, Math.min(NCOLS, c));
        if (curSheet && curSheet.skip && curSheet.skip[key(r, c)]) {
            const anchorKey = curSheet.skip[key(r, c)];
            const [ar, ac] = anchorKey.split(":").map(Number);
            r = ar;
            c = ac;
        }
        if (extend) S.cur = { r, c };
        else { S.anchor = { r, c }; S.cur = { r, c }; }
        updateSelection();
        const td = cellEls[r] && cellEls[r][c];
        if (td) td.scrollIntoView({ block: "nearest", inline: "nearest" });
    }

    function jumpToNameBox(val) {
        const clean = String(val || "").trim().toUpperCase();
        const m = /^([A-Z]{1,3})(\d{1,7})$/.exec(clean);
        if (!m) { updateSelection(); return; }
        const c = colIndex(m[1]), r = +m[2];
        if (curSheet && curSheet.sparse && curSheet.sparse[clean] !== undefined) {
            $("formulaInput").value = curSheet.sparse[clean];
            toast(`${clean}: ${curSheet.sparse[clean]}`);
            return;
        }
        select(r, c, false);
    }

    // ------------------------------------------------------------------ редактирование
    function commitFormula() {
        const { r, c } = S.anchor, d = curSheet.cells[key(r, c)] || {};
        const val = $("formulaInput").value;
        if (d.editableAction) {
            const t = val.trim();
            if (t) { S.custom = ""; sendAction(t, -1); }
            return;
        }
        if (isLocked(r, c)) return;
        (S.edits[S.sheet] = S.edits[S.sheet] || {})[key(r, c)] = val.trim() === "" ? undefined : val;
        if (val.trim() === "") delete S.edits[S.sheet][key(r, c)];
        if (val.trim() === "" && curSheet.cells[key(r, c)] && curSheet.cells[key(r, c)].raw !== undefined) {
            S.edits[S.sheet][key(r, c)] = "";
        }
        syncing();
        render();
        select(r + 1, c);
        $("formulaInput").blur();
    }

    function cancelEdit() {
        syncFormulaBar();
        $("formulaInput").blur();
    }

    function insertFn() {
        const f = $("formulaInput");
        if (!f.readOnly) {
            if (!f.value.startsWith("=")) f.value = "=СУММ()";
            f.focus();
        } else {
            toast("Выберите редактируемую ячейку для вставки функции");
        }
    }

    // ------------------------------------------------------------------ игра
    function logTurn(action) {
        S.log.push({ turn: S.log.length + 1, time: clockStr(), action, roi: S.stats.ROI, kpi: S.stats.KPI, anx: S.stats.Anxiety, formula: S.auditLog });
        S.history.push({ ...S.stats });
    }

    function syncing() {
        const el = $("saveStatus");
        if (!el) return;
        el.classList.add("syncing");
        setTimeout(() => { el.classList.remove("syncing"); }, 500);
    }

    function applyData(data, action) {
        const wasOver = S.over;
        const prevItems = S.inventory.length;
        S.narrative = data.narrative;
        S.stats = data.stats || { ...START_STATS };
        S.options = data.options || [];
        S.over = Boolean(data.game_over) || S.stats.ROI <= 0 || S.stats.KPI <= 0 || S.stats.Anxiety >= 100;
        S.auditLog = data.audit_log || "";
        S.sceneId = data.scene_id || null;
        if (S.sceneId) S.seen.push(S.sceneId);
        if (data.inventory_add) S.inventory.push(data.inventory_add);
        S.loading = null;
        S.custom = "";
        if (action) logTurn(action);
        if (!S.over && clockMinutes() >= 18 * 60) {
            S.over = true;
            S.won = true;
            S.narrative = `🎉 18:00 — ВЕЧЕРНИЙ ПРАЗДНИК В ОУКХЕЙВЕНЕ! Вы отлично провели день, исследовали долину и заслужили почётное место в Гильдии у мастера ${S.boss}! Итоговый счёт: ${score()}.`;
            S.options = [];
        }
        if (S.over) {
            const sc = S.won ? score() : Math.round(score() / 2);
            if (S.turn > 1 && sc > S.best) { S.best = sc; store.set("te_best", String(sc)); }
        }
        if (S.over && !wasOver && adventureId) {
            const before = DailyAllowance.remaining(store);
            const remaining = DailyAllowance.complete(store, adventureId);
            if (before > 0 && remaining === 0) offerPending = true;
        }
        if (S.inventory.length > prevItems) toast("🎒 Получен предмет: " + S.inventory[S.inventory.length - 1]);
        render();
        const td = cellEls[2] && cellEls[2][2];
        if (td) td.classList.add("fade");
        if (offerPending && !offerDismissed && !S.bossMode) offerSubscription();
    }

    async function post(url, body) {
        const res = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
        if (!res.ok) throw new Error("HTTP " + res.status);
        return res.json();
    }

    async function restart() {
        if (S.loading) return;
        adventureId = crypto.randomUUID();
        offerPending = false;
        offerDismissed = false;
        Object.assign(S, {
            stats: { ...START_STATS }, options: [], turn: 1, over: false, won: false,
            narrative: "", auditLog: "", sceneId: null, seen: [], inventory: [], log: [],
            history: [{ ...START_STATS }], custom: "",
        });
        S.loading = "Загрузка нового приключения…";
        S.bossMode = false;
        S.stealthTabs = false;
        S.sheet = "adv";
        render();
        select(2, 2);
        syncing();
        try {
            const data = await post("/api/start", { role: S.role, boss: S.boss, api_key: S.apiKey });
            applyData(data, null);
        } catch {
            S.loading = null;
            S.narrative = "⚠ Нет связи с сервером. Нажмите «Новая игра» ещё раз.";
            render();
        }
    }

    async function sendAction(text, idx) {
        if (S.loading || S.over || S.bossMode) return;
        const prev = { narrative: S.narrative, options: S.options };
        S.loading = "Выполняется ход: " + text + " …";
        render();
        syncing();
        try {
            const data = await post("/api/action", {
                action: text, stats: S.stats, role: S.role, boss: S.boss, api_key: S.apiKey,
                choice_idx: idx, scene_id: S.sceneId, seen: S.seen.slice(-12), clock: clockStr(), inventory: S.inventory,
            });
            S.turn += 1;
            workspaceActionPending = S.bossMode;
            applyData(data, text);
            workspaceActionPending = false;
            return true;
        } catch {
            S.loading = null;
            S.narrative = prev.narrative;
            S.options = prev.options;
            if (!S.bossMode) render();
            toast("⚠ Ошибка связи с сервером. Повторите ход.");
            return false;
        }
    }

    function choose(i) {
        if (S.sheet !== "adv" || S.loading) return;
        if (S.over) {
            if (i === 0) restart();
            else if (i === 2) exportReport();
            return;
        }
        if (S.options[i]) sendAction(S.options[i], i);
    }

    async function useItem(name) {
        if (S.over || S.loading) { toast("Сейчас предметы использовать нельзя"); return; }
        const i = S.inventory.indexOf(name);
        if (i < 0) return;
        S.inventory.splice(i, 1);
        S.sheet = "adv";
        const applied = await sendAction("Использовать предмет: " + name, -1);
        if (!applied) { S.inventory.splice(i, 0, name); if (!S.bossMode) render(); }
    }

    function exportReport() {
        const lines = [
            "TabEscape — Дневник приключений Оукхейвена", `Герой: ${S.role}`, `Мастер Гильдии: ${S.boss}`,
            `Исход: ${S.won ? "успешное завершение дня" : "приключение прервано"}`, `Очки: ${score()} · Ходов: ${S.log.length}`, "",
            ...S.log.map((e) => `${String(e.turn).padStart(2, "0")} [${e.time}] ${e.action}  (Здоровье ${e.roi} / Бодрость ${e.kpi} / Неприятности ${e.anx})`), "", "Финал: " + S.narrative,
        ];
        const a = document.createElement("a");
        a.href = URL.createObjectURL(new Blob([lines.join("\n")], { type: "text/plain;charset=utf-8" }));
        a.download = "adventure_diary.txt";
        a.click();
        URL.revokeObjectURL(a.href);
        toast("Дневник похода сохранён");
    }

    // ------------------------------------------------------------------ босс-режим и навигация по листам
    function bossTarget() {
        return S.imported ? "imp" + Math.min(S.impIdx, S.imported.sheets.length - 1) : "boss";
    }

    let gameView = null;
    let workView = null;
    function setBoss(on) {
        if (on === S.bossMode) return;
        const snapshot = {sheet: S.sheet, cur: {...S.cur}, anchor: {...S.anchor},
            x: $("gridScroll").scrollLeft, y: $("gridScroll").scrollTop,
            focus: document.activeElement, focusId: document.activeElement.id, scroll: window.scrollY};
        if (on) gameView = snapshot; else workView = snapshot;
        S.bossMode = on;
        S.stealthTabs = on;
        if (on) {
            S.prevSheet = S.sheet;
            S.sheet = workView && isBossSheet(workView.sheet) ? workView.sheet : bossTarget();
        } else S.sheet = gameView ? gameView.sheet : "adv";
        $("toast").classList.remove("show");
        if (tourStep === 3 && on) tourVisitedWork = true;
        render();
        const target = on ? workView : gameView;
        if (target) {
            S.cur = target.cur; S.anchor = target.anchor;
            updateSelection();
            $("gridScroll").scrollTo(target.x, target.y);
            const focus = target.focusId ? $(target.focusId) : target.focus;
            if (focus && focus.isConnected) focus.focus({preventScroll:true});
            window.scrollTo(0, target.scroll);
        }
    }

    function toggleBoss() {
        setBoss(!S.bossMode);
    }

    function gotoSheet(id) {
        $("sheetsPopup").hidden = true;
        if (isBossSheet(id)) {
            if (!isBossSheet(S.sheet)) S.prevSheet = S.sheet;
            if (id.startsWith("imp")) S.impIdx = +id.slice(3);
        } else {
            S.bossMode = false;
            S.stealthTabs = false;
        }
        S.sheet = id;
        render();
        $("gridScroll").scrollTo(0, 0);
        select(id === "adv" ? 2 : 1, id === "adv" ? 2 : 1);
    }

    // ------------------------------------------------------------------ лента (Office 365 Excel Web)
    const TABS = [
        "Файл", "Главная", "Вставка", "Поделиться", "Разметка страницы",
        "Формулы", "Данные", "Рецензирование", "Вид", "Автоматизация", "Справка", "Рисование",
    ];
    const TAB_IDS = [
        "file", "home", "insert", "share", "layout",
        "formulas", "data", "review", "view", "automate", "help", "draw",
    ];
    const TIPS = [
        "Совет: нажмите «Новая книга» или «+» внизу, чтобы создать чистую таблицу Excel с нуля.",
        "Совет: нажмите «Открыть .xlsx» или просто перетащите любой файл Excel в окно браузера.",
        "Совет: спокойные вопросы жителям и выбор дороги не отнимают ни Здоровье, ни Бодрость!",
    ];

    function selectedRange() { return rangeBounds(); }
    function forRange(fn) {
        const { r1, r2, c1, c2 } = selectedRange();
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) fn(r, c);
    }

    function toggleStyle(k) {
        const st = S.styles[S.sheet] = S.styles[S.sheet] || {};
        let all = true;
        forRange((r, c) => {
            const cur = mergedCellStyle(curSheet, r, c);
            if (!(cur && cur[k])) all = false;
        });
        forRange((r, c) => {
            const o = st[key(r, c)] = st[key(r, c)] || {};
            o[k] = !all;
        });
        render();
    }

    function setStyleProp(prop, val) {
        const st = S.styles[S.sheet] = S.styles[S.sheet] || {};
        forRange((r, c) => {
            const o = st[key(r, c)] = st[key(r, c)] || {};
            o[prop] = val;
            if (prop === "border") {
                const css = val === "none" ? "" : (val === "medium" ? "1.5px solid #000" : "1px solid #000");
                o.bt = css; o.br = css; o.bb = css; o.bl = css;
            }
        });
        render();
    }

    function setFmt(f) { setStyleProp("fmt", f); }

    function stOn(k) {
        const o = curSheet ? mergedCellStyle(curSheet, S.anchor.r, S.anchor.c) : null;
        return Boolean(o && o[k]);
    }

    function curFontName() {
        const o = curSheet ? mergedCellStyle(curSheet, S.anchor.r, S.anchor.c) : null;
        return (o && o.font) || "Calibri";
    }

    function curFontSize() {
        const o = curSheet ? mergedCellStyle(curSheet, S.anchor.r, S.anchor.c) : null;
        return String(Math.round((o && o.sz) || 11));
    }

    function copyCell() {
        const td = cellEls[S.anchor.r] && cellEls[S.anchor.r][S.anchor.c];
        try {
            navigator.clipboard.writeText(td ? td.innerText : "");
            toast("Скопировано в буфер обмена");
        } catch {
            toast("Не удалось скопировать");
        }
    }

    function autosum() {
        const { r, c } = S.anchor;
        if (isLocked(r, c) || r < 2) { toast("Выберите ячейку под числами"); return; }
        $("formulaInput").readOnly = false;
        $("formulaInput").value = `=СУММ(${refOf(Math.max(1, r - 5), c)}:${refOf(r - 1, c)})`;
        $("formulaInput").focus();
    }

    function toggleTheme() {
        S.theme = S.theme === "dark" ? "light" : "dark";
        document.documentElement.dataset.theme = S.theme;
        store.set("te_theme", S.theme);
        render();
    }

    function toggleRibbon() {
        S.ribbonCollapsed = !S.ribbonCollapsed;
        $("ribbonBody").classList.toggle("collapsed", S.ribbonCollapsed);
        const path = $("ribbonTogglePath");
        if (path) path.setAttribute("d", S.ribbonCollapsed ? "M3.5 6L8 10.5 12.5 6" : "M3.5 10L8 5.5 12.5 10");
    }

    function ribbonDef() {
        const G = (title, items) => ({ title, items });
        const B = (label, ico, act, extra = {}) => Object.assign({ label, ico, act }, extra);
        switch (S.bossMode ? "home" : S.tab) {
            case "file": return [
                G("Книга Excel", [
                    B("Новая книга (с нуля)", "＋", createNewExcelWorkbook, { primary: true, tip: "Создать чистую пустую книгу Excel с нуля" }),
                    B("Открыть .xlsx", "⤒", pickFile, { tip: "Загрузить любой файл .xlsx или .csv с компьютера" }),
                    B("Экспорт .xlsx", "⤓", exportXlsx, { tip: "Скачать книгу .xlsx со всеми цветами, рамками и формулами" }),
                    B("Переименовать", "✎", renameWorkbook, { tip: "Изменить имя файла .xlsx" }),
                ]),
                G("Игра и отчёт", [
                    B("Новое приключение", "▶", () => { S.sheet = "adv"; restart(); }, {}),
                    B("Дневник (.txt)", "▤", exportReport, {}),
                    B("Параметры ИИ", "⚙", openSettings, {}),
                ]),
            ];
            case "insert": return [
                G("Таблицы и листы", [
                    B("Новый пустой лист", "＋", addSheet, { primary: true }),
                    B("Загрузить .xlsx", "⤒", pickFile, {}),
                    B("Объединить ячейки", "⊞", toggleMergeCells, {}),
                ]),
                G("Диаграммы", [
                    B("График похода", "⌁", () => gotoSheet("log"), {}),
                ]),
            ];
            case "share": return [
                G("Экспорт и сохранение", [
                    B("Скачать книгу (.xlsx)", "⤓", exportXlsx, { primary: true, tip: "Сохранить .xlsx со всеми стилями и цветами" }),
                    B("Скачать дневник (.txt)", "▤", exportReport, {}),
                ]),
            ];
            case "layout": return [
                G("Тема", [
                    B(S.theme === "dark" ? "Светлая тема" : "Тёмная тема", "◐", toggleTheme, {}),
                    B("Сетка", "▦", () => toast("Линии сетки отображаются"), { on: true }),
                ]),
            ];
            case "formulas": return [
                G("Функции", [
                    B("Автосумма", "Σ", autosum, {}),
                    B("Вставить fx", "fx", insertFn, {}),
                    B("Пересчитать", "↻", () => { render(); toast("Все формулы пересчитаны"); }, {}),
                ]),
            ];
            case "data": return [
                G("Управление данными", [
                    B("Новая книга", "＋", createNewExcelWorkbook, {}),
                    B("Открыть .xlsx", "⤒", pickFile, { primary: true }),
                    B("Экспорт .xlsx", "⤓", exportXlsx, {}),
                ]),
            ];
            case "review": return [
                G("Проверка", [
                    B("Статистика книги", "📊", showWorkbookStats, {}),
                    B("Совет", "?", () => toast(TIPS[Math.floor(Math.random() * TIPS.length)]), {}),
                    B("Рекорд: " + S.best, "★", () => toast("Лучший результат: " + S.best + " очков"), {}),
                ]),
            ];
            case "view": return [
                G("Режимы просмотра", [
                    B("Босс-режим (Esc)", "👁", toggleBoss, { on: S.bossMode }),
                    B(S.theme === "dark" ? "Светлая тема" : "Тёмная тема", "◐", toggleTheme, {}),
                    B("100%", "⌕", () => setZoom(100), {}),
                    B("Во весь экран", "⛶", toggleFullscreen, {}),
                ]),
            ];
            default: return [
                G("Книга", [
                    B("Новая книга", "＋", createNewExcelWorkbook, { tip: "Создать чистую пустую книгу Excel с нуля" }),
                    B("Открыть", "⤒", pickFile, { tip: "Загрузить свой .xlsx файл" }),
                    B("Экспорт .xlsx", "⤓", exportXlsx, { primary: true, tip: "Сохранить книгу Excel (.xlsx) со всеми цветами, рамками и формулами" }),
                ]),
                G("Шрифт и цвета", [
                    { fontSelect: curFontName() },
                    { sizeSelect: curFontSize() },
                    B("", "<b>Ж</b>", () => toggleStyle("b"), { raw: true, on: stOn("b"), tip: "Полужирный" }),
                    B("", "<i>К</i>", () => toggleStyle("i"), { raw: true, on: stOn("i"), tip: "Курсив" }),
                    B("", "<u>Ч</u>", () => toggleStyle("u"), { raw: true, on: stOn("u"), tip: "Подчёркнутый" }),
                    B("", "▦", () => setStyleProp("border", "thin"), { tip: "Все границы (тонкие)" }),
                    B("", "▣", () => setStyleProp("border", "medium"), { tip: "Толстая внешняя граница" }),
                    { swatchFill: S.fillColor },
                    { swatchFont: S.fontColor },
                ]),
                G("Выравнивание", [
                    B("", "⇤", () => setStyleProp("align", "left"), { tip: "По левому краю" }),
                    B("", "⇔", () => setStyleProp("align", "center"), { tip: "По центру" }),
                    B("", "⇥", () => setStyleProp("align", "right"), { tip: "По правому краю" }),
                    B("", "↵", () => toggleStyle("wrap"), { on: stOn("wrap"), tip: "Перенос текста" }),
                    B("Объединить", "⊞", toggleMergeCells, { tip: "Объединить и поместить в центре выделенные ячейки" }),
                ]),
                G("Число", [
                    B("", "₽", () => setFmt("money"), { tip: "Денежный формат" }),
                    B("", "%", () => setFmt("pct"), { tip: "Процентный формат" }),
                    B("", "000", () => setFmt("int"), { tip: "Целый с разделителями" }),
                    B("", ".00", () => setFmt("dec"), { tip: "Два знака после запятой" }),
                ]),
                G("Режим", [
                    B("Игра", "🧭", () => gotoSheet("adv"), { on: S.sheet === "adv", tip: "Открыть вкладку Приключение" }),
                    B("Лист Excel", "📋", () => gotoSheet("imp" + S.impIdx), { on: S.sheet.startsWith("imp"), tip: "Перейти к редактированию таблицы Excel (Лист1)" }),
                ]),
            ];
        }
    }

    let ribbonActs = [];
    function renderRibbon() {
        $("ribbonTabs").innerHTML = TABS.map((n, i) =>
            `<button class="rtab ${TAB_IDS[i] === S.tab ? "active" : ""}" data-tab="${TAB_IDS[i]}">${n}</button>`
        ).join("");
        ribbonActs = [];
        $("ribbonBody").innerHTML = ribbonDef().filter(g => !S.bossMode || g.title !== "Режим").map((g) => `<div class="rgroup"><div class="rg-items">` + g.items.map((it) => {
            if (it.fontSelect) {
                const fonts = ["Calibri", "Times New Roman", "Arial", "Arial Cyr", "Segoe UI"];
                return `<select class="rselect-input" data-font-select="1" title="Шрифт">${fonts.map((f) => `<option value="${f}" ${f === it.fontSelect ? "selected" : ""}>${f}</option>`).join("")}</select>`;
            }
            if (it.sizeSelect) {
                const sizes = ["10", "11", "12", "13", "14", "15", "16", "18", "20", "22"];
                return `<select class="rselect-input" data-size-select="1" title="Размер шрифта">${sizes.map((s) => `<option value="${s}" ${s === it.sizeSelect ? "selected" : ""}>${s}</option>`).join("")}</select>`;
            }
            if (it.swatchFill) {
                return `<button class="swatch-btn" data-fill-apply="1" title="Залить выбранным цветом"><span>🪣</span><i class="swatch-bar" style="background:${it.swatchFill}"></i></button>`
                    + `<button class="icon-btn tiny" data-fill-pick="1" title="Выбрать любой цвет заливки">▾</button>`;
            }
            if (it.swatchFont) {
                return `<button class="swatch-btn" data-fg-apply="1" title="Применить цвет текста"><b>A</b><i class="swatch-bar" style="background:${it.swatchFont}"></i></button>`
                    + `<button class="icon-btn tiny" data-fg-pick="1" title="Выбрать цвет текста">▾</button>`;
            }
            ribbonActs.push(it.act);
            const idx = ribbonActs.length - 1;
            return `<button class="rbtn ${it.on ? "on" : ""} ${it.primary ? "primary" : ""}" data-ract="${idx}" title="${esc(it.tip || it.label || "")}"><span class="ico">${it.raw ? it.ico : esc(it.ico)}</span>${it.label ? `<span>${esc(it.label)}</span>` : ""}</button>`;
        }).join("") + `</div></div>`).join("");
    }

    // ------------------------------------------------------------------ вспомогательные действия UI
    let toastTimer = null;
    function toast(msg) {
        if (S.bossMode) return;
        const t = $("toast");
        t.textContent = msg;
        t.classList.add("show");
        clearTimeout(toastTimer);
        toastTimer = setTimeout(() => t.classList.remove("show"), 2600);
    }

    function setZoom(z) {
        S.zoom = Math.max(60, Math.min(160, z));
        $("zoomRange").value = S.zoom;
        $("zoomLabel").textContent = S.zoom + " % ▾";
        render();
    }
    function zoomBy(d) { setZoom(S.zoom + d); }
    function resetZoom() { setZoom(100); }

    function toggleFullscreen() {
        if (!document.fullscreenElement) {
            document.documentElement.requestFullscreen().catch(() => {});
        } else {
            document.exitFullscreen().catch(() => {});
        }
    }

    function showWorkbookStats() {
        const shCount = S.imported ? S.imported.sheets.length : 1;
        const cellCount = curSheet ? Object.keys(curSheet.cells).length + Object.keys(S.edits[curSheet.id] || {}).length : 0;
        toast(`Статистика книги: листов Excel — ${shCount}, активных ячеек на листе — ${cellCount}`);
    }

    function openSettings() {
        $("inputRole").value = S.role;
        $("inputBoss").value = S.boss;
        $("inputApiKey").value = S.apiKey;
        $("settingsModal").hidden = false;
        $("inputRole").focus();
    }
    function closeSettings() { $("settingsModal").hidden = true; $("switchView")?.focus(); }
    function saveSettings() {
        S.role = $("inputRole").value.trim() || "Начинающий искатель";
        S.boss = $("inputBoss").value.trim() || "Мастер Гильдии Морган";
        S.apiKey = $("inputApiKey").value.trim();
        store.set("te_role", S.role);
        store.set("te_boss", S.boss);
        session.set("te_key", S.apiKey);
        closeSettings();
        restart();
    }

    // ------------------------------------------------------------------ события
    function bind() {
        const grid = $("grid");
        grid.addEventListener("mousedown", (e) => {
            $("sheetsPopup").hidden = true;
            const td = e.target.closest("td"), th = e.target.closest("th");
            if (e.target.closest("button")) return;
            if (td) {
                S.drag = true;
                select(+td.dataset.r, +td.dataset.c, e.shiftKey);
                if (!e.shiftKey) $("formulaInput").blur();
            } else if (th && th.dataset.hc) {
                S.anchor = { r: 1, c: +th.dataset.hc };
                S.cur = { r: curSheet.rows, c: +th.dataset.hc };
                updateSelection();
            } else if (th && th.dataset.hr) {
                S.anchor = { r: +th.dataset.hr, c: 1 };
                S.cur = { r: +th.dataset.hr, c: NCOLS };
                updateSelection();
            } else if (th && th.classList.contains("corner")) {
                S.anchor = { r: 1, c: 1 };
                S.cur = { r: curSheet.rows, c: NCOLS };
                updateSelection();
            }
        });
        grid.addEventListener("mouseover", (e) => {
            if (!S.drag) return;
            const td = e.target.closest("td");
            if (td) { S.cur = { r: +td.dataset.r, c: +td.dataset.c }; updateSelection(); }
        });
        document.addEventListener("mouseup", () => { S.drag = false; });

        grid.addEventListener("click", (e) => {
            const use = e.target.closest("[data-use]");
            if (use) { useItem(use.dataset.use); return; }
            const row = e.target.closest("tr[data-opt]");
            if (row && S.sheet === "adv" && e.target.closest("td") && rangeBounds().r1 === rangeBounds().r2 && rangeBounds().c1 === rangeBounds().c2) {
                choose(+row.dataset.opt);
            }
        });
        grid.addEventListener("dblclick", (e) => {
            const td = e.target.closest("td");
            if (!td) return;
            const f = $("formulaInput");
            if (!f.readOnly) { f.focus(); f.setSelectionRange(f.value.length, f.value.length); }
        });

        $("ribbon").addEventListener("click", (e) => {
            const t = e.target.closest("[data-tab]");
            if (t) {
                if (S.ribbonCollapsed) toggleRibbon();
                S.tab = t.dataset.tab;
                renderRibbon();
                return;
            }
            const b = e.target.closest("[data-ract]");
            if (b) { ribbonActs[+b.dataset.ract](); return; }
            if (e.target.closest("[data-fill-apply]")) { setStyleProp("fill", S.fillColor); return; }
            if (e.target.closest("[data-fill-pick]")) { $("bgColorPicker").click(); return; }
            if (e.target.closest("[data-fg-apply]")) { setStyleProp("color", S.fontColor); return; }
            if (e.target.closest("[data-fg-pick]")) { $("fgColorPicker").click(); return; }
        });

        $("ribbon").addEventListener("change", (e) => {
            if (e.target.matches("[data-font-select]")) {
                setStyleProp("font", e.target.value);
            } else if (e.target.matches("[data-size-select]")) {
                setStyleProp("sz", parseFloat(e.target.value));
            }
        });

        $("bgColorPicker").addEventListener("input", (e) => {
            S.fillColor = e.target.value;
            setStyleProp("fill", S.fillColor);
        });
        $("fgColorPicker").addEventListener("input", (e) => {
            S.fontColor = e.target.value;
            setStyleProp("color", S.fontColor);
        });

        $("sheetTabs").addEventListener("click", (e) => {
            const t = e.target.closest("[data-sheet]");
            if (t) gotoSheet(t.dataset.sheet);
        });
        $("sheetsPopup").addEventListener("click", (e) => {
            const t = e.target.closest("[data-pop-sheet]");
            if (t) gotoSheet(t.dataset.popSheet);
        });

        $("fileInput").addEventListener("change", (e) => {
            importFile(e.target.files[0]);
            e.target.value = "";
        });
        window.addEventListener("dragover", (e) => e.preventDefault());
        window.addEventListener("drop", (e) => {
            e.preventDefault();
            if (e.dataTransfer.files && e.dataTransfer.files[0]) importFile(e.dataTransfer.files[0]);
        });
        $("settingsModal").addEventListener("mousedown", (e) => {
            if (e.target.id === "settingsModal") closeSettings();
        });

        const nb = $("nameBox");
        nb.addEventListener("keydown", (e) => {
            if (e.key === "Enter") {
                e.preventDefault();
                jumpToNameBox(nb.value);
                e.stopPropagation();
                nb.blur();
            }
        });

        const fi = $("formulaInput");
        fi.addEventListener("input", () => {
            const d = curSheet.cells[key(S.anchor.r, S.anchor.c)] || {};
            if (d.editableAction) {
                S.custom = fi.value;
                const td = cellEls[S.anchor.r] && cellEls[S.anchor.r][S.anchor.c];
                if (td) {
                    td.innerHTML = esc(S.custom || "✎ Свой ход: напишите действие в строке формул (B3) и нажмите Enter");
                    td.classList.toggle("editing", Boolean(S.custom));
                }
            }
        });
        fi.addEventListener("keydown", (e) => {
            if (e.key === "Enter") { e.preventDefault(); e.stopPropagation(); commitFormula(); }
            else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); if (!e.repeat) cancelEdit(); }
        });

        document.addEventListener("keydown", (e) => {
            if ($("subscriptionDialog").open) return;
            if (!$("welcome").hidden) return;
            if (e.key === "Escape") {
                e.preventDefault();
                if (e.repeat) return;
                if (!$("settingsModal").hidden) closeSettings();
                else toggleBoss();
                return;
            }
            if (!$("settingsModal").hidden || e.ctrlKey || e.metaKey || e.altKey) return;
            if (mask !== "excel") {
                if (!S.bossMode && !e.target.closest("input,textarea,select,[contenteditable],button") && ["1","2","3"].includes(e.key)) choose(+e.key-1);
                return;
            }
            const tag = document.activeElement && document.activeElement.tagName;
            if (tag === "INPUT" || tag === "SELECT" || tag === "TEXTAREA" || tag === "BUTTON" || document.activeElement.isContentEditable) return;
            const { r, c } = S.cur;
            const ext = e.shiftKey;
            switch (e.key) {
                case "ArrowDown": e.preventDefault(); select(r + 1, c, ext); return;
                case "ArrowUp": e.preventDefault(); select(r - 1, c, ext); return;
                case "ArrowLeft": e.preventDefault(); select(r, c - 1, ext); return;
                case "ArrowRight": e.preventDefault(); select(r, c + 1, ext); return;
                case "Tab": e.preventDefault(); select(S.anchor.r, S.anchor.c + (ext ? -1 : 1)); return;
                case "F2": { const f = $("formulaInput"); if (!f.readOnly) f.focus(); return; }
                case "Enter": {
                    const a = S.anchor;
                    if (S.sheet === "adv" && a.r >= 4 && a.r <= 6) { choose(a.r - 4); return; }
                    if (S.sheet === "adv" && a.r === 3 && a.c === 2) { $("formulaInput").focus(); return; }
                    select(a.r + 1, a.c);
                    return;
                }
                case "Delete":
                case "Backspace": {
                    if (isLocked(S.anchor.r, S.anchor.c)) return;
                    const ed = S.edits[S.sheet] = S.edits[S.sheet] || {};
                    forRange((rr, cc) => { if (!isLocked(rr, cc)) ed[key(rr, cc)] = ""; });
                    render();
                    return;
                }
                default: break;
            }
            if (S.sheet === "adv" && !S.over && ["1", "2", "3"].includes(e.key)) { choose(+e.key - 1); return; }
            if (S.over && S.sheet === "adv" && ["1", "2", "3"].includes(e.key)) { choose(+e.key - 1); return; }
            if (e.key.length === 1 && !isLocked(S.anchor.r, S.anchor.c)) {
                const f = $("formulaInput");
                f.readOnly = false;
                f.focus();
                f.value = e.key;
                f.dispatchEvent(new Event("input"));
                e.preventDefault();
            }
        });
    }

    // One RPG state, four presentation configurations. Work DOM is mounted once.
    const MASKS = {
        excel: {name:"Excel", color:"#19734b", icon:"X", guide:"excel", file:"Q4_Прогноз.xlsx", description:["Отчёт, ячейки и формулы","Report, cells and formulas"]},
        word: {name:"Word", color:"#285b9f", icon:"W", guide:"blue", file:"Project review.docx", description:["Документ и заметки","Document and notes"]},
        outlook: {name:"Outlook", color:"#087da8", icon:"O", guide:"green", file:"Inbox", description:["Письма и область чтения","Inbox and reading pane"]},
        autocad: {name:"AutoCAD", color:"#a33437", icon:"A", guide:"green", file:"Floor plan.dwg", description:["Чертёж и слои","Drawing and layers"]}
    };
    let mask = store.get("te_mask","excel");
    if (!MASKS[mask]) mask = "excel";
    let language = store.get("te_language","ru");
    let started = false;
    let tourStep = -1;
    let tourVisitedWork = false;
    let workspaceActionPending = false;
    let tourStartTurn = 1;
    const t = (ru,en) => language === "en" ? en : ru;
    const button = (label, fn, parent) => {
        const b = document.createElement("button");
        b.type="button"; b.textContent=label; b.onclick=fn; parent.append(b); return b;
    };
    function initWorkspace() {
        const original = document.createElement("div");
        original.id="excelShell";
        const modal=$("settingsModal");
        for (const child of Array.from(document.body.children)) {
            if (child.tagName !== "SCRIPT" && child !== modal) original.append(child);
        }
        document.body.prepend(original);
        document.body.insertAdjacentHTML("afterbegin", '<div id="compactTools" hidden></div><dialog id="subscriptionDialog" aria-label="TabEscape Plus"></dialog>');
        $("subscriptionDialog").addEventListener("cancel", event => {event.preventDefault(); closeOffer();});
        document.body.insertAdjacentHTML("afterbegin", '<main id="welcome"></main><nav id="workspaceTools" hidden></nav><main id="alternateGame" hidden></main><main id="alternateWork" hidden></main><aside id="guide" hidden aria-label="Обучение / Tutorial"></aside>');
        showWelcome();
    }
    function showWelcome() {
        if (S.bossMode) setBoss(false);
        tourStep=-1;
        $("welcome").hidden=false;
        $("welcome").innerHTML='<header><b>TabEscape</b><span>B2C · ENTERTAINMENT</span></header><section><p class="welcome-kicker">YOUR NEXT CHAPTER, BETWEEN TASKS</p><h1>'+t("Фэнтези-приключение<br>на короткий перерыв","A fantasy adventure.<br>A little break.")+'</h1><p>'+t("Выберите рабочий интерфейс. Нажмите Esc, чтобы переключиться между ним и игрой","Choose a work interface. Press Esc to switch between it and your adventure.")+'</p><div class="maskCards"></div><p class="sessionNote">'+t("Без аккаунта и ключа ИИ. Партия продолжается в этой вкладке до перезагрузки.","No account or AI key. Your adventure continues in this tab until you reload.")+'</p></section>';
        button(language==="en"?"Русский":"English",()=>{language=language==="en"?"ru":"en";store.set("te_language",language);showWelcome();},$("welcome").querySelector("header"));
        for (const [id,config] of Object.entries(MASKS)) {
            const b=button("",()=>enterMask(id),$("welcome").querySelector(".maskCards"));
            b.className="maskCard "+id; b.style.setProperty("--accent",config.color);
            b.innerHTML='<div class="preview"><div class="previewBar">'+config.icon+' &nbsp; '+config.file+'</div><div class="previewBody"><i></i><i></i><i></i><i></i></div></div><h2>'+config.name+'</h2><p>'+config.description[language==="en"?1:0]+'</p>';
        }
        const quota = document.createElement("p");
        quota.className = "quotaNote";
        quota.textContent = t("Демо · истории без ограничений", "Demo · unlimited stories");
        $("welcome").querySelector("section").append(quota);
        syncWorkspace();
        $("welcome").scrollTop=0;
    }
    async function enterMask(id) {
        mask=id; store.set("te_mask",mask); $("welcome").hidden=true;
        if (!started) {
            started=true;
            try {
                const response = await fetch("/api/workbook/default");
                if (!response.ok) throw new Error("No default workbook");
                const buffer = await response.arrayBuffer();
                const sheets = await parseXlsx(buffer);
                if (sheets.length) {
                    S.imported = {name:"Q4_Прогноз.xlsx",fromScratch:false,sheets};
                    S.sourceB64 = arrayBufferToBase64(buffer);
                }
            } catch { /* The built-in Q4 forecast remains available. */ }
            await restart();
        }
        if (!S.bossMode) render();
        if (store.get("te_tutorial","")!=="done") startTour();
        else syncWorkspace();
    }
    function startTour() {
        if (S.bossMode) setBoss(false);
        S.sheet="adv"; tourStep=0; tourStartTurn=S.turn; tourVisitedWork=false; render();
    }
    function finishTour() {tourStep=-1; store.set("te_tutorial","done");syncWorkspace();}
    function syncWorkspace() {
        if (!$("welcome")) return;
        document.body.append($("guide"));
        const home=!$("welcome").hidden;
        document.body.dataset.mask=mask;
        document.body.classList.toggle("work-mode",S.bossMode);
        $("excelShell").hidden=home || mask!=="excel";
        $("alternateGame").hidden=home || mask==="excel" || S.bossMode;
        $("alternateWork").hidden=home || mask==="excel" || !S.bossMode;
        $("workspaceTools").hidden=home;
        $("workspaceTools").replaceChildren();
        const cfg=MASKS[mask];
        const brand=document.createElement("strong"); brand.textContent=cfg.name; $("workspaceTools").append(brand);
        if (S.bossMode && mask==="excel") {
            button(t("Открыть .xlsx / .csv","Open .xlsx / .csv"),pickFile,$("workspaceTools"));
            button(t("Новая книга","New workbook"),createNewExcelWorkbook,$("workspaceTools"));
            button(t("Экспорт .xlsx","Export .xlsx"),exportXlsx,$("workspaceTools"));
        }
        if (!S.bossMode) {
            button(t("Маски","Interfaces"),showWelcome,$("workspaceTools"));
            button(t("Как играть","How to play"),startTour,$("workspaceTools"));
            button(t("Параметры героя","Hero settings"),openSettings,$("workspaceTools"));
            button(t("Новая игра","Restart"),()=>{if(!S.loading)restart();},$("workspaceTools"));
            button(t("Дневник ↓","Diary ↓"),exportReport,$("workspaceTools"));
        }
        button(S.bossMode?t("Переключить · Esc","Switch · Esc"):t("Рабочий экран · Esc","Work screen · Esc"),toggleBoss,$("workspaceTools")).id="switchView";
        const hide = button(t("Скрыть панель", "Hide toolbar"), () => {
            panelHidden = true;
            store.set("te_panel_hidden", "true");
            updatePanelVisibility();
            $("compactTools").querySelector("button").focus();
        }, $("workspaceTools"));
        hide.id = "hideToolbar";
        updatePanelVisibility();
        if (!home && mask==="excel" && !S.bossMode && language==="en" && S.sheet==="adv") {
            for (const [r,label] of [[2,"Health"],[3,"Energy"],[4,"Trouble"],[5,"Time"],[6,"Inventory"]]) {
                const cell=document.querySelector('td[data-r="'+r+'"][data-c="4"]');
                if(cell) cell.textContent=label;
            }
        }
        if (!home && mask!=="excel") {
            if (!S.bossMode) renderAlternateGame();
            else mountWork();
        }
        document.title=home?"TabEscape":cfg.file+" — "+cfg.name;
        document.querySelector('link[rel="icon"]').href="data:image/svg+xml,"+encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32"><rect width="32" height="32" rx="5" fill="'+cfg.color+'"/><text x="16" y="23" text-anchor="middle" fill="white" font-size="23">'+cfg.icon+'</text></svg>');
        renderTour();
        if (DailyAllowance.remaining(store) > 0) offerPending = false;
        if (offerPending && !offerDismissed && !S.bossMode && !home) offerSubscription();
    }
    function renderAlternateGame() {
        const root=$("alternateGame");
        root.className="alternative "+mask;
        root.innerHTML='<aside class="folders"><b>'+MASKS[mask].file+'</b><p>'+t("Оукхейвен","Oakhaven")+'</p><p>'+t("Архив приключения","Adventure archive")+'</p></aside><section class="gameDocument"><div id="altScene"><small>'+t("Текущая сцена","Current scene")+'</small><h1>Oakhaven</h1><p>'+esc(S.loading||S.narrative)+'</p></div><div id="altChoices"></div><form id="customForm"><label for="customAction">'+t("Свой ход","Your action")+'</label><input id="customAction" maxlength="500"><button>'+t("Отправить","Send")+'</button></form></section><aside id="altStats"></aside>';
        const stats=$("altStats");
        stats.innerHTML='<h2>'+clockStr()+'</h2><p>'+t("Здоровье","Health")+' '+S.stats.ROI+'%</p><p>'+t("Бодрость","Energy")+' '+S.stats.KPI+'%</p><p>'+t("Неприятности","Trouble")+' '+S.stats.Anxiety+'%</p><p>'+t("Рекорд","Best")+' '+S.best+'</p><h3>'+t("Инвентарь","Inventory")+'</h3>';
        S.inventory.forEach(name=>button(name,()=>useItem(name),stats));
        const details=document.createElement("details");
        details.innerHTML='<summary>'+t("Журнал","Journal")+' ('+S.log.length+')</summary>'+S.log.map(e=>'<p>'+e.time+' · '+esc(e.action)+'</p>').join("");
        stats.append(details);
        if (S.over) button(t("Сыграть снова","Play again"),restart,$("altChoices"));
        else S.options.forEach((option,i)=>{const b=button((i+1)+". "+option,()=>{S.sheet="adv";choose(i);},$("altChoices"));b.disabled=!!S.loading;});
        $("customForm").onsubmit=e=>{e.preventDefault();const value=$("customAction").value.trim();if(value)sendAction(value,-1);};
        $("customForm").hidden=S.over;
        $("customAction").value=S.custom;
        $("customAction").oninput=e=>S.custom=e.target.value;
        $("customAction").disabled=!!S.loading;
        if(mask==="outlook") {
            const folder=root.querySelector(".folders");
            folder.innerHTML="<b>Inbox</b><p>Guild (1)</p><p>Sent</p><p>Archive</p>";
            const list=document.createElement("aside");list.className="gameMailList";
            list.innerHTML="<b>Guild correspondence</b><p>Morgan · "+clockStr()+"</p><p>Oakhaven — current assignment</p>";
            root.insertBefore(list,root.querySelector(".gameDocument"));
        }
        if(mask==="autocad") {
            const drawing=document.createElement("div");drawing.className="adventureMap";
            drawing.innerHTML="<div>OAKHAVEN</div><div>GUILD</div><div>RIVER</div>";
            root.querySelector(".gameDocument").prepend(drawing);
        }
        const chart=document.createElement("div");chart.className="historyBars";
        chart.setAttribute("aria-label",t("Динамика здоровья по ходам","Health by turn"));
        for(const entry of S.history) {
            const bar=document.createElement("i");bar.style.height=Math.max(2,entry.ROI)+"%";
            bar.title=String(entry.ROI)+"%";chart.append(bar);
        }
        stats.append(chart);
        if (mask === "word") wordChrome(root, false);
    }
    const mail = [
        ["Alex Morgan","Q4 planning review","Hi team, please review the forecast before Thursday. Focus on operating expenses and the updated regional totals. Thanks, Alex."],
        ["Sam Lee","Office floor plan","The revised layout includes a meeting room and two collaboration areas. Please check the furniture layer before our review."],
        ["Taylor Reed","Weekly project update","Milestone one is complete. The next review is scheduled for Friday at 10:00. Please add your notes to the project document."]
    ];
    function mountWork() {
        const root=$("alternateWork");
        for(const child of root.children) child.hidden=child.dataset.work!==mask;
        let pane=root.querySelector('[data-work="'+mask+'"]');
        if(pane) return;
        pane=document.createElement("section");pane.dataset.work=mask;pane.className="workPane "+mask;root.append(pane);
        if(mask==="word") pane.innerHTML='<article contenteditable="true" aria-label="Project document" role="textbox" aria-multiline="true"><p>OPERATIONS / NOVEMBER 2026</p><h1>Quarterly project review</h1><h2>Executive summary</h2><p>The team completed the first delivery milestone. This document records progress, open questions and next steps for the quarterly review.</p><h2>Priorities</h2><p>Review the forecast, confirm ownership of outstanding tasks and prepare the revised timeline.</p><h2>Next steps</h2><p>Gather feedback by Thursday. Share the updated document with the project team on Friday.</p></article>';
        if (mask === "word") wordChrome(pane, true);
        if(mask==="outlook") {
            pane.innerHTML='<aside><h2>Folders</h2><p>Inbox (3)</p><p>Sent</p><p>Drafts</p></aside><div class="mailList"></div><article class="reading" tabindex="0"></article>';
            const read=i=>{pane.querySelector(".reading").innerHTML='<small>'+esc(mail[i][0])+' · 10:15 AM</small><h1>'+esc(mail[i][1])+'</h1><p>'+esc(mail[i][2])+'</p>';};
            mail.forEach((m,i)=>button(m[0]+" — "+m[1],()=>read(i),pane.querySelector(".mailList")));
            read(0);
        }
        if(mask==="autocad") {
            pane.innerHTML='<aside><h2>Layers</h2><label><input type="checkbox" id="furniture" checked> Furniture</label><label>Zoom <input id="cadZoom" type="range" min="50" max="180" value="100"></label></aside><div class="drawing"><div class="floorPlan"><div class="room">MEETING ROOM</div><div class="room">OFFICE 01</div><div class="room">OFFICE 02</div><div class="room">LOBBY</div><div class="furniture">▤ &nbsp; ▤ &nbsp; ▤</div></div><small>FLOOR PLAN · SCALE 1:100 · A-101</small></div>';
            pane.querySelector("#cadZoom").oninput=e=>pane.querySelector(".floorPlan").style.transform="scale("+e.target.value/100+")";
            pane.querySelector("#furniture").onchange=e=>pane.querySelector(".furniture").hidden=!e.target.checked;
        }
    }
    function renderTour() {
        document.querySelectorAll(".tourFocus").forEach(el=>el.classList.remove("tourFocus"));
        const guide=$("guide");guide.hidden=tourStep<0 || S.bossMode || !$("welcome").hidden;
        if(guide.hidden) return;
        const excel=mask==="excel";
        const selectors=excel?['td[data-r="2"][data-c="2"]','td[data-r="2"][data-c="4"],td[data-r="2"][data-c="5"],td[data-r="2"][data-c="6"],td[data-r="3"][data-c="4"],td[data-r="3"][data-c="5"],td[data-r="3"][data-c="6"],td[data-r="4"][data-c="4"],td[data-r="4"][data-c="5"],td[data-r="4"][data-c="6"]','tr[data-opt]','#switchView']:["#altScene","#altStats","#altChoices","#switchView"];
        const targetSelector = tourStep === 3 && panelHidden ? "#compactTools" : selectors[tourStep];
        const targets=document.querySelectorAll(targetSelector);
        targets.forEach(el=>el.classList.add("tourFocus"));
        if (window.matchMedia("(max-width:800px)").matches && !excel && targets.length) {
            targets[targets.length-1].insertAdjacentElement("afterend",guide);
        }
        const texts=[
            t("Здесь появляется текущая сцена. Твоя цель — добраться до 18:00. Каждый ход занимает 30 минут.","This is the current scene. Reach 18:00. Each action takes 30 minutes."),
            t("Здоровье и бодрость начинаются со 100%, неприятности — с 10%. При нуле здоровья или бодрости либо при 100% неприятностей приключение заканчивается.","Health and energy start at 100%, trouble at 10%. Zero health or energy, or 100% trouble, ends the adventure.")+" "+t("Сейчас","Now")+": "+S.stats.ROI+" / "+S.stats.KPI+" / "+S.stats.Anxiety+"%.",
            t("Нажми на один из трёх вариантов или используй 1, 2, 3. ","Choose an option or press 1, 2, 3. ")+(excel?t("В Excel можно выделить вариант и нажать Enter. Для своего действия выбери B3, введи текст в строке формул и нажми Enter.","In Excel select an option and press Enter. For a custom action select B3, type in the formula bar and press Enter."):t("Для своего действия используй поле «Свой ход» и кнопку «Отправить».","For a custom action use Your action and Send."))+t(" Попробуй настоящий ход."," Try a real turn."),
            t("Нажми Esc, чтобы открыть рабочий экран. Нажми ещё раз, чтобы продолжить с того же места.","Press Esc to open the work screen. Press it again to continue from the same scene.")
        ];
        guide.innerHTML='<img src="/static/guides/'+MASKS[mask].guide+'.jpeg" alt=""><div><small>'+t("Проводница","Your guide")+' · '+(tourStep+1)+' / 4</small><p>'+texts[tourStep]+'</p><div class="tourButtons"></div></div>';
        const actions=guide.querySelector(".tourButtons");
        const back=button(t("Назад","Back"),()=>{tourStep--;renderTour();},actions);back.disabled=tourStep===0;
        button(t("Пропустить","Skip"),finishTour,actions);
        const next=button(tourStep===3?t("К приключению","Start adventuring"):t("Далее","Next"),()=>{if(tourStep===3)finishTour();else{tourStep++;renderTour();}},actions);
        next.disabled=(tourStep===2 && S.turn===tourStartTurn)||(tourStep===3 && !tourVisitedWork);
    }

    let panelHidden = store.get("te_panel_hidden", "false") === "true";
    let adventureId = null;
    let offerPending = false;
    let offerDismissed = false;
    let offerFocus = null;

    function updatePanelVisibility() {
        const home = !$("welcome").hidden;
        $("workspaceTools").hidden = home || panelHidden;
        $("compactTools").hidden = home || !panelHidden;
        $("compactTools").replaceChildren();
        button(t("Показать панель", "Show toolbar"), () => {
            panelHidden = false;
            store.set("te_panel_hidden", "false");
            updatePanelVisibility();
            $("hideToolbar")?.focus();
        }, $("compactTools"));
        button(t("Переключить · Esc", "Switch · Esc"), toggleBoss, $("compactTools"));
    }

    function offerSubscription() {
        offerPending = true;
        offerDismissed = false;
        if (S.bossMode) return;
        const dialog = $("subscriptionDialog");
        if (dialog.open) return;
        offerFocus = document.activeElement;
        dialog.innerHTML = '<div class="offerLayout"><div class="offerPortrait"><img src="/static/guides/' +
            MASKS[mask].guide + '.jpeg" alt=""></div><section><span class="offerEyebrow">TABESCAPE PLUS</span><h2>' +
            t("На сегодня — целое приключение", "You’ve had quite an adventure") + '</h2><p>' +
            t("Ты завершил три истории! В демо можно продолжать играть бесплатно и без ограничений.",
              "You’ve finished three stories! Keep playing this demo for free, with no limits.") +
            '</p><div class="plusFeatures"><h3>' + t("Больше миров с Plus", "More worlds with Plus") +
            '</h3><ul><li>' + t("Новые сюжеты с ИИ", "New AI-generated stories") + '</li><li>' +
            t("Выбор жанров и атмосферы", "Choose genres and atmosphere") + '</li><li>' +
            t("Бесконечный режим", "Endless mode") + '</li></ul><p>' +
            t("Подписка готовится к запуску. Цена и оплата появятся позже.",
              "Subscription coming soon. Pricing and checkout will be available at launch.") +
            '</p></div><div class="offerActions"></div><p class="offerNote" role="status"></p></section></div>';
        const actions = dialog.querySelector(".offerActions");
        button(t("Продолжить бесплатно", "Continue for free"), () => {
            closeOffer();
            if (S.over) restart();
        }, actions).className = "primary";
        button(t("Купить подписку · скоро", "Get Plus · coming soon"), () => {
            dialog.querySelector(".offerNote").textContent = t(
                "Plus ещё в разработке. Генерация историй и оплата станут доступны после запуска подписки.",
                "Plus is in development. Story generation and checkout will become available when subscriptions launch.");
        }, actions);
        button(t("Прийти завтра", "Come back tomorrow"), () => {
            closeOffer();
            showWelcome();
        }, actions);
        button(started ? t("Посмотреть финал", "Review ending") : t("Закрыть", "Close"), closeOffer, actions);
        dialog.showModal();
    }

    function closeOffer() {
        $("subscriptionDialog").close();
        offerDismissed = true;
        if (offerFocus?.isConnected) offerFocus.focus();
    }

    function wordChrome(root, editing) {
        const chrome = document.createElement("header");
        chrome.className = "wordChrome";
        chrome.innerHTML = '<div class="wordTitle"><div class="wordQuick"></div><strong>' +
            (editing ? t("Документ1", "Document1") : t("Оукхейвен — дневник", "Oakhaven — journal")) +
            '</strong><div class="wordWindowActions"></div></div><nav class="wordTabs" aria-label="Word"></nav>' +
            '<div class="wordCommands"></div><div class="wordRuler" aria-hidden="true"><span>' +
            Array.from({length:19}, (_, i) => '<i>' + i + '</i>').join("") +
            '</span></div>';
        root.prepend(chrome);
        const editor = editing ? root.querySelector("[contenteditable]") : root.querySelector(".gameDocument");
        const quick = chrome.querySelector(".wordQuick");
        const home = button("⌂", showWelcome, quick);
        home.title = t("Выбрать интерфейс", "Choose interface");
        home.setAttribute("aria-label", home.title);

        function command(label, name, value, parent, title) {
            const control = button(label, () => {
                editor.focus();
                document.execCommand(name, false, value);
            }, parent);
            control.disabled = !editing;
            control.title = title || label;
            control.setAttribute("aria-label", title || label);
            control.onmousedown = event => event.preventDefault();
            return control;
        }

        command("↶", "undo", null, quick, t("Отменить", "Undo"));
        command("↷", "redo", null, quick, t("Повторить", "Redo"));
        button(t("Переключить · Esc", "Switch · Esc"), toggleBoss,
            chrome.querySelector(".wordWindowActions"));

        function group(label) {
            const section = document.createElement("section");
            section.className = "wordToolGroup";
            section.setAttribute("aria-label", label);
            section.innerHTML = '<div class="wordToolItems"></div><small>' + label + '</small>';
            chrome.querySelector(".wordCommands").append(section);
            return section.firstElementChild;
        }

        function selectControl(label, options, current, change, parent) {
            const select = document.createElement("select");
            select.setAttribute("aria-label", label);
            for (const value of options) {
                const option = document.createElement("option");
                option.value = value;
                option.textContent = value;
                option.selected = String(value) === String(current);
                select.append(option);
            }
            select.disabled = !editing;
            select.onchange = () => change(select.value);
            parent.append(select);
        }

        function renderRibbon(tab) {
            chrome.querySelector(".wordCommands").replaceChildren();
            for (const item of chrome.querySelectorAll(".wordTabs button")) {
                item.setAttribute("aria-pressed", String(item.dataset.tab === tab));
            }
            if (tab === "home") {
                const font = group(t("Шрифт", "Font"));
                selectControl(t("Шрифт", "Font"), ["Calibri", "Arial", "Georgia", "Times New Roman"],
                    editor.style.fontFamily.replaceAll('"', "") || "Calibri",
                    value => editor.style.fontFamily = value, font);
                selectControl(t("Размер текста", "Text size"), [10, 11, 12, 14, 16, 18, 20, 24],
                    parseInt(editor.style.fontSize) || 16,
                    value => editor.style.fontSize = value + "px", font);
                command(t("Ж", "B"), "bold", null, font, t("Полужирный", "Bold"));
                command(t("К", "I"), "italic", null, font, t("Курсив", "Italic"));
                command(t("Ч", "U"), "underline", null, font, t("Подчёркнутый", "Underline"));
                command("ab", "strikeThrough", null, font, t("Зачёркнутый", "Strikethrough"));
                const paragraph = group(t("Абзац", "Paragraph"));
                command("☷", "insertUnorderedList", null, paragraph, t("Маркированный список", "Bulleted list"));
                command("1.", "insertOrderedList", null, paragraph, t("Нумерованный список", "Numbered list"));
                command("≡", "justifyLeft", null, paragraph, t("По левому краю", "Align left"));
                command("≡", "justifyCenter", null, paragraph, t("По центру", "Center"));
                command("≡", "justifyRight", null, paragraph, t("По правому краю", "Align right"));
                const styles = group(t("Стили", "Styles"));
                styles.classList.add("wordStyleGallery");
                for (const [label, tag] of [[t("Обычный","Normal"),"p"],[t("Заголовок 1","Heading 1"),"h1"],[t("Заголовок 2","Heading 2"),"h2"],[t("Подзаголовок","Subtitle"),"h3"]]) {
                    const tile = command(label, "formatBlock", tag, styles);
                    tile.className = "wordStyleTile";
                    const sample = document.createElement("span");
                    sample.textContent = "AaБбВв";
                    sample.className = "styleSample " + tag;
                    tile.prepend(sample);
                }
            }
            if (tab === "insert") {
                const insert = group(t("Элементы документа", "Document elements"));
                command(t("Горизонтальная линия", "Horizontal line"), "insertHorizontalRule", null, insert);
                command(t("Список", "List"), "insertUnorderedList", null, insert);
                command(t("Дата", "Date"), "insertText", new Date().toLocaleDateString(), insert);
            }
            if (tab === "layout") {
                const layout = group(t("Поля страницы", "Page margins"));
                for (const [label, padding] of [[t("Обычные","Normal"),"64px"],[t("Узкие","Narrow"),"32px"]]) {
                    button(label, () => editor.style.padding = padding, layout);
                }
            }
            if (tab === "view") {
                const view = group(t("Отображение", "Display"));
                button(t("Линейка", "Ruler"), () => {
                    chrome.querySelector(".wordRuler").hidden = !chrome.querySelector(".wordRuler").hidden;
                }, view);
                button(t("Масштаб +", "Zoom in"), () => editor.style.zoom =
                    Math.min(1.5, (parseFloat(editor.style.zoom) || 1) + .1), view);
                button(t("100%", "100%"), () => editor.style.zoom = 1, view);
                if (!editing) button(t("Показатели", "Stats"), () => $("altStats").scrollIntoView({block:"center"}), view);
            }
        }

        for (const [id, label] of [["home",t("Главная","Home")],["insert",t("Вставка","Insert")],
            ["layout",t("Макет","Layout")],["view",t("Вид","View")]]) {
            const control = button(label, () => renderRibbon(id), chrome.querySelector(".wordTabs"));
            control.dataset.tab = id;
        }
        renderRibbon("home");
        const footer = document.createElement("footer");
        footer.className = "wordStatus";
        footer.textContent = editing ? t("Страница 1 из 1 · Редактирование", "Page 1 of 1 · Editing") :
            t("Оукхейвен · ", "Oakhaven · ") + clockStr();
        root.append(footer);
    }

    // ------------------------------------------------------------------ публичный API
    window.App = {
        openSettings, closeSettings, saveSettings, toggleBoss, commitFormula,
        cancelEdit, insertFn, zoomBy, resetZoom, toggleFullscreen, toggleRibbon,
        scrollTabs, toggleSheetsMenu, addSheet, showWorkbookStats, exportXlsx,
        createNewExcelWorkbook, renameWorkbook, toggleMergeCells, gotoSheet, toast,
    };

    document.addEventListener("DOMContentLoaded", async () => {
        document.documentElement.dataset.theme = S.theme;
        initWorkspace();
        bind();
        render();
    });
})();
