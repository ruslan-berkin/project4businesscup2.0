/* Anchors are resolved again after the game renders or changes workspace. */
class WorkspaceTour {
    constructor(options) {
        Object.assign(this, options);
        this.step = -1;
        this.frame = 0;
        this.lastKey = "";
        this.shade = document.createElement("div");
        this.shade.className = "spotlight";
        this.shade.hidden = true;
        this.shade.setAttribute("aria-hidden", "true");
        document.body.append(this.shade);
        this.guide.setAttribute("role", "dialog");
        this.guide.setAttribute("aria-labelledby", "tourTitle");
        this.guide.setAttribute("aria-describedby", "tourText");
        this.guide.innerHTML = `<img class="guideSprite" src="${portableImages["mira.webp"]}" alt="">
            <div class="speech"><div class="speaker"><strong>Mira</strong><span id="tourCount"></span></div>
            <div aria-live="polite" aria-atomic="true"><h2 id="tourTitle"></h2><p id="tourText"></p></div>
            <div class="tourButtons"><button data-go="skip"></button><button data-go="back"></button><button data-go="next"></button></div></div>`;
        this.control = name => this.guide.querySelector('[data-go="' + name + '"]');
        this.control("skip").onclick = () => this.close();
        this.control("back").onclick = () => this.move(-1);
        this.control("next").onclick = () => this.move(1);
        this.observer = new ResizeObserver(() => this.schedule());
        document.addEventListener("scroll", () => this.schedule(), true);
        window.addEventListener("resize", () => this.refresh());
        window.visualViewport?.addEventListener("resize", () => this.refresh());
        document.addEventListener("keydown", event => this.keydown(event), true);
        for (const type of ["pointerdown", "mousedown", "click", "focusin"]) {
            document.addEventListener(type, event => {
                if (!this.active || this.allowed(event.target)) return;
                event.preventDefault();
                event.stopImmediatePropagation();
                if (type === "focusin") this.control("next").focus({preventScroll:true});
            }, true);
        }
    }
    visible(el) {
        return el && el.isConnected && el.getClientRects().length > 0 && getComputedStyle(el).visibility !== "hidden";
    }
    get active() { return this.step >= 0; }
    start(home = false) {
        this.returnFocus = document.activeElement;
        this.fromHome = home;
        this.step = home ? 0 : 1;
        this.lastKey = "";
        this.refresh();
        this.control("next").focus({preventScroll:true});
    }
    move(direction) {
        if (this.entering) return;
        if (direction > 0 && this.step === 5) { this.close(); return; }
        if (this.step === 0 && direction > 0) { this.enter(this.selected()); return; }
        this.step += direction;
        if (!this.step) this.home();
        if (this.step >= 1) this.setWork(this.step === 5);
        this.refresh();
        this.control("next").focus({preventScroll:true});
    }
    async enter(id) {
        if (this.entering) return;
        this.entering = true;
        this.step = 1;
        try {
            const entered = await this.openMask(id);
            if (entered === false) this.step = this.fromHome ? 0 : 1;
        }
        finally {
            this.entering = false;
            if (this.active) {
                this.refresh();
                this.control("next").focus({preventScroll:true});
            }
        }
    }
    close() {
        this.step = -1;
        this.save();
        this.observer.disconnect();
        cancelAnimationFrame(this.frame);
        this.frame = 0;
        this.lastKey = "";
        this.guide.hidden = this.shade.hidden = true;
        this.targets = [];
        document.body.classList.remove("tourActive", "tourDock");
        document.body.style.removeProperty("--dock-height");
        if (this.isWork()) this.setWork(false);
        const fallback = [...document.querySelectorAll(".tutorialLaunch, [data-home-help]")].find(el => this.visible(el));
        const focus = this.returnFocus?.isConnected && this.visible(this.returnFocus) && this.returnFocus !== document.body ? this.returnFocus : fallback;
        focus?.focus({preventScroll:true});
        this.afterClose();
    }
    allowed(el) {
        return document.querySelector("#gameLobby[open], #upcomingDialog[open]") || this.isWork() || this.guide.contains(el) || this.targets?.some(target => target.contains(el));
    }
    keydown(event) {
        if (!this.active || document.querySelector("#gameLobby[open], #upcomingDialog[open]")) return;
        if (this.isWork() && event.key !== "Escape") return;
        if (event.key === "Tab") {
            const clickable = (this.targets || []).flatMap(el => el.matches("button:enabled") ? [el] : [...el.querySelectorAll("button:enabled")]);
            const controls = [...clickable, ...this.guide.querySelectorAll("button:enabled")].filter(el => this.visible(el));
            const delta = event.shiftKey ? -1 : 1;
            const next = (controls.indexOf(document.activeElement) + delta + controls.length) % controls.length;
            event.preventDefault();
            controls[next]?.focus({preventScroll:true});
        } else if (event.key === "Escape") {
            event.preventDefault();
            event.stopImmediatePropagation();
            if (event.repeat) return;
            if (this.step >= 4) this.move(1);
            else this.close();
        } else if (this.step === 2 && !this.ended() && /^[123]$/.test(event.key) && !event.ctrlKey && !event.altKey && !event.metaKey) {
            event.preventDefault();
            if (!event.repeat) this.choose(Number(event.key) - 1);
        }
    }
    describe() {
        const t = this.t;
        const grid = this.isGrid();
        const activeSwitch = "#switchView, .compactSwitch, .wordSwitch";
        const end = this.ended();
        const steps = [
            [".maskCards", t("Ваше рабочее окно","Pick your workspace"), t("Я Мира. Выберите оформление или нажмите «Далее», чтобы открыть выбранное окно. Esc закрывает обучение.","I’m Mira. Pick a workspace, or use Next to open " + this.selectedName() + ". Esc skips the tour.")],
            [grid ? "td.story" : "#altScene", t("Ваша история","Read the scene"), this.goalText()],
            [grid ? "td.optcell" : "#altChoices", end ? t("Финал","Adventure complete") : t("Ваш выбор","Choose your next move"), end ? t("Здесь можно начать новую игру.","This is where you can start another adventure.") : t("Попробуйте вариант или нажмите 1–3. Настоящий ход занимает 30 игровых минут. Можно просто продолжить обучение.","Try an option or press 1–3. A real turn takes 30 game minutes. You can also select Next to keep looking around.")],
            [grid ? 'td[data-r="2"][data-c="4"],td[data-r="4"][data-c="5"]' : "#altStats .vitals", t("Показатели героя","Check your vitals"), this.endless() ? t("Игра закончится, когда здоровье упадёт до нуля. Нулевая бодрость или 100% неприятностей отнимают здоровье каждый ход.", "Only zero health ends this mode. Zero energy or 100% trouble costs health every turn.") : t("Следите за здоровьем и бодростью. Если они упадут до нуля или неприятности достигнут 100%, приключение закончится.","Keep health and energy above zero, and trouble below 100%, to stay in the adventure.")],
            [activeSwitch, t("Рабочий экран","Try the work screen"), t("Нажмите выделенную кнопку, Esc или «Далее», чтобы открыть рабочий экран. Сцена сохранится.","Use the highlighted switch, press Esc, or select Next to open the work screen. Your scene stays right here.")],
            [activeSwitch, t("Вернитесь в игру","Return to your adventure"), t("Вы на рабочем экране. Нажмите Esc, выделенную кнопку или «Завершить», чтобы продолжить с той же сцены.","You’re in the work view. Press Esc, use the switch, or select Finish to return to the same scene.")]
        ];
        return steps[this.step];
    }
    refresh() {
        if (!this.active) return;
        if (this.isWork() || document.querySelector("#gameLobby[open], #upcomingDialog[open]")) {
            this.guide.hidden = this.shade.hidden = true;
            document.body.classList.remove("tourActive", "tourDock");
            document.body.style.removeProperty("--dock-height");
            return;
        }
        let [selector, title, copy] = this.describe();
        let matches = [...document.querySelectorAll(selector)].filter(el => this.visible(el));
        // A failed initial request can leave a scene message but no available actions.
        if (this.step === 2 && (!this.hasActions() || !matches.length)) {
            selector = this.isGrid() ? "td.story" : "#altScene";
            matches = [...document.querySelectorAll(selector)].filter(el => this.visible(el));
            title = this.t("Действия", "Story choices");
            copy = this.t("Действия появятся после загрузки сцены. Если возникла ошибка, закройте обучение и нажмите Restart.", "Choices appear when the scene loads. If the server is unavailable, skip the tour and use Restart to try again.");
        }
        if (!matches.length) {
            this.targets = [];
            // Between screens there is no target. Keep Skip available while loading.
            this.shade.hidden = true;
            this.guide.hidden = false;
            this.control("next").disabled = true;
            this.guide.querySelector("#tourTitle").textContent = this.t("Открываем окно", "Opening workspace");
            this.guide.querySelector("#tourText").textContent = this.t("Можно пропустить обучение.", "You can skip the tutorial at any time.");
            return;
        }
        this.targets = this.step >= 4 ? matches.slice(0,1) : matches;
        this.guide.hidden = this.shade.hidden = false;
        document.body.classList.add("tourActive");
        const key = this.step + ":" + selector + ":" + this.fromHome;
        const changed = key !== this.lastKey;
        this.lastKey = key;
        this.guide.dataset.step = this.step;
        this.guide.querySelector("#tourCount").textContent = (this.step + (this.fromHome ? 1 : 0)) + this.t(" из ", " of ") + (this.fromHome ? 6 : 5);
        this.guide.querySelector("#tourTitle").textContent = title;
        this.guide.querySelector("#tourText").textContent = copy;
        this.control("back").textContent = this.t("Назад", "Back");
        this.control("back").disabled = this.step === (this.fromHome ? 0 : 1);
        this.control("skip").textContent = this.t("Пропустить", "Skip");
        this.control("next").textContent = this.step === 5 ? this.t("Завершить", "Finish") : this.t("Далее", "Next");
        this.control("next").disabled = this.entering === true;
        const dock = innerWidth < 700 || innerHeight < 570;
        document.body.classList.toggle("tourDock", dock);
        if (dock) document.body.style.setProperty("--dock-height", (this.guide.offsetHeight + 18) + "px");
        this.observer.disconnect();
        this.observer.observe(this.guide);
        this.targets.forEach(el => this.observer.observe(el));
        if (changed || !this.inView()) {
            const behavior = matchMedia("(prefers-reduced-motion: reduce)").matches ? "instant" : "smooth";
            this.targets[0].scrollIntoView({block:"center", inline:"nearest", behavior});
        }
        this.layout();
        this.schedule();
    }
    schedule() {
        if (!this.active || this.frame) return;
        this.frame = requestAnimationFrame(() => {this.frame = 0; this.layout();});
    }
    bounds() {
        const boxes = this.targets.map(el => el.getBoundingClientRect());
        const first = this.targets[0];
        let left = 0, top = 0, right = innerWidth, bottom = innerHeight;
        for (let p = first.parentElement; p && p !== document.body; p = p.parentElement) {
            const style = getComputedStyle(p);
            const r = p.getBoundingClientRect();
            if (/auto|scroll|hidden|clip/.test(style.overflowX)) { left = Math.max(left,r.left); right = Math.min(right,r.right); }
            if (/auto|scroll|hidden|clip/.test(style.overflowY)) { top = Math.max(top,r.top); bottom = Math.min(bottom,r.bottom); }
        }
        if (document.body.classList.contains("tourDock")) bottom = Math.min(bottom,innerHeight - this.guide.offsetHeight - 16);
        return {
            left: Math.max(left, Math.min(...boxes.map(r => r.left)) - 3),
            right: Math.min(right, Math.max(...boxes.map(r => r.right)) + 3),
            top: Math.max(top, Math.min(...boxes.map(r => r.top)) - 3),
            bottom: Math.min(bottom, Math.max(...boxes.map(r => r.bottom)) + 3)
        };
    }
    inView() {
        const r = this.bounds();
        return r.bottom-r.top > 32 && r.right-r.left > 30;
    }
    layout() {
        if (!this.active || !this.targets?.length || this.guide.hidden) return;
        const r = this.bounds(), gap = 14;
        Object.assign(this.shade.style, {left:r.left+"px",top:r.top+"px",width:Math.max(0,r.right-r.left)+"px",height:Math.max(0,r.bottom-r.top)+"px"});
        if (document.body.classList.contains("tourDock")) return;
        const w = this.guide.offsetWidth, h = this.guide.offsetHeight;
        const clip = (n,max) => Math.max(gap, Math.min(n,max-gap));
        let x = clip((r.left+r.right-w)/2, innerWidth-w);
        let y = r.bottom+gap;
        if (y+h+gap > innerHeight) {
            if (r.top > h+gap*2) y = r.top-h-gap;
            else {
                x = r.right+w+gap*2 <= innerWidth ? r.right+gap : r.left-w-gap;
                y = clip((r.top+r.bottom-h)/2, innerHeight-h);
            }
        }
        this.guide.style.left = clip(x,innerWidth-w)+"px";
        this.guide.style.top = clip(y,innerHeight-h)+"px";
    }
}
