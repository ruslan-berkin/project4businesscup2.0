/* Optional first-party playtest. No events are sent before explicit consent. */
(() => {
    const KEY = 'te_research_v1';
    let context = () => ({}), lang = () => 'en';
    const tr = (ru,en) => lang() === 'ru' ? ru : en;
    const sources = ['direct','story-community','office-community','creator','paid-search','team','jury'];
    function settings() {
        try {
            const value=JSON.parse(localStorage.getItem(KEY));
            return value?.consent === true && value.expires > Date.now() ? value : null;
        } catch { return null; }
    }
    async function event(name, extra={}) {
        const prefs=settings();
        if (!prefs) return false;
        const game=context();
        const payload={event_id:crypto.randomUUID(),visitor_id:prefs.visitor,consent:true,
            cohort:prefs.cohort,source:prefs.source,eligible:prefs.eligible,workspace:game.workspace||'excel',name,...extra};
        if (['adventure_started','turn_completed','adventure_completed','adventure_resumed'].includes(name)) {
            if (!game.story_token || !game.adventure_id) return false;
            Object.assign(payload,{story_token:game.story_token,adventure_id:game.adventure_id});
        }
        try {
            const response=await fetch('/api/research/event',{method:'POST',headers:{'Content-Type':'application/json'},
                body:JSON.stringify(payload),keepalive:true,signal:AbortSignal.timeout(5000)});
            return response.ok;
        } catch { return false; } // Telemetry must never hold up gameplay.
    }
    function open({premium=false}={}) {
        let dialog=document.getElementById('researchDialog');
        if (!dialog) {
            dialog=document.createElement('dialog');dialog.id='researchDialog';
            dialog.setAttribute('aria-labelledby','researchTitle');document.body.append(dialog);
        }
        if (dialog.open) return;
        const before=settings(), trigger=document.activeElement;
        dialog.innerHTML=`<button type="button" class="researchClose" aria-label="${tr('Закрыть','Close')}">×</button>
            <p class="eyeline">${tr('ДОБРОВОЛЬНЫЙ ПЛЕЙТЕСТ','OPTIONAL PLAYTEST')}</p>
            <h2 id="researchTitle">${tr('Поможете сделать истории лучше?','Help us make better stories?')}</h2>
            <p>${tr('Мы проверяем, хочется ли возвращаться к игре. Можно поделиться впечатлением и разрешить анонимную статистику. Игра останется доступна при любом ответе.','We’re testing whether people want to come back. Share your feedback and optionally join the anonymous playtest. Every answer leaves the game available.')}</p>
            <form>
            <label>${tr('Я здесь как','I’m here as')}<select name="cohort"><option value="pilot">${tr('Потенциальный игрок','A potential player')}</option><option value="demo">${tr('Жюри или участник команды','Jury or team member')}</option></select></label>
            <label class="researchCheck"><input type="checkbox" name="eligible">${tr('Мне 25–45, я работаю в офисной роли в США и люблю сюжетные игры.','I’m 25–45, work in an office role in the US and enjoy story games.')}</label>
            <label>${tr('Насколько было интересно?','How engaging was it?')}<select name="rating"><option value="">${tr('Пока не отвечать','Skip this question')}</option>${[1,2,3,4,5].map(n=>`<option value="${n}">${n} / 5</option>`).join('')}</select></label>
            <label>${tr('Главное впечатление','Main takeaway')}<select name="reason">
            ${[['keep-playing','Хочу продолжить','I want to keep playing'],['too-much-text','Слишком много текста','Too much text'],['unclear-choices','Непонятны последствия выбора','Choices feel unclear'],['not-my-story','История не зацепила','The story did not grab me'],['technical','Техническая проблема','A technical problem'],['other','Другое','Something else']].map(([v,ru,en])=>`<option value="${v}">${tr(ru,en)}</option>`).join('')}</select></label>
            <fieldset><legend>Premium · $6.99 ${tr('в месяц','/ month')}</legend><p>${tr('Новые ИИ-истории, выбор жанра и бесконечный сюжет. Это будущая цена, оплата не подключена.','New AI stories, genre choice and endless adventures. This is a proposed price; billing is not connected.')}</p>
            <label>${tr('Рассмотрели бы такую подписку?','Would you consider this subscription?')}<select name="interest"><option value="">${tr('Пока не отвечать','Skip this question')}</option><option value="yes">${tr('Да','Yes')}</option><option value="maybe">${tr('Возможно','Maybe')}</option><option value="no">${tr('Нет','No')}</option></select></label></fieldset>
            <label class="researchCheck"><input type="checkbox" name="consent" required>${tr('Согласен отправлять анонимные события игры и эти ответы.','I agree to share anonymous game events and these answers.')}</label>
            <p class="researchPrivacy">${tr('Случайный ID браузера, источник приглашения, режим, ходы и ответы — до 90 дней. Не отправляем имена, рабочие документы или текст ваших действий. Без стороннего трекера. Отозвать согласие можно здесь; это остановит новые события.','A random browser ID, invitation source, mode, turns and answers are kept for up to 90 days. No names, work documents or action text. No third-party tracker. Withdraw here to stop future events.')}</p>
            <div class="researchActions"><button type="submit" class="primary">${tr('Отправить и участвовать','Share and join')}</button><button type="button" class="researchStop">${tr('Не участвовать','Do not participate')}</button></div><p role="status" id="researchStatus"></p></form>`;
        const form=dialog.querySelector('form');
        form.elements.cohort.value=before?.cohort||(['team','jury'].includes(new URLSearchParams(location.search).get('source'))?'demo':'pilot');
        form.elements.eligible.checked=before?.eligible||false;
        form.elements.consent.checked=!!before;
        const close=()=>{dialog.close();if(trigger?.isConnected) trigger.focus({preventScroll:true});};
        dialog.querySelector('.researchClose').onclick=close;
        dialog.querySelector('.researchStop').onclick=()=>{try{localStorage.removeItem(KEY);}catch{}close();};
        form.onsubmit=async e=>{
            e.preventDefault();
            const button=form.querySelector('[type="submit"]');button.disabled=true;
            const source=new URLSearchParams(location.search).get('source');
            const current=settings();
            const prefs={consent:true,visitor:current?.visitor||before?.visitor||crypto.randomUUID(),cohort:form.elements.cohort.value,
                source:before?.source||(sources.includes(source)?source:'direct'),eligible:form.elements.eligible.checked,
                expires:Date.now()+90*86400000};
            try {localStorage.setItem(KEY,JSON.stringify(prefs));}
            catch {document.getElementById('researchStatus').textContent=tr('Браузер не разрешил сохранение. Игра доступна без участия.','This browser blocked storage. You can still play without joining.');button.disabled=false;return;}
            const responses=[event('enrolled')];
            if (premium || form.elements.interest.value) responses.push(event('premium_viewed'));
            if (form.elements.rating.value) responses.push(event('feedback',{rating:Number(form.elements.rating.value),reason:form.elements.reason.value}));
            if (form.elements.interest.value) responses.push(event('premium_interest',{interest:form.elements.interest.value}));
            const ok=(await Promise.all(responses)).every(Boolean);
            document.getElementById('researchStatus').textContent=ok ? tr('Спасибо! Это обратная связь, а не покупка. Возвращайтесь к истории в следующий перерыв.','Thank you! This is feedback, not a purchase. Come back to your story on your next break.') : tr('Ответ не доставлен полностью. Можно повторить позже; игра работает.','Some feedback could not be delivered. Try again later; the game still works.');
            button.disabled=false;
        };
        dialog.showModal();dialog.querySelector('.researchClose').focus();
    }
    globalThis.TabEscapeResearch={event,open,init(options){context=options.context;lang=options.language;event('page_view');}};
})();
