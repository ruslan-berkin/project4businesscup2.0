// Exercise real turn orchestration with a controllable clock and network.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '../static/js/app.js'), 'utf8');
let now = 10000, networkResolve, networkReject, requests = 0, renders = 0, applied = 0;
let timers = [];
const state = {narrative:'The old scene',options:['Listen','Help','Leave'],readUntil:0,loading:null,
    over:false,bossMode:false,storyToken:'signed',seen:[],inventory:[],turn:1,sheet:'adv'};
const context = vm.createContext({S:state, Boolean, Date:{now:()=>now}, Promise,
    $:()=>null, document:{querySelectorAll:()=>[]},
    t:ru=>ru, esc:s=>s, clockStr:()=> '09:00', syncing:()=>{}, toast:()=>{},
    render:()=>{renders++;}, workspaceActionPending:false,
    setTimeout:(fn,delay)=>{const item={fn,at:now+delay};timers.push(item);return item;},
    clearTimeout:item=>{timers=timers.filter(x=>x!==item);},
    post:()=>{requests++;return new Promise((resolve,reject)=>{networkResolve=resolve;networkReject=reject;});},
    applyData:data=>{applied++;state.narrative=data.narrative;state.loading=null;context.pauseForReading();},
});
vm.runInContext(source.slice(source.indexOf('    let readingTimer;'),source.indexOf('    function logTurn('))+
    source.slice(source.indexOf('    async function sendAction('),source.indexOf('    function choose('))+
    '\nglobalThis.testLocked = turnLocked;', context);
async function advance(ms){now+=ms;const ready=timers.filter(x=>x.at<=now);timers=timers.filter(x=>x.at>now);ready.forEach(x=>x.fn());await Promise.resolve();await Promise.resolve();}
module.exports = (async()=>{
    const pending=context.sendAction('Listen',0);
    await context.sendAction('Help',1);
    assert.equal(requests,1,'rapid input must not create a second request');
    assert.equal(state.narrative,'The old scene','previous scene stays visible');
    networkResolve({narrative:'A new scene'});
    await advance(1199);
    assert.equal(applied,0,'fast response waits for the transition');
    state.bossMode=true;
    await advance(1);await pending;
    assert.equal(applied,1);assert.equal(state.turn,2);
    assert.equal(state.bossMode,true,'late response preserves work mode');
    assert.equal(renders,1,'late response does not repaint work view in orchestration');
    state.bossMode=false;
    await context.sendAction('Help',1);
    assert.equal(requests,1,'new scene has a brief input guard');
    await advance(1420);
    assert.equal(context.testLocked(),false);
    const failed=context.sendAction('Help',1);
    networkReject(new Error('unavailable'));
    assert.equal(await failed,false);
    assert.equal(state.narrative,'A new scene');
    assert.equal(state.turn,2,'failed request consumes no turn');
    assert.equal(context.testLocked(),false,'errors unlock retry');
    console.log('Turn pacing, duplicate input, work-mode response and retry checks passed.');
})();
