// Run actual persistence functions with browser storage that survives a fresh context.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '../static/js/app.js'), 'utf8');
const code = source.slice(source.indexOf('    const SAVE_KEY ='), source.indexOf('    function syncing()'));
const storage = new Map();
const localStorage = {getItem:k=>storage.get(k),setItem:(k,v)=>storage.set(k,v)};
function browser() {
    const context = vm.createContext({S:{}, localStorage, started:false, adventureId:'party-1',
        toast:message=>{context.warning=message;}, t:(_,en)=>en, Date,
        post:async()=>({story_token:'signed',narrative:'Same scene',stats:{ROI:80,KPI:75,Anxiety:20},
            inventory:['Kit'],story_turn:4,options:['A','B','C'],mode:'free',campaign_id:'oakhaven',
            campaign_title:'Oakhaven',genre:'fantasy',game_over:false})});
    vm.runInContext(code, context); return context;
}
(async()=>{
    let b=browser();
    Object.assign(b.S,{storyToken:'signed',campaignTitle:'Oakhaven',turn:5,campaignId:'oakhaven',
        playMode:'free',genre:'fantasy',log:[{action:'Help Lena',time:'11:00'}],history:[{ROI:80,KPI:75,Anxiety:20}]});
    b.saveAdventure();
    b=browser(); b.loadAdventure();
    assert.equal(b.started,true); assert.equal(b.S.log[0].action,'Help Lena');
    await b.resumeAdventure();
    assert.equal(b.S.turn,5); assert.equal(b.S.inventory.length,1);
    assert.equal(b.S.log.length,1); assert.equal(b.adventureId,'party-1');
    const committed=storage.get('te_adventure_v1');
    b.S.loading='Pending turn'; // Reload during a request recovers the last committed save.
    const fresh=browser(); fresh.loadAdventure(); assert.equal(fresh.S.turn,5);
    b.localStorage={getItem:localStorage.getItem,setItem:()=>{throw Error('quota');}};
    b.saveAdventure(); assert.match(b.warning,/Could not save/);
    assert.equal(storage.get('te_adventure_v1'),committed);
    storage.set('te_adventure_v1','broken'); b=browser();b.loadAdventure();
    assert.equal(b.started,false);assert.match(b.warning,/Could not read/);
    console.log('Save, reload, journal, inventory, interrupted turn, quota and corrupt-save checks passed.');
})().catch(error=>{console.error(error);process.exitCode=1;});
