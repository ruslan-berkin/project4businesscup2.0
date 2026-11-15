// Exercise accepted and failed starts: no lost party, duplicate request or premature reset.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname,'../static/js/app.js'),'utf8');
let resolve, reject, requests=0;
const state={storyToken:'old-signed',narrative:'Old scene',inventory:['Kit'],log:[{action:'Help'}],
    loading:null,over:true,role:'Hero',boss:'Mentor',campaignId:'oakhaven',playMode:'free',genre:'fantasy'};
const context=vm.createContext({S:state, START_STATS:{ROI:100,KPI:100,Anxiety:10},
    adventureId:'previous-party',offerPending:true,offerDismissed:true,cellEls:[],
    render:()=>{},syncing:()=>{},t:(_,en)=>en,$:()=>({open:true}),toast:()=>{},
    post:()=>{requests++;return new Promise((yes,no)=>{resolve=yes;reject=no;});},
    applyData:data=>Object.assign(state,{storyToken:data.story_token,narrative:data.narrative,loading:null})});
vm.runInContext(source.slice(source.indexOf('    let launchError ='),source.indexOf('    async function sendAction(')),context);
(async()=>{
    const first=context.restart();
    assert.equal(state.storyToken,'old-signed');assert.equal(state.log.length,1);
    assert.equal(await context.restart(),false);assert.equal(requests,1);
    reject(new Error('Offline'));assert.equal(await first,false);
    assert.equal(state.storyToken,'old-signed');assert.equal(state.inventory[0],'Kit');
    assert.equal(context.adventureId,'previous-party');assert.equal(state.loading,null);
    const retry=context.restart();resolve({story_token:'new-signed',narrative:'New opening'});
    assert.equal(await retry,true);assert.equal(state.storyToken,'new-signed');
    assert.equal(state.log.length,0);assert.equal(state.inventory.length,0);
    assert.notEqual(context.adventureId,'previous-party');
    console.log('Launch keeps the previous party on failure, blocks duplicate starts and commits only on success.');
})().catch(error=>{console.error(error);process.exitCode=1;});

// A late demo-file response must never overwrite a user's workbook.
async function checkWorkbookPreparation() {
    let finishFetch, calls=0;
    const state={imported:null};
    const ctx=vm.createContext({S:state, MASKS:{excel:{file:'Demo.xlsx'}},
        arrayBufferToBase64:()=> 'demo', parseXlsx:async()=>[{name:'Forecast'}],
        fetchWithTimeout:()=>{calls++;return new Promise(resolve=>{finishFetch=resolve;});}});
    vm.runInContext(source.slice(source.indexOf('    let defaultWorkbookRequest ='),source.indexOf('    function loadCatalog(')),ctx);
    const pending=ctx.prepareDefaultWorkbook();
    const same=ctx.prepareDefaultWorkbook();assert.equal(calls,1);
    const userBook={name:'My workbook.xlsx',sheets:[{name:'My sheet'}]};state.imported=userBook;
    finishFetch({ok:true,arrayBuffer:async()=>new ArrayBuffer(1)});
    await Promise.all([pending,same]);assert.equal(state.imported,userBook);
    const failed=vm.createContext({S:{imported:null},fetchWithTimeout:async()=>{throw new Error('Timed out');}});
    vm.runInContext(source.slice(source.indexOf('    let defaultWorkbookRequest ='),source.indexOf('    function loadCatalog(')),failed);
    await failed.prepareDefaultWorkbook();assert.equal(failed.S.imported,null);
    console.log('Workbook preparation reuses one request, preserves an uploaded book and allows built-in fallback.');
}
checkWorkbookPreparation().catch(error=>{console.error(error);process.exitCode=1;});
