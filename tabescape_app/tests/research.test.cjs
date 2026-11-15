const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const records=[],store=new Map();
const ctx=vm.createContext({crypto:require('node:crypto').webcrypto,AbortSignal,Date,JSON,
 localStorage:{getItem:k=>store.get(k)||null},
 fetch:async(url,options)=>{records.push({url,payload:JSON.parse(options.body)});return{ok:true};}});
vm.runInContext(fs.readFileSync(path.join(__dirname,'../static/js/research.js'),'utf8'),ctx);
(async()=>{
 ctx.TabEscapeResearch.init({language:()=> 'en',context:()=>({workspace:'excel',story_token:'signed',adventure_id:'party',privateAction:'never send'})});
 await ctx.TabEscapeResearch.event('page_view');assert.equal(records.length,0);
 store.set('te_research_v1',JSON.stringify({consent:true,visitor:'random-id',cohort:'demo',source:'jury',eligible:false,expires:Date.now()+100000}));
 await ctx.TabEscapeResearch.event('turn_completed');assert.equal(records.length,1);
 assert.equal(records[0].payload.cohort,'demo');assert.equal(records[0].payload.story_token,'signed');assert.equal(records[0].payload.privateAction,undefined);
 store.delete('te_research_v1');await ctx.TabEscapeResearch.event('premium_viewed');assert.equal(records.length,1);
 console.log('Research sends nothing before consent or after withdrawal; jury cohort and minimal context preserved.');
})().catch(e=>{console.error(e);process.exitCode=1});
