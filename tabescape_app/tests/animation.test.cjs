const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname,'../static/js/app.js'),'utf8');
let now=10000, reduced=false, writes=0;
const visible={set textContent(v){this.value=v;writes++;}}, pending={set textContent(v){this.value=v;writes++;}};
const S={narrative:'A'.repeat(800)+'🌙',bossMode:false,readUntil:0,loading:null,sheet:'adv'};
const context=vm.createContext({S,Date:{now:()=>now},window:{matchMedia:()=>({matches:reduced})},
  document:{querySelectorAll:selector=>selector==='.sceneProse'?[{querySelector:s=>s==='.proseVisible'?visible:pending}]:[]},
  $:()=>({hidden:true}),setTimeout:()=>1,clearTimeout:()=>{},t:s=>s,esc:s=>s,
  offerPending:false,offerDismissed:false});
vm.runInContext(source.slice(source.indexOf('    let readingTimer;'),source.indexOf('    function logTurn(')),context);
context.pauseForReading(); context.paintProse();
assert.equal(visible.value,'');
now+=1200;context.paintProse();
assert(visible.value.length>0 && pending.value.length>0);
assert.equal(visible.value+pending.value,S.narrative);
S.bossMode=true;const before=writes;now+=100;context.paintProse();assert.equal(writes,before,'work screen must not be modified');
S.bossMode=false;now+=2000;context.paintProse();assert.equal(visible.value,S.narrative);assert.equal(pending.value,'');
reduced=true;context.pauseForReading();context.paintProse();assert.equal(visible.value,S.narrative,'reduced motion is immediate');
assert(source.includes('rc.includes("opt") ?'),'Excel click targets survive the reading lock');
console.log('Typewriter progress, work-mode isolation and reduced motion passed.');
