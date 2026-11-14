// Exercise the actual spreadsheet evaluator without a DOM or external packages.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '../static/js/app.js'), 'utf8');
const values = {A1: 10, A2: 20, A3: 30};
const context = vm.createContext({
    cellValue: (_sheet, ref) => values[ref] ?? 0,
});
const prefix = source.slice(0, source.indexOf('    function rawAt('));
vm.runInContext(prefix + '\nglobalThis.calculate = formula => evalFormula("imp0", formula, 0);\n})();', context);
for (const [formula, expected] of [
    ['=2+3*4', 14], ['=SUM(A1:A3)', 60], ['=СУММ(A1;A2)', 30],
    ['=IF(A1=10;"yes";"no")', 'yes'], ['=ROUND(2.126;2)', 2.13],
    ['=50%*A2', 10], ['=A3-A1', 20], ['=1/0', '#ДЕЛ/0!'],
]) {
    assert.equal(context.calculate(formula), expected, formula);
}
// A workbook formula must never obtain the JavaScript Function constructor.
assert.equal(context.calculate('=[]["filter"]["constructor"]("return 731")()'), '#ИМЯ?');
assert.equal(context.calculate('=`template`'), '#ИМЯ?');
assert.equal(context.calculate('=({})'), '#ИМЯ?');
console.log('Spreadsheet arithmetic, references and unsafe-expression checks passed.');
