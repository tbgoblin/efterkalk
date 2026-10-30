const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const assert=require('node:assert/strict');
const asar=require('@electron/asar');
const directory=path.resolve(process.argv[2] || 'dist');
const version=process.argv[3] || require('../package.json').version;
const archive=path.join(directory,'win-unpacked','resources','app.asar');
assert.ok(fs.existsSync(archive),'Missing packaged application: '+archive);
const files=new Set(asar.listPackage(archive).map(file=>file.replace(/\\/g,'/').replace(/^\//,'')));
// @electron/asar splits lookups on path.sep, so nested paths need native separators on Windows.
const read=file=>asar.extractFile(archive,file.split('/').join(path.sep)).toString('utf8');
assert.equal(JSON.parse(read('package.json')).version,version);
for(const file of ['views/index.html','assets/monthly-summary.css','assets/js/monthly-summary.js',
    'assets/js/monthly-summary-model.js','assets/js/bilancio-formula.js','services/monthlySummaryService.js',
    'services/ubOrdersService.js','services/bilancioFormula.js','assets/js/kundefaktura.js']) {
    assert.ok(files.has(file),'Missing release file: '+file);
}
const html=read('views/index.html');
for(const match of html.matchAll(/(?:src|href)=["']\/(assets\/[^"'?#]+)(?:[^"']*)["']/g)) assert.ok(files.has(match[1]),'Missing linked asset: '+match[1]);
new vm.Script(html.slice(html.indexOf('<script>')+8,html.lastIndexOf('</script>')),{filename:'views/index.html:inline'});
let checked=0;
for(const file of files) {
    if(!file.endsWith('.js') || file.startsWith('node_modules/'))continue;
    new vm.Script('(function(require,module,exports,__dirname,__filename){\n'+read(file)+'\n})',{filename:file});
    checked++;
}
assert.ok(![...files].some(file=>file.startsWith('docs/audits/') || file.startsWith('.verify-') || file.startsWith('.claude/worktrees/')));
console.log('Packaged v'+version+' verified: HTML, referenced assets, summary modules and '+checked+' JavaScript files.');
