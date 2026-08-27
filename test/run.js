// Registra la suite nello stesso processo: funziona anche negli ambienti
// aziendali/sandbox dove la creazione di processi figli è disabilitata.
const fs = require('fs');
const path = require('path');

for (const fileName of fs.readdirSync(__dirname).filter(name => name.endsWith('.test.js')).sort()) {
    require(path.join(__dirname, fileName));
}
