// Small arithmetic language. No JavaScript evaluation or executable expressions.
function parseFormula(expression) {
    if (typeof expression !== 'string' || !expression.trim() || expression.length > 2000) throw Error('Formel: skriv 1–2000 tegn.');
    const tokens = expression.match(/\[[A-Za-z][A-Za-z0-9_-]{0,63}\]|(?:\d+(?:\.\d+)?|\.\d+)|[+*/()%\-]|\s+|./g).filter(t => !/^\s+$/.test(t));
    let index = 0;
    const refs = new Set();
    function primary(depth) {
        if (depth > 40) throw Error('Formlen har for mange parenteser.');
        const token = tokens[index++];
        let node;
        if (token === '+' || token === '-') node = { op: 'unary' + token, left: primary(depth + 1) };
        else if (token === '(') { node = add(depth + 1); if (tokens[index++] !== ')') throw Error('Formel: manglende ).'); }
        else if (/^\[/.test(token || '')) { const id = token.slice(1, -1); refs.add(id); node = { ref: id }; }
        else if (/^(?:\d+(?:\.\d+)?|\.\d+)$/.test(token || '')) { node = { value: Number(token) }; if (!Number.isFinite(node.value)) throw Error('Ugyldigt tal.'); }
        else throw Error('Formel: brug tal, [række-id], +, -, *, /, % og parenteser.');
        while (tokens[index] === '%') { index++; node = { op: '/', left: node, right: { value: 100 } }; }
        return node;
    }
    function multiply(depth) {
        let node = primary(depth);
        while (['*', '/'].includes(tokens[index])) { const op = tokens[index++]; node = { op, left: node, right: primary(depth) }; }
        return node;
    }
    function add(depth) {
        let node = multiply(depth);
        while (['+', '-'].includes(tokens[index])) { const op = tokens[index++]; node = { op, left: node, right: multiply(depth) }; }
        return node;
    }
    const tree = add(0);
    if (index !== tokens.length) throw Error('Formel: uventet tegn eller manglende operator.');
    return { tree, refs: [...refs] };
}
function evaluateFormula(tree, resolve) {
    if ('value' in tree) return tree.value;
    if (tree.ref) return resolve(tree.ref);
    const left = evaluateFormula(tree.left, resolve);
    if (left == null) return null;
    if (tree.op === 'unary-') return -left;
    if (tree.op === 'unary+') return left;
    const right = evaluateFormula(tree.right, resolve);
    if (right == null || (tree.op === '/' && right === 0)) return null;
    const value = tree.op === '+' ? left + right : tree.op === '-' ? left - right : tree.op === '*' ? left * right : left / right;
    return Number.isFinite(value) ? value : null;
}
module.exports = { parseFormula, evaluateFormula };
