import ts from 'typescript';
import fs from 'fs';
import path from 'path';

const ROOT = path.resolve('./src/js');
const files = [];
(function walk(dir) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) walk(p);
    else if (e.name.endsWith('.js')) files.push(p);
  }
})(ROOT);

const SF = ts.SyntaxKind;
let total = 0;

for (const file of files) {
  const text = fs.readFileSync(file, 'utf-8');
  const sf = ts.createSourceFile(file, text, ts.ScriptTarget.ES2020, true, ts.ScriptKind.JS);
  const fileEdits = [];

  function visit(n) {
    if (ts.isForOfStatement(n)) {
      const expr = n.expression;
      if (expr && ts.isCallExpression(expr) && ts.isPropertyAccessExpression(expr.expression) && expr.expression.name.text === 'entries') {
        const X = expr.expression.expression;
        const decl = n.initializer;
        if (decl && ts.isVariableDeclarationList(decl) && decl.declarations.length === 1) {
          const d = decl.declarations[0];
          if (ts.isArrayBindingPattern(d.name) && d.name.elements.length === 2) {
            const e0 = d.name.elements[0], e1 = d.name.elements[1];
            if (e0 && e1 && ts.isBindingElement(e0) && ts.isBindingElement(e1) && ts.isIdentifier(e0.name) && ts.isIdentifier(e1.name)) {
              const A = e0.name.text, B = e1.name.text;
              let used = false;
              const body = n.statement;
              function scan(x) {
                if (used) return;
                if (ts.isIdentifier(x) && x.text === B) { used = true; return; }
                ts.forEachChild(x, scan);
              }
              scan(body);
              if (!used) {
                const Xtext = text.substring(X.getStart(sf), X.getEnd(sf));
                const newHeader = `for (const ${A} of ${Xtext}.keys())`;
                fileEdits.push({ start: n.getStart(sf), end: body.getStart(sf), text: newHeader });
                total++;
              }
            }
          }
        }
      }
    }
    ts.forEachChild(n, visit);
  }
  visit(sf);

  if (fileEdits.length === 0) continue;
  const kept = [];
  fileEdits.sort((a,b)=>a.start-b.start);
  for (const e of fileEdits) { if (kept.some(k=>e.start>=k.start&&e.end<=k.end)) continue; kept.push(e); }
  kept.sort((a,b)=>b.start-a.start);
  let out = text;
  for (const e of kept) out = out.slice(0, e.start) + e.text + out.slice(e.end);
  fs.writeFileSync(file, out);
}
fs.writeFileSync('/tmp/cleanup_stats.txt', `cleaned unused item: ${total}\n`);
