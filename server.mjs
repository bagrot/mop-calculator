import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { FileBlob, SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const root = process.env.MOP_APP_ROOT || process.cwd();
const publicDir = path.join(root, 'public');
const port = Number(process.env.MOP_PORT || 9222);
const resultsDir = process.env.MOP_RESULTS_DIR || path.join(root, 'results');
const mergeResultsDir = process.env.MOP_MERGE_RESULTS_DIR || path.join(root, 'results', 'Готовые МОПы');

const headers = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Content-Type', 'Access-Control-Allow-Methods': 'GET,POST,OPTIONS' };
const send = (res, status, body, type='application/json; charset=utf-8') => { res.writeHead(status, { ...headers, 'Content-Type': type }); res.end(typeof body === 'string' ? body : JSON.stringify(body)); };
const readBody = req => new Promise((resolve, reject) => { const chunks=[]; req.on('data', c=>chunks.push(c)); req.on('end',()=>resolve(Buffer.concat(chunks))); req.on('error', reject); });

function parseMultipart(body, contentType) {
  const match = contentType.match(/boundary=(?:"([^"]+)"|([^;]+))/i); if (!match) throw Error('Не найден boundary загрузки');
  const boundary = Buffer.from(`--${match[1] || match[2]}`); const start = body.indexOf(boundary); if (start < 0) throw Error('Пустая загрузка');
  const headerStart = start + boundary.length + 2; const headerEnd = body.indexOf(Buffer.from('\r\n\r\n'), headerStart); if (headerEnd < 0) throw Error('Неверный формат загрузки');
  const headersText = body.subarray(headerStart, headerEnd).toString('utf8'); const fileName = (headersText.match(/filename="([^"]+)"/i) || [,'upload.xlsx'])[1];
  const dataStart = headerEnd + 4; const end = body.indexOf(Buffer.from(`\r\n--${boundary.toString('utf8').slice(2)}`), dataStart); if (end < 0) throw Error('Не найден конец файла');
  return { fileName, data: body.subarray(dataStart, end) };
}

const cell = (v) => v instanceof Date ? v.toISOString() : v;
const normalize = v => String(v ?? '').replace(/\s+/g, ' ').trim();
const colNames = ['ID голоса','ID голоса по помещению','Дата голоса','Форма голоса','Поступил в систему','SSOID/CORPID','Собственник','Дата рождения','Кадастровый номер','№ помещения','Общая площадь помещения','Доля в собственности','Учтенная площадь','Площадь помещения в собственности','Тип собственности','Номер правоустанавливающего документа','Дата правоустанавливающего документа','Количество голосов (в %)','Общее количество голосов МКД (100%)','Статус помещения','Статус голоса','Статус подтвержден модератором'];

async function parseWorkbook(filePath) {
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(filePath));
  const meta = wb.worksheets.getItemOrNullObject?.('Данные о выгрузке');
  const sheet = wb.worksheets.items.find(s => s.name === 'Реестр голосов') || wb.worksheets.items[0];
  const used = sheet.getUsedRange(); const values = used.values || [];
  const header = values[0]?.map(normalize) || []; const indexes = colNames.map(n => header.indexOf(n));
  const rows = values.slice(1).filter(r => r.some(v => v != null && v !== '')).map((r, i) => {
    const o = { row: i + 2 }; colNames.forEach((n, j) => o[n] = cell(indexes[j] >= 0 ? r[indexes[j]] : '')); return o;
  });
  let oss = '';
  const metaSheet = wb.worksheets.items.find(s => s.name === 'Данные о выгрузке');
  if (metaSheet) { const mv = metaSheet.getUsedRange()?.values || []; const rr = mv.find(r => normalize(r[0]) === 'Номер ОСС'); oss = rr ? normalize(rr[1]) : ''; }
  return { fileName: path.basename(filePath), sheet: sheet.name, oss, rows, headers: colNames };
}

async function parseGenericWorkbook(filePath) {
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(filePath));
  return {
    fileName: path.basename(filePath),
    sheets: wb.worksheets.items.map(sheet => {
      const values = sheet.getUsedRange()?.values || [];
      const headers = (values[0] || []).map((v, i) => String(v ?? '').trim() || `Колонка ${i + 1}`);
      return { name: sheet.name, headers, rows: values.slice(1).filter(r => r.some(v => v != null && v !== '')) };
    }),
  };
}

async function exportMerge(model) {
  const wb = Workbook.create();
  const sheet = wb.worksheets.add('Объединенный реестр');
  const headers = model.headers || [];
  const rows = model.rows || [];
  sheet.getRange(`A1:${String.fromCharCode(64 + Math.max(1, headers.length))}${Math.max(1, rows.length + 1)}`).values = [headers, ...rows];
  sheet.getUsedRange().format = { font: { name: 'Calibri', size: 10 }, wrapText: false };
  sheet.getRange('1:1').format = { fill: '#273047', font: { bold: true, color: '#FFFFFF' } };
  sheet.getUsedRange().format.autofitColumns();
  const dir = mergeResultsDir; await fs.mkdir(dir, { recursive: true });
  const out = path.join(dir, `Объединенный_реестр_${Date.now()}.xlsx`);
  const blob = await SpreadsheetFile.exportXlsx(wb); await blob.save(out); return out;
}

async function exportResult(model) {
  const wb = Workbook.create(); const result = wb.worksheets.add('Итоговый реестр'); const manual = wb.worksheets.add('Дополнительные комнаты');
  const outputHeaders = ['Номер ОСС','ID голоса','ID голоса по помещению','Форма голоса','Кадастровый номер','Актуальный номер помещения','Общая площадь','Площадь МОП комнаты','Площадь с МОП','Доля в собственности','Площадь в собственности','Тип собственности','Номер права','Дата права'];
  const areaHeaders = new Set(['Общая площадь','Площадь МОП комнаты','Площадь с МОП','Площадь в собственности']);
  const decimal = value => { if (value === null || value === undefined || value === '') return ''; const n = Number(value); if (!Number.isFinite(n)) return value; return n.toFixed(3).replace(/0+$/, '').replace(/\.$/, ''); };
  const resultRows = model.results.map(r => outputHeaders.map(h => areaHeaders.has(h) ? decimal(r[h]) : (r[h] ?? '')));
  result.getRange(`A1:N${Math.max(1, model.results.length + 1)}`).values = [outputHeaders, ...resultRows];
  manual.getRange(`A1:E${Math.max(1, model.manual.length + 1)}`).values = [['Исходное помещение','Актуальный номер помещения','Площадь','Комментарий','Источник'], ...model.manual.map(r => [r.source, r.actual, decimal(r.area), r.comment || '', 'Добавлено вручную'])];
  for (const s of [result, manual]) { s.getUsedRange().format = { font: { name: 'Calibri', size: 10 }, wrapText: false }; s.getRange('1:1').format = { fill: '#273047', font: { bold: true, color: '#FFFFFF' } }; s.getUsedRange().format.autofitColumns(); }
  const dir = resultsDir; await fs.mkdir(dir, { recursive: true }); const out = path.join(dir, `МОП_${Date.now()}.xlsx`); const blob = await SpreadsheetFile.exportXlsx(wb); await blob.save(out); return out;
}

const server = http.createServer(async (req, res) => {
  try {
    if (req.method === 'OPTIONS') return send(res, 204, '');
    if (req.method === 'GET' && req.url === '/') return send(res, 200, await fs.readFile(path.join(publicDir, 'index.html'), 'utf8'), 'text/html; charset=utf-8');
    if (req.method === 'GET' && req.url.startsWith('/public/')) return send(res, 200, await fs.readFile(path.join(root, req.url.slice(1)), 'utf8'), req.url.endsWith('.js') ? 'text/javascript' : 'text/css');
    if (req.method === 'POST' && req.url === '/api/upload') { const body=await readBody(req); const f=parseMultipart(body, req.headers['content-type']||''); const tmp=path.join(os.tmpdir(), `mop-${Date.now()}-${f.fileName}`); await fs.writeFile(tmp,f.data); const data=await parseWorkbook(tmp); await fs.unlink(tmp); return send(res,200,data); }
    if (req.method === 'POST' && req.url === '/api/export') { const model=JSON.parse((await readBody(req)).toString('utf8')); const out=await exportResult(model); return send(res,200,{ file: path.relative(root,out).replaceAll('\\','/') }); }
    send(res,404,{error:'Not found'});
  } catch (e) { send(res,500,{error:e.message}); }
});
export { server, parseWorkbook, parseGenericWorkbook, exportResult, exportMerge };
if (!process.env.MOP_NO_SERVER) server.listen(port, '127.0.0.1', () => console.log(`МОП app: http://127.0.0.1:${server.address().port}`));
