import { app, BrowserWindow, dialog, ipcMain, shell } from 'electron';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '..');
let win;

async function createWindow() {
  process.env.MOP_APP_ROOT = root;
  process.env.MOP_NO_SERVER = '1';
  process.env.MOP_MERGE_RESULTS_DIR = path.join(app.getPath('documents'), 'Обработка голосов', 'results', 'Готовые МОПы');
  process.env.MOP_RESULTS_DIR = path.join(app.getPath('documents'), 'Обработка голосов', 'results', 'Калькулятор МОП');
  process.env.MOP_RESULTS_DIR = path.join(app.getPath('documents'), 'Обработка голосов', 'results');
  const module = await import(pathToFileURL(path.join(root, 'server.mjs')).href);
  ipcMain.handle('upload-workbook', async (_event, bytes, name) => {
    const filePath = path.join(os.tmpdir(), `mop-${Date.now()}-${String(name || 'upload.xlsx').replace(/[^\\w.()-]/g, '_')}`);
    await fs.writeFile(filePath, Buffer.from(bytes));
    try { return await module.parseWorkbook(filePath); } finally { await fs.rm(filePath, { force: true }); }
  });
  ipcMain.handle('export-workbook', async (_event, model) => ({ file: await module.exportResult(model) }));
  ipcMain.handle('upload-generic-workbook', async (_event, bytes, name) => {
    const filePath = path.join(os.tmpdir(), `mop-generic-${Date.now()}-${String(name || 'upload.xlsx').replace(/[^\\w.()-]/g, '_')}`);
    await fs.writeFile(filePath, Buffer.from(bytes));
    try { const data = await module.parseGenericWorkbook(filePath); data.fileName = name || data.fileName; return data; } finally { await fs.rm(filePath, { force: true }); }
  });
  ipcMain.handle('export-merge', async (_event, model) => ({ file: await module.exportMerge(model) }));
  ipcMain.handle('open-results', async () => { const dir = path.join(app.getPath('documents'), 'Обработка голосов', 'results'); await fs.mkdir(dir, { recursive: true }); return shell.openPath(dir); });
  win = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1100,
    minHeight: 700,
    backgroundColor: '#f4f6f8',
    title: 'Обработка голосов',
    autoHideMenuBar: true,
    webPreferences: { contextIsolation: true, nodeIntegration: false, preload: path.join(here, 'preload.cjs') },
  });
  await win.loadFile(path.join(root, 'public', 'index.html'));
  const styleUrl = pathToFileURL(path.join(root, 'public', 'style.css')).href;
  const scriptUrl = pathToFileURL(path.join(root, 'public', 'app.js')).href;
  await win.webContents.executeJavaScript(`
    const link=document.createElement('link'); link.rel='stylesheet'; link.href=${JSON.stringify(styleUrl)}; document.head.appendChild(link);
    const nativeFetch=window.fetch;
    window.fetch=async (url, options={}) => {
      if (url === '/api/upload') { const file=options.body?.get('file'); const data=await window.mop.uploadWorkbook(new Uint8Array(await file.arrayBuffer()), file.name); return new Response(JSON.stringify(data), {headers:{'Content-Type':'application/json'}}); }
      if (url === '/api/export') { const data=await window.mop.exportWorkbook(JSON.parse(options.body)); return new Response(JSON.stringify(data), {headers:{'Content-Type':'application/json'}}); }
      return nativeFetch(url, options);
    };
    const script=document.createElement('script'); script.src=${JSON.stringify(scriptUrl)}; document.body.appendChild(script);
  `);
  win.on('closed', () => { win = null; });
}

app.whenReady().then(createWindow).catch(error => {
  dialog.showErrorBox('Не удалось запустить приложение', `${error?.message || error}\n\nПопробуйте закрыть старую версию приложения и запустить снова.`);
  app.quit();
});
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
app.on('activate', () => { if (!win) createWindow(); });
app.on('web-contents-created', (_event, contents) => {
  contents.setWindowOpenHandler(({ url }) => { shell.openExternal(url); return { action: 'deny' }; });
});
