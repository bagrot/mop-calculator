(function () {
  const nav = document.querySelector('nav');
  const btn = document.createElement('button');
  btn.dataset.page = 'merge';
  btn.textContent = '⇄  Готовые МОПы';
  nav.insertBefore(btn, nav.children[1] || null);
  const section = document.createElement('section');
  section.className = 'page hidden'; section.id = 'merge';
  section.innerHTML = `<div class="eyebrow">ВТОРОЙ РЕЖИМ</div><h1>Сверка МОПов и реестра</h1><p class="muted">Загрузите рассчитанные МОПы и реестр голосов. Связь выполняется по кадастровому номеру.</p><div class="merge-grid"><div class="card"><h2>Файл подсчитанных МОПов</h2><input id="mopFile" type="file" accept=".xlsx,.xlsm"><p id="mopState" class="muted">Файл не выбран</p><div id="mopConfig" class="merge-config hidden"></div></div><div class="card"><h2>Реестр голосов</h2><input id="registryFile" type="file" accept=".xlsx,.xlsm"><p id="registryState" class="muted">Файл не выбран</p><div id="registryConfig" class="merge-config hidden"></div></div></div><div class="card hidden" id="mergePreview"><div class="cardHead"><div><h2>Предпросмотр объединения</h2><p id="mergeInfo" class="muted"></p></div><button id="runMerge" class="primary">Сформировать итоговый XLSX</button></div><div class="tableWrap"><table id="mergeTable"><thead></thead><tbody></tbody></table></div></div>`;
  document.querySelector('main').appendChild(section);
  const state = { mop: null, registry: null, config: null, merged: null };
  const $ = s => section.querySelector(s);
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const options = (headers, optional = false) => `${optional ? '<option value="">— не использовать —</option>' : ''}${headers.map((h,i)=>`<option value="${i}">${esc(h)}</option>`).join('')}`;
  function select(label, id, html) { return `<label>${label}<select id="${id}">${html}</select></label>`; }
  function bestSheet(data, kind) {
    let best = 0, score = -1;
    data.sheets.forEach((s, i) => { const h=s.headers.map(x=>String(x).toLowerCase()); let n=0;
      if (h.some(x=>x.includes('кадастр'))) n+=4;
      if (kind==='mop' && h.some(x=>x.includes('моп'))) n+=5;
      if (kind==='mop' && h.some(x=>x.includes('квартира'))) n+=2;
      if (kind==='registry' && h.some(x=>x.includes('id голоса'))) n+=6;
      if (n>score){score=n;best=i;}
    }); return best;
  }
  function renderConfig(kind) {
    const data = state[kind]; if (!data) return;
    const box = $(`#${kind}Config`); const sheetId = `${kind}Sheet`;
    const current = box.querySelector(`#${sheetId}`)?.value;
    const selected = current == null ? bestSheet(data, kind) : Number(current);
    box.innerHTML = select('Лист', sheetId, data.sheets.map((s,i)=>`<option value="${i}" ${i===selected?'selected':''}>${esc(s.name)} (${s.rows.length} строк)</option>`).join(''));
    box.classList.remove('hidden');
    const sheet = data.sheets[selected];
    if (kind === 'mop') box.innerHTML += `<div class="merge-fields">${select('Кадастровый номер', 'mopCad', options(sheet.headers))}${select('Площадь с МОП', 'mopArea', options(sheet.headers))}${select('Номер помещения — часть 1', 'mopRoom1', options(sheet.headers))}${select('Номер помещения — часть 2', 'mopRoom2', options(sheet.headers, true))}</div>`;
    else box.innerHTML += `<div class="merge-fields">${select('Кадастровый номер', 'regCad', options(sheet.headers))}</div>`;
    box.querySelector(`#${sheetId}`).onchange = () => renderConfig(kind);
    box.querySelectorAll('select:not(#'+sheetId+')').forEach(x => x.onchange = buildMerge);
    if (kind === 'mop') {
      const h=sheet.headers.map(x=>String(x).toLowerCase());
      const set=(id,tests)=>{const i=h.findIndex(x=>tests.some(t=>x.includes(t)));if(i>=0)box.querySelector('#'+id).value=String(i);};
      set('mopCad',['кадастр']); set('mopArea',['с моп','мопами']); set('mopRoom1',['квартира','№ квартиры']);
      const room2=box.querySelector('#mopRoom2'); const i2=h.findIndex(x=>x.includes('комната')); if(i2>=0)room2.value=String(i2);
    } else { const h=sheet.headers.map(x=>String(x).toLowerCase()); const i=h.findIndex(x=>x.includes('кадастр')); if(i>=0)box.querySelector('#regCad').value=String(i); }
    buildMerge();
  }
  async function readFile(input) { const f = input.files[0]; if (!f) return null; return window.mop.uploadGenericWorkbook(new Uint8Array(await f.arrayBuffer()), f.name); }
  $('#mopFile').onchange = async e => { $('#mopState').textContent = 'Загрузка…'; state.mop = await readFile(e.target); $('#mopState').textContent = state.mop.fileName; renderConfig('mop'); };
  $('#registryFile').onchange = async e => { $('#registryState').textContent = 'Загрузка…'; state.registry = await readFile(e.target); $('#registryState').textContent = state.registry.fileName; renderConfig('registry'); };
  function values(data, kind) { const sheet = data.sheets[Number($(`#${kind}Sheet`).value)]; return sheet; }
  function norm(v) { return String(v ?? '').trim().replace(/\s+/g, ' '); }
  function excelDate(v, withTime) { const n=Number(v); if(!Number.isFinite(n)||n<20000||n>70000)return v; const d=new Date(Date.UTC(1899,11,30)+n*86400000); const pad=x=>String(x).padStart(2,'0'); const out=`${pad(d.getUTCDate())}.${pad(d.getUTCMonth()+1)}.${d.getUTCFullYear()}`; return withTime?`${out} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`:out; }
  function outputValue(v, header) { const h=String(header||'').toLowerCase(); if(h.includes('дата'))return excelDate(v,false); if(h.includes('поступил')||h.includes('время'))return excelDate(v,true); if(h.includes('площадь с моп')){const n=Number(v);return Number.isFinite(n)?n.toFixed(3):'';} return v; }
  function buildMerge() {
    if (!state.mop || !state.registry || !$('#mopCad') || !$('#regCad')) return;
    const ms=values(state.mop,'mop'), rs=values(state.registry,'registry'); const mi={cad:+$('#mopCad').value,area:+$('#mopArea').value,room1:+$('#mopRoom1').value,room2:$('#mopRoom2').value===''?null:+$('#mopRoom2').value}; const ri={cad:+$('#regCad').value};
    const map=new Map(ms.rows.map(r=>[norm(r[mi.cad]),r])); const headers=[...rs.headers,'Актуальный номер помещения','Площадь с МОП']; const rawRows=rs.rows.map(r=>{const m=map.get(norm(r[ri.cad]));let room=m?norm(m[mi.room1]):'';if(m&&mi.room2!==null){const part2=norm(m[mi.room2]);if(part2)room=`${room} к.${part2}`;}return [...r,room,m?m[mi.area]:''];}); const rows=rawRows.map(r=>r.map((v,i)=>outputValue(v,headers[i]))); state.merged={headers,rows};
    $('#mergePreview').classList.remove('hidden'); $('#mergeInfo').textContent=`Строк реестра: ${rows.length} · найдено кадастров: ${rows.filter(r=>r.at(-1)!=='').length} · не найдено: ${rows.filter(r=>r.at(-1)==='').length}`;
    $('#mergeTable thead').innerHTML='<tr>'+headers.map(h=>`<th>${esc(h)}</th>`).join('')+'</tr>'; $('#mergeTable tbody').innerHTML=rows.slice(0,30).map(r=>'<tr>'+r.map(v=>`<td>${esc(v)}</td>`).join('')+'</tr>').join('');
  }
  $('#runMerge').onclick = async () => { if (!state.merged) return; const d=await window.mop.exportMerge(state.merged); $('#mergeInfo').textContent=`Файл сформирован: ${d.file}`; };
  btn.onclick = () => { document.querySelectorAll('.page').forEach(x=>x.classList.add('hidden')); section.classList.remove('hidden'); document.querySelectorAll('nav button').forEach(x=>x.classList.toggle('active',x===btn)); };
})();
const mergeTitle=document.querySelector('#merge h1');if(mergeTitle)mergeTitle.textContent='Готовые МОПы';
