/* FF14 潛水艇計時器：數據總表（xlsx）→ 網站 data/*.js
 * 瀏覽器（tools/excel-to-data.html）與 Node 都可以使用：convertWorkbook(XLSX, workbook)
 * 回傳 { files: {檔名: 內容}, warnings: [...], summary: {...}, version } */
(function (root) {
  const FILES = {
    'sectors.js': ['海域、地點、航行距離、等級、部件', ['MAP_NAMES', 'SUB_HOMES', 'SUB_SECTORS', 'SUB_PATCH', 'SUB_TANK', 'SUB_DIST', 'SUB_RANKS', 'SUB_CAP', 'PARTS']],
    'loot.js': ['掉落物：層級、機率、數量', ['SUB_LOOT', 'SUB_LOOT_TIER', 'GP_TIER_PROC', 'GP_TIER_PROC_D', 'GP_ITEM_PROB', 'SECTOR_LOOT']],
    'items.js': ['物品名稱、高優先物品、常見材料', ['ITEM_DB', 'ITEM_DB_EN', 'PRIORITY_LOOT_GROUPS', 'COMMON_LOOT']],
    'presets.js': ['常用配件預設', ['GP_PRESETS', 'GP_PRESET_HINTS']],
    'craft.js': ['配件製作配方', ['CRAFT_PARTS', 'CRAFT_MATS', 'CRAFT_BASE']],
  };
  const PART_KEY = { '船體': 'hull', '船尾': 'stern', '船首': 'bow', '艦橋': 'bridge' };

  function blank(v) { return v === null || v === undefined || (typeof v === 'string' && v.trim() === ''); }
  function num(v) { if (blank(v)) return null; const n = Number(String(v).replace(/,/g, '')); return Number.isFinite(n) ? n : null; }
  function str(v) { return blank(v) ? null : String(v).trim(); }

  function convertWorkbook(XLSX, wb) {
    const warnings = [];
    const warn = (msg) => { if (warnings.length < 300) warnings.push(msg); };

    function table(name, required) {
      const ws = wb.Sheets[name];
      if (!ws) throw new Error(`找不到「${name}」分頁`);
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: false });
      const head = (rows[0] || []).map(h => (h === null ? '' : String(h).trim()));
      for (const r of required) if (!head.includes(r)) throw new Error(`「${name}」分頁缺少欄位「${r}」`);
      const idx = Object.fromEntries(head.map((h, i) => [h, i]));
      return rows.slice(1).filter(r => r.some(v => !blank(v))).map((r, i) => {
        const o = { _row: i + 2 }; for (const h of head) if (h) o[h] = r[idx[h]]; return o;
      });
    }

    // ── 海域 ──
    const MAP_NAMES = {}, SUB_HOMES = {}, coords = {};
    for (const r of table('海域', ['海域ID', '名稱', '起點節點ID'])) {
      const m = num(r['海域ID']), h = num(r['起點節點ID']);
      if (m === null || h === null) { warn(`海域 第${r._row}列：海域ID 或起點節點ID 空白，已略過`); continue; }
      MAP_NAMES[m] = str(r['名稱']); SUB_HOMES[m] = h;
      const x = num(r['起點X']), y = num(r['起點Y']), z = num(r['起點Z']);
      if (x !== null && y !== null && z !== null) coords[h] = [x, y, z];
    }

    // ── 地點 ──
    const SUB_SECTORS = [], SUB_PATCH = {}, SUB_TANK = {}, GP_TIER_PROC = {}, GP_TIER_PROC_D = {}, slThr = {};
    const seenId = new Set();
    for (const r of table('地點', ['地點ID', '海域ID', '字母', '中文名', '英文名', '解鎖等級', '經驗值', '偵查時間(分鐘×70)', '偵查航距消耗',
      '恩惠門檻', 'T2探索門檻', 'T3探索門檻', '收集門檻(普通)', '收集門檻(優秀)'])) {
      const id = num(r['地點ID']), m = num(r['海域ID']);
      if (id === null || m === null) { warn(`地點 第${r._row}列：地點ID 或海域ID 空白，已略過`); continue; }
      if (seenId.has(id)) { warn(`地點 第${r._row}列：地點ID ${id} 重複，已略過`); continue; }
      if (!(m in MAP_NAMES)) warn(`地點 第${r._row}列：海域ID ${m} 不在「海域」分頁裡`);
      seenId.add(id);
      const letter = str(r['字母']);
      SUB_SECTORS.push([id, m, letter, str(r['英文名']), str(r['中文名']), num(r['解鎖等級']), num(r['經驗值']),
        num(r['偵查時間(分鐘×70)']), num(r['偵查航距消耗']), num(r['恩惠門檻']), num(r['T2探索門檻']), num(r['T3探索門檻']),
        num(r['收集門檻(普通)']), num(r['收集門檻(優秀)'])]);
      if (!blank(r['開放版本'])) SUB_PATCH[id] = String(r['開放版本']).trim();
      if (num(r['青磷水消耗']) !== null) SUB_TANK[id] = num(r['青磷水消耗']);
      const x = num(r['座標X']), y = num(r['座標Y']), z = num(r['座標Z']);
      if (x !== null && y !== null && z !== null) coords[id] = [x, y, z];
      const tp = ['T1 Proc%', 'T2 Proc%', 'T3 Proc%'].map(k => num(r[k])), td = ['T1 D%', 'T2 D%', 'T3 D%'].map(k => num(r[k]));
      if (tp.some(v => v !== null)) GP_TIER_PROC[id] = tp;
      if (td.some(v => v !== null)) GP_TIER_PROC_D[id] = td;
      const a = num(r['數量表收集門檻(普通)']), b = num(r['數量表收集門檻(優秀)']);
      if (a !== null || b !== null) slThr[`${m}_${letter}`] = [a, b];
      ['解鎖等級', '經驗值', '偵查時間(分鐘×70)', '偵查航距消耗'].forEach(k => { if (num(r[k]) === null) warn(`地點 ${id}（${str(r['中文名']) || letter}）：「${k}」空白`); });
    }
    SUB_SECTORS.sort((a, b) => a[0] - b[0]);
    const secById = new Map(SUB_SECTORS.map(s => [s[0], s]));
    const secByKey = new Map(SUB_SECTORS.map(s => [`${s[1]}_${s[2]}`, s]));

    // ── 路線距離：有座標就自動計算（直線距離 × 39.9／× 0.035 取整），否則用「路線距離」分頁 ──
    const sheetDist = {};
    if (wb.Sheets['路線距離']) for (const r of table('路線距離', ['出發ID', '目的ID', '航行距離', '航距消耗'])) {
      const f = num(r['出發ID']), t = num(r['目的ID']);
      if (f !== null && t !== null && num(r['航行距離']) !== null) sheetDist[`${f}>${t}`] = [num(r['航行距離']), num(r['航距消耗']) || 0];
    }
    const SUB_DIST = {}; let computed = 0, fromSheet = 0, missing = 0;
    const calc = (a, b) => { const e = Math.hypot(a[0] - b[0], a[1] - b[1], a[2] - b[2]); return [Math.floor(e * 39.9), Math.floor(e * 0.035)]; };
    for (const [m, home] of Object.entries(SUB_HOMES)) {
      const ids = [home, ...SUB_SECTORS.filter(s => s[1] === +m).map(s => s[0])];
      for (const f of ids) {
        const row = {};
        for (const t of ids) {
          if (t === f || t === home) continue;
          let v = null;
          if (coords[f] && coords[t]) { v = calc(coords[f], coords[t]); computed++; }
          else if (sheetDist[`${f}>${t}`]) { v = sheetDist[`${f}>${t}`]; fromSheet++; }
          else missing++;
          if (v) row[t] = v;
        }
        if (f !== home) {
          let v = null;
          if (coords[f] && coords[home]) { v = calc(coords[f], coords[home]); computed++; }
          else if (sheetDist[`${f}>${home}`]) { v = sheetDist[`${f}>${home}`]; fromSheet++; }
          else missing++;
          if (v) row[home] = v;
        }
        SUB_DIST[f] = row;
      }
    }
    if (missing) warn(`有 ${missing} 組地點之間沒有航行距離（缺座標，「路線距離」分頁也沒有）；這些組合不會出現在路線規劃`);

    // ── 等級 ──
    const SUB_RANKS = {}, SUB_CAP = {};
    for (const r of table('等級', ['等級', '升級所需經驗', '部件容量上限', '探索', '收集', '航速', '航距', '恩惠'])) {
      const lv = num(r['等級']); if (lv === null) continue;
      SUB_RANKS[lv] = [num(r['升級所需經驗']), num(r['航速']), num(r['航距']), num(r['探索']), num(r['收集']), num(r['恩惠'])].map(v => v === null ? 0 : v);
      if (num(r['部件容量上限']) !== null) SUB_CAP[lv] = num(r['部件容量上限']);
    }

    // ── 部件 ──
    const PARTS = {};
    for (const r of table('部件', ['部位代碼', '英文代碼', '級別', '探索', '收集', '航速', '航距', '恩惠', '造價(容量)'])) {
      const p = str(r['部位代碼']) || PART_KEY[str(r['部位'])];
      if (!p) { warn(`部件 第${r._row}列：部位代碼空白`); continue; }
      (PARTS[p] = PARTS[p] || []).push([str(r['英文代碼']), str(r['級別']), ...['探索', '收集', '航速', '航距', '恩惠', '造價(容量)'].map(k => num(r[k]))]);
    }

    // ── 配件預設 ──
    const GP_PRESETS = {}, GP_PRESET_HINTS = {};
    for (const r of table('配件預設', ['預設代碼', '船體', '船尾', '船首', '艦橋'])) {
      const k = str(r['預設代碼']); if (!k) continue;
      GP_PRESETS[k] = { hull: String(r['船體'] ?? ''), stern: String(r['船尾'] ?? ''), bow: String(r['船首'] ?? ''), bridge: String(r['艦橋'] ?? '') };
      if (str(r['說明'])) GP_PRESET_HINTS[k] = str(r['說明']);
    }

    // ── 物品 ──
    const ITEM_DB = {}, ITEM_DB_EN = {}, PRIORITY_LOOT_GROUPS = {}, common = [];
    for (const r of table('物品', ['物品ID', '中文名', '英文名'])) {
      const id = num(r['物品ID']); if (id === null) continue;
      if (str(r['中文名'])) ITEM_DB[id] = str(r['中文名']);
      if (str(r['英文名'])) ITEM_DB_EN[id] = str(r['英文名']);
      const g = str(r['高優先分類']); if (g) (PRIORITY_LOOT_GROUPS[g] = PRIORITY_LOOT_GROUPS[g] || []).push(id);
      const c = num(r['常見材料順序']); if (c !== null && ITEM_DB[id]) common.push([c, ITEM_DB[id]]);
    }
    const COMMON_LOOT = common.sort((a, b) => a[0] - b[0]).map(x => x[1]);
    const enToId = new Map(Object.entries(ITEM_DB_EN).map(([k, v]) => [v, +k]));

    // ── 掉落 ──
    const SUB_LOOT = {}, SUB_LOOT_TIER = {}, GP_ITEM_PROB = {}, SECTOR_LOOT = {};
    for (const [k, v] of Object.entries(slThr)) SECTOR_LOOT[k] = [v[0], v[1], {}];
    for (const r of table('掉落', ['地點ID', '門檻層級', '物品ID', '英文名', '路線用', '在掉落清單'])) {
      const sid = num(r['地點ID']); const sec = secById.get(sid);
      if (!sec) { warn(`掉落 第${r._row}列：地點ID ${sid} 不在「地點」分頁裡，已略過`); continue; }
      const tierTxt = str(r['門檻層級']); const tier = tierTxt ? tierTxt.replace(/^T/i, '') : null;
      if (tier && !['1', '2', '3'].includes(tier)) { warn(`掉落 第${r._row}列：門檻層級「${tierTxt}」只能是 T1／T2／T3`); continue; }
      let iid = num(r['物品ID']); let en = str(r['英文名']);
      if (iid === null && en && enToId.has(en)) iid = enToId.get(en);
      if (iid !== null && !en) en = ITEM_DB_EN[iid] || null;
      if (iid !== null && !(iid in ITEM_DB) && !(iid in ITEM_DB_EN)) warn(`掉落 第${r._row}列：物品ID ${iid} 不在「物品」分頁裡`);
      const yes = v => str(v) && /^(y|yes|v|✓|是|1)$/i.test(str(v));
      if (yes(r['路線用'])) {
        if (!tier || iid === null) warn(`掉落 第${r._row}列：「路線用」需要門檻層級和物品ID`);
        else { const t = (SUB_LOOT_TIER[sid] = SUB_LOOT_TIER[sid] || {}); (t[tier] = t[tier] || []).push(iid); }
      }
      if (yes(r['在掉落清單']) && iid !== null) (SUB_LOOT[sid] = SUB_LOOT[sid] || []).push(iid);
      const tp = num(r['T%(層內機率)']), dp = num(r['D%(直接機率)']);
      if ((tp !== null || dp !== null) && tier && iid !== null) {
        const g = (GP_ITEM_PROB[sid] = GP_ITEM_PROB[sid] || {}); (g[tier] = g[tier] || {})[iid] = { t: tp, d: dp };
      }
      const q = ['收集不足-最少', '收集不足-最多', '收集普通-最少', '收集普通-最多', '收集優秀-最少', '收集優秀-最多'].map(k => num(r[k]));
      if (q.some(v => v !== null) && tier) {
        if (!en) { warn(`掉落 第${r._row}列：有數量但沒有英文名，數量表需要英文名`); continue; }
        const key = `${sec[1]}_${sec[2]}`;
        const e = (SECTOR_LOOT[key] = SECTOR_LOOT[key] || [null, null, {}]);
        const z = v => v === null ? 0 : v;
        (e[2]['t' + tier] = e[2]['t' + tier] || []).push({ n: en, p: [z(q[0]), z(q[1])], nm: [z(q[2]), z(q[3])], o: [z(q[4]), z(q[5])] });
      }
    }
    for (const k of Object.keys(SUB_LOOT)) SUB_LOOT[k] = [...new Set(SUB_LOOT[k])].sort((a, b) => a - b);
    for (const sid of Object.keys(SUB_LOOT_TIER)) {
      const t = SUB_LOOT_TIER[sid]; const o = {};
      for (const k of ['1', '2', '3']) o[k] = t[k] || [];
      SUB_LOOT_TIER[sid] = o;
    }
    for (const v of Object.values(SECTOR_LOOT)) for (const t of Object.keys(v[2])) v[2][t].sort((a, b) => (a.n < b.n ? -1 : a.n > b.n ? 1 : 0));
    for (const s of SUB_SECTORS) if (!SUB_LOOT_TIER[s[0]]) warn(`地點 ${s[0]}（${s[4]}）沒有任何「路線用」掉落物`);

    // ── 製作 ──
    const CRAFT_PARTS = {}, CRAFT_MATS = {}, CRAFT_BASE = {};
    for (const r of table('製作-部件', ['部件名稱', '級別', '部位', '級別編號', '排序', '物品等級', '素材', '數量'])) {
      const n = str(r['部件名稱']); if (!n) continue;
      const p = CRAFT_PARTS[n] || (CRAFT_PARTS[n] = { class: num(r['級別編號']), className: str(r['級別']), part: str(r['部位']), sort: num(r['排序']),
        recipe: {}, frame: str(r['框架']), frameQty: num(r['框架數量']), itemLevel: num(r['物品等級']) });
      if (str(r['素材'])) p.recipe[str(r['素材'])] = num(r['數量']);
    }
    for (const r of table('製作-素材', ['素材名稱', '分類', '製作職業', '材料', '數量'])) {
      const n = str(r['素材名稱']); if (!n) continue;
      let m = CRAFT_MATS[n];
      if (!m) {
        m = CRAFT_MATS[n] = { category: str(r['分類']), job: str(r['製作職業']), recipe: {} };
        if (num(r['製作階層']) !== null) m.level = num(r['製作階層']);
        if (num(r['產出數量']) !== null) m.yield = num(r['產出數量']);
      }
      if (str(r['材料'])) m.recipe[str(r['材料'])] = num(r['數量']);
    }
    for (const r of table('基礎素材', ['名稱', '分類'])) if (str(r['名稱'])) CRAFT_BASE[str(r['名稱'])] = str(r['分類']);

    const data = { MAP_NAMES, SUB_HOMES, SUB_SECTORS, SUB_PATCH, SUB_TANK, SUB_DIST, SUB_RANKS, SUB_CAP, PARTS,
      SUB_LOOT, SUB_LOOT_TIER, GP_TIER_PROC, GP_TIER_PROC_D, GP_ITEM_PROB, SECTOR_LOOT,
      ITEM_DB, ITEM_DB_EN, PRIORITY_LOOT_GROUPS, COMMON_LOOT, GP_PRESETS, GP_PRESET_HINTS, CRAFT_PARTS, CRAFT_MATS, CRAFT_BASE };

    const now = new Date(); const pad = n => String(n).padStart(2, '0');
    const version = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}`;
    const files = {};
    for (const [fn, [title, keys]] of Object.entries(FILES)) {
      const body = keys.map(k => `  ${JSON.stringify(k)}: ${JSON.stringify(data[k])}`).join(',\n');
      files[fn] = `/* FF14 潛水艇計時器 — ${title}\n * 由 tools/excel-to-data.html 從「FF14潛水艇數據表.xlsx」產生，也可以直接手動編輯。\n` +
        ` * 每個欄位的值都是標準 JSON；各欄位說明請見數據總表的「說明」分頁。 */\n` +
        `window.SUB_DATA = Object.assign(window.SUB_DATA || {}, {\n${body}\n});\n`;
    }
    const summary = {
      海域: Object.keys(MAP_NAMES).length, 地點: SUB_SECTORS.length, 航行距離: { 由座標計算: computed, 取自分頁: fromSheet, 缺少: missing },
      等級: Object.keys(SUB_RANKS).length, 部件: Object.values(PARTS).reduce((a, b) => a + b.length, 0), 物品: Object.keys(ITEM_DB).length,
      路線用掉落: Object.values(SUB_LOOT_TIER).reduce((a, t) => a + t['1'].length + t['2'].length + t['3'].length, 0),
      製作部件: Object.keys(CRAFT_PARTS).length,
    };
    files['version.js'] = `/* 數據版本號：由 tools/excel-to-data.html 自動產生，每次更新數據時會一起換掉。 */\nwindow.SUB_DATA_VERSION = '${version}';\n`;
    return { files, warnings, summary, version, data };
  }

  if (typeof module !== 'undefined' && module.exports) module.exports = { convertWorkbook };
  else root.convertWorkbook = convertWorkbook;
})(typeof self !== 'undefined' ? self : this);
