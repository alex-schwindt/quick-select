import * as XLSX from 'xlsx';

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

function cleanBlank(value) {
  if (value === null || value === undefined) return null;
  const text = String(value).trim();
  return text === '' ? null : text;
}

function asBlank(value) {
  return value === null || value === undefined || value === '' ? '' : value;
}

function normalizeText(value) {
  return String(value ?? '').trim();
}

function normalizeFamily(value) {
  const v = normalizeText(value).toLowerCase();
  if (['ac', 'gas pack', 'gaspack'].includes(v)) return 'AC';
  if (['heat pump', 'heatpump', 'hp'].includes(v)) return 'Heat Pump';
  return normalizeText(value);
}

function normalizeEfficiency(value) {
  const v = normalizeText(value).toLowerCase();
  if (['standard', 'std'].includes(v)) return 'Standard';
  if (['high', 'high efficiency', 'high-efficiency'].includes(v)) return 'High';
  return normalizeText(value);
}

function normalizeVoltage(value) {
  const s = normalizeText(value).toLowerCase().replace(/\s+/g, '').replace(/-/g, '/');
  if (/^208(\/230)?(\/3|\/1)?/.test(s) || s === '208' || s.startsWith('208/')) return '208/230/3';
  if (s === '2083' || s === '20830' || s === '208/3') return '208/230/3';
  if (s === '460' || s === '4603' || s === '460/3' || s.startsWith('460/')) return '460/3';
  return normalizeText(value);
}

function normalizeHeatType(value) {
  const v = normalizeText(value).toLowerCase();
  if (['aluminum gas heat', 'gas heat', 'gas'].includes(v)) return 'Aluminum Gas Heat';
  if (['electric heat', 'electric'].includes(v)) return 'Electric Heat';
  if (['none', 'no heat', ''].includes(v)) return 'None';
  return normalizeText(value);
}

function normalizeHeatCapacity(value) {
  return normalizeText(value);
}

function normalizeNumericText(value) {
  const text = normalizeText(value).replace(/,/g, '');
  if (!text) return '';
  const n = Number(text);
  return Number.isFinite(n) ? String(n) : text;
}

function buildSelectionCode(unit) {
  const familyCode = normalizeFamily(unit.family) === 'Heat Pump' ? 'HP' : 'AC';
  const efficiencyCode = normalizeEfficiency(unit.efficiency) === 'High' ? 'HI' : 'STD';
  const tonnageCode = String(unit.tonnage).replace('.', 'P');
  const voltageCode = normalizeVoltage(unit.voltage) === '460/3' ? '460' : '208';
  const heatTypeKey = normalizeHeatType(unit.heatType);
  const normalizedHeatCapacity = normalizeHeatCapacity(unit.heatCapacity).replace(/\s+/g, '');
  const heatCode =
    heatTypeKey === 'None'
      ? 'NOHEAT'
      : heatTypeKey === 'Electric Heat'
      ? `ELEC-${normalizedHeatCapacity}`
      : `GAS-${normalizedHeatCapacity}`;
  const reheatCode = unit.hotGasReheat ? 'HGRH' : 'NOHGRH';
  const econCode = unit.economizer === 'barometric' ? 'ECO-BARO' : unit.economizer === 'powered' ? 'ECO-PE' : 'NOECO';
  return `${familyCode}-${efficiencyCode}-${tonnageCode}-${voltageCode}-${heatCode}-${reheatCode}-${econCode}`;
}

function optionSummary(unit) {
  return [
    unit.remarks || null,
    unit.hotGasReheat ? 'Hot Gas Reheat' : null,
    unit.economizer === 'barometric' ? 'Economizer w/ Barometric Relief' : null,
    unit.economizer === 'powered' ? 'Economizer w/ Powered Exhaust' : null,
  ]
    .filter(Boolean)
    .join(', ');
}

function parseCsvRows(text) {
  const workbook = XLSX.read(text, { type: 'string' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

function parseWorkbookRows(buffer) {
  const workbook = XLSX.read(buffer, { type: 'array', cellText: true, cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

function mapCatalogUploadRow(raw) {
  const family = normalizeFamily(raw.family || raw.Family);
  const efficiency = normalizeEfficiency(raw.efficiency || raw.Efficiency || 'Standard');
  const tonnage = Number(raw.tonnage || raw.Tonnage || 0);
  const voltage = normalizeVoltage(raw.voltage || raw.Voltage || raw['Volt/Ph']);
  const heatType = normalizeHeatType(raw.heat_type || raw.heatType || raw['Heat Type']);
  const heatCapacity = normalizeHeatCapacity(raw.heat_capacity || raw.heatCapacity || raw['Heat Capacity']);
  const modelNumber = normalizeText(raw.model_number || raw['Model Number']);
  return {
    family,
    efficiency,
    tonnage,
    voltage,
    heat_type: heatType,
    heat_capacity: heatType === 'None' ? '' : heatCapacity,
    model_number: modelNumber,
    cutsheet_url: cleanBlank(raw.cutsheet_url || raw['Cutsheet URL']),
    accessories_url: cleanBlank(raw.accessories_url || raw['Accessories URL']),
    wiring_url: cleanBlank(raw.wiring_url || raw['Wiring URL']),
    iom_url: cleanBlank(raw.iom_url || raw['IOM URL']),
  };
}

function inferCatalogRowFromScheduleLike(raw) {
  const descriptor = normalizeText(raw['Tag #'] || raw.tag || raw.descriptor || '');
  const modelNumber = normalizeText(raw['Model Number'] || raw.model_number || '');
  const voltage = normalizeVoltage(raw['Voltage'] || raw['Volt/Ph'] || raw.voltage || '');
  const tonnageMatch = descriptor.match(/^(\d+(?:\.\d+)?)\s*-?\s*ton/i);
  const tonnage = tonnageMatch ? Number(tonnageMatch[1]) : Number(raw.Tonnage || raw.tonnage || 0);
  const hpMbh = normalizeNumericText(raw['Heat Pump Ratings MBH'] || raw['HP Ratings MBH'] || raw.heatpump_capacity_mbh || '');
  const gasInput = normalizeNumericText(raw['Gas Heat MBH'] || raw['Heating Input MBH'] || raw.gas_heat_input_mbh || '');
  const elecKw = normalizeNumericText(raw['Electric Heater kW'] || raw.electric_heat_kw || '');
  const family = hpMbh ? 'Heat Pump' : 'AC';
  const heatType = gasInput ? 'Aluminum Gas Heat' : elecKw ? 'Electric Heat' : 'None';
  const heatCapacity = gasInput || elecKw || '';
  return {
    family,
    efficiency: normalizeEfficiency(raw.Efficiency || raw.efficiency || 'Standard'),
    tonnage,
    voltage,
    heat_type: heatType,
    heat_capacity: heatCapacity,
    model_number: modelNumber,
    cutsheet_url: cleanBlank(raw['Cutsheet URL'] || raw.cutsheet_url),
    accessories_url: cleanBlank(raw['Accessories URL'] || raw.accessories_url),
    wiring_url: cleanBlank(raw['Wiring URL'] || raw.wiring_url),
    iom_url: cleanBlank(raw['IOM URL'] || raw.iom_url),
  };
}

function looksLikeScheduleStyleRow(raw) {
  const keys = Object.keys(raw || {}).map((k) => String(k).toLowerCase());
  return keys.includes('tag #') || keys.includes('model number') || keys.includes('brand') || keys.includes('qty');
}

async function upsertCatalogRow(env, row) {
  const existing = await env.DB.prepare(`SELECT id FROM unit_models WHERE model_number=? LIMIT 1`).bind(row.model_number).first();
  if (existing?.id) {
    await env.DB.prepare(`UPDATE unit_models SET family=?, efficiency=?, tonnage=?, voltage=?, heat_type=?, heat_capacity=? WHERE id=?`)
      .bind(row.family, row.efficiency, row.tonnage, row.voltage, row.heat_type, row.heat_capacity, existing.id)
      .run();
    const docExisting = await env.DB.prepare(`SELECT id FROM unit_documents WHERE model_id=? LIMIT 1`).bind(existing.id).first();
    if (docExisting?.id) {
      await env.DB.prepare(`UPDATE unit_documents SET cutsheet_url=?, accessories_url=?, wiring_url=?, iom_url=? WHERE model_id=?`)
        .bind(row.cutsheet_url, row.accessories_url, row.wiring_url, row.iom_url, existing.id)
        .run();
    } else {
      await env.DB.prepare(`INSERT INTO unit_documents (model_id, cutsheet_url, accessories_url, wiring_url, iom_url) VALUES (?, ?, ?, ?, ?)`)
        .bind(existing.id, row.cutsheet_url, row.accessories_url, row.wiring_url, row.iom_url)
        .run();
    }
    return 'updated';
  }
  const inserted = await env.DB.prepare(`INSERT INTO unit_models (family, efficiency, tonnage, voltage, heat_type, heat_capacity, model_number) VALUES (?, ?, ?, ?, ?, ?, ?)`)
    .bind(row.family, row.efficiency, row.tonnage, row.voltage, row.heat_type, row.heat_capacity, row.model_number)
    .run();
  await env.DB.prepare(`INSERT INTO unit_documents (model_id, cutsheet_url, accessories_url, wiring_url, iom_url) VALUES (?, ?, ?, ?, ?)`)
    .bind(inserted.meta.last_row_id, row.cutsheet_url, row.accessories_url, row.wiring_url, row.iom_url)
    .run();
  return 'inserted';
}

async function handleAdminImportCatalog(request, env) {
  const ct = request.headers.get('Content-Type') || '';
  if (!ct.toLowerCase().includes('multipart/form-data')) return json({ error: 'Expected multipart/form-data upload.' }, 400);
  const formData = await request.formData();
  const file = formData.get('file');
  if (!file || typeof file === 'string') return json({ error: 'Missing file field.' }, 400);

  const name = normalizeText(file.name).toLowerCase();
  const buffer = await file.arrayBuffer();
  let rows = [];
  if (name.endsWith('.csv')) rows = parseCsvRows(new TextDecoder().decode(buffer));
  else if (name.endsWith('.xlsx') || name.endsWith('.xls')) rows = parseWorkbookRows(buffer);
  else return json({ error: 'Unsupported file type. Use CSV or Excel.' }, 400);

  const issues = [];
  let inserted = 0;
  let updated = 0;
  let rowsRead = 0;

  for (let i = 0; i < rows.length; i += 1) {
    const sourceRow = rows[i] || {};
    const mapped = looksLikeScheduleStyleRow(sourceRow) ? inferCatalogRowFromScheduleLike(sourceRow) : mapCatalogUploadRow(sourceRow);
    if (!mapped.model_number) continue;
    rowsRead += 1;
    if (!mapped.family || !mapped.tonnage || !mapped.voltage || !mapped.model_number || !mapped.heat_type) {
      issues.push(`Row ${i + 2}: missing required catalog fields`);
      continue;
    }
    if (mapped.heat_type !== 'None' && !mapped.heat_capacity) {
      issues.push(`Row ${i + 2}: heat capacity is required when heat type is not None`);
      continue;
    }
    const action = await upsertCatalogRow(env, mapped);
    if (action === 'inserted') inserted += 1;
    if (action === 'updated') updated += 1;
  }

  return json({ ok: issues.length === 0, rows_read: rowsRead, inserted, updated, issues }, issues.length ? 400 : 200);
}

async function listCatalog(env, filters) {
  let sql = `SELECT m.*, d.cutsheet_url, d.accessories_url, d.wiring_url, d.iom_url FROM unit_models m LEFT JOIN unit_documents d ON d.model_id=m.id WHERE 1=1`;
  const binds = [];
  if (filters.family) { sql += ' AND m.family=?'; binds.push(filters.family); }
  if (filters.efficiency) { sql += ' AND m.efficiency=?'; binds.push(filters.efficiency); }
  if (filters.tonnage != null) { sql += ' AND m.tonnage=?'; binds.push(filters.tonnage); }
  if (filters.voltage) { sql += ' AND m.voltage=?'; binds.push(filters.voltage); }
  if (filters.heatType) { sql += ' AND m.heat_type=?'; binds.push(filters.heatType); }
  if (normalizeHeatType(filters.heatType) !== 'None') { sql += ' AND m.heat_capacity=?'; binds.push(filters.heatCapacity); }
  else { sql += " AND COALESCE(m.heat_capacity,'')=''"; }
  sql += ' ORDER BY m.model_number';
  return (await env.DB.prepare(sql).bind(...binds).all()).results;
}

async function findMatchingImportedRow(env, unit) {
  const family = normalizeFamily(unit.family);
  const tonnage = String(unit.tonnage ?? '');
  const voltage = normalizeVoltage(unit.voltage);
  const heatType = normalizeHeatType(unit.heatType);
  const heatCap = normalizeHeatCapacity(unit.heatCapacity);
  const exact = await env.DB.prepare(`SELECT m.*, d.cutsheet_url, d.accessories_url, d.wiring_url, d.iom_url FROM unit_models m LEFT JOIN unit_documents d ON d.model_id=m.id WHERE m.family=? AND CAST(m.tonnage AS REAL)=CAST(? AS REAL) AND m.voltage=? AND m.heat_type=? AND COALESCE(m.heat_capacity,'')=? ORDER BY m.id DESC LIMIT 1`)
    .bind(family, tonnage, voltage, heatType, heatType === 'None' ? '' : heatCap)
    .first();
  if (exact) return exact;
  const relaxed = await env.DB.prepare(`SELECT m.*, d.cutsheet_url, d.accessories_url, d.wiring_url, d.iom_url FROM unit_models m LEFT JOIN unit_documents d ON d.model_id=m.id WHERE m.family=? AND CAST(m.tonnage AS REAL)=CAST(? AS REAL) AND m.voltage=? ORDER BY m.id DESC LIMIT 1`)
    .bind(family, tonnage, voltage)
    .first();
  return relaxed ?? null;
}

function buildResolvedScheduleRow(unit, match, index = 0) {
  return {
    tag: normalizeText(unit.tag) || `RTU-${index + 1}`,
    areaServed: asBlank(unit.areaServed),
    manufacturer: 'Tempmaster',
    modelNumber: asBlank(match?.model_number) || buildSelectionCode(unit),
    nominalTons: asBlank(match?.tonnage) !== '' ? match.tonnage : unit.tonnage,
    unitType: asBlank(match?.family) || unit.family,
    unitEer: '',
    seerIeerr: '',
    supplyCfm: '',
    supplyEsp: '',
    supplyQty: 1,
    supplyBhp: '',
    supplyHp: '',
    supplyRpm: '',
    coolingEat: '',
    coolingLat: '',
    coolingSensible: '',
    coolingTotal: '',
    heatingCfm: '',
    heatingEat: '',
    heatingLat: '',
    heatingInput: unit.heatType === 'None' ? 'No heat' : unit.heatCapacity,
    heatingTotalCapacity: '',
    heatingOutput: '',
    voltPh: asBlank(match?.voltage) || unit.voltage,
    mca: '',
    mocp: '',
    weight: '',
    remarks: optionSummary(unit),
    selectionCode: buildSelectionCode(unit),
    matchFound: Boolean(match),
    cutsheetUrl: asBlank(match?.cutsheet_url),
    accessoriesUrl: asBlank(match?.accessories_url),
    wiringUrl: asBlank(match?.wiring_url),
    iomUrl: asBlank(match?.iom_url),
  };
}

async function resolveScheduleRows(env, units) {
  const rows = [];
  for (let i = 0; i < units.length; i++) {
    const unit = units[i];
    const match = await findMatchingImportedRow(env, unit);
    rows.push(buildResolvedScheduleRow(unit, match, i));
  }
  return rows;
}

const COL = {
  tag: 2, areaServed: 3, manufacturer: 4, modelNumber: 5, nominalTons: 6,
  unitType: 7, unitEer: 8, seerIeerr: 9, supplyCfm: 10, supplyEsp: 11,
  supplyQty: 12, supplyBhp: 13, supplyHp: 14, supplyRpm: 15,
  coolingEat: 16, coolingLat: 17, coolingSensible: 18, coolingTotal: 19,
  heatingCfm: 20, heatingEat: 21, heatingLat: 22, heatingInput: 23,
  heatingOutput: 24, voltPh: 25, mca: 26, mocp: 27, weight: 28, remarks: 29,
};

function setCellValue(sheet, col, row, value) {
  const addr = XLSX.utils.encode_cell({ c: col - 1, r: row - 1 });
  if (!sheet[addr]) sheet[addr] = {};
  const cell = sheet[addr];
  if (value === '' || value == null) {
    cell.t = 's';
    cell.v = '';
    cell.w = '';
  } else if (typeof value === 'number') {
    cell.t = 'n';
    cell.v = value;
    cell.w = String(value);
  } else {
    cell.t = 's';
    cell.v = String(value);
    cell.w = String(value);
  }
}

function expandSheetRange(sheet, maxRow) {
  if (!sheet['!ref']) return;
  const range = XLSX.utils.decode_range(sheet['!ref']);
  if (maxRow - 1 > range.e.r) {
    range.e.r = maxRow - 1;
    sheet['!ref'] = XLSX.utils.encode_range(range);
  }
}

async function createWorkbook(env, units) {
  const obj = await env.TEMPLATES.get('SSR-Schedule-Example.xlsx');
  if (!obj) throw new Error('Template workbook not found in R2 bucket.');
  const buf = await obj.arrayBuffer();
  const workbook = XLSX.read(buf, { type: 'array', cellStyles: true, cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const templateRow = 4;
  const rows = await resolveScheduleRows(env, units);
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const rowNum = templateRow + i;
    setCellValue(sheet, COL.tag, rowNum, r.tag);
    setCellValue(sheet, COL.areaServed, rowNum, r.areaServed);
    setCellValue(sheet, COL.manufacturer, rowNum, r.manufacturer);
    setCellValue(sheet, COL.modelNumber, rowNum, r.modelNumber);
    setCellValue(sheet, COL.nominalTons, rowNum, r.nominalTons);
    setCellValue(sheet, COL.unitType, rowNum, r.unitType);
    setCellValue(sheet, COL.unitEer, rowNum, r.unitEer);
    setCellValue(sheet, COL.seerIeerr, rowNum, r.seerIeerr);
    setCellValue(sheet, COL.supplyCfm, rowNum, r.supplyCfm);
    setCellValue(sheet, COL.supplyEsp, rowNum, r.supplyEsp);
    setCellValue(sheet, COL.supplyQty, rowNum, r.supplyQty);
    setCellValue(sheet, COL.supplyBhp, rowNum, r.supplyBhp);
    setCellValue(sheet, COL.supplyHp, rowNum, r.supplyHp);
    setCellValue(sheet, COL.supplyRpm, rowNum, r.supplyRpm);
    setCellValue(sheet, COL.coolingEat, rowNum, r.coolingEat);
    setCellValue(sheet, COL.coolingLat, rowNum, r.coolingLat);
    setCellValue(sheet, COL.coolingSensible, rowNum, r.coolingSensible);
    setCellValue(sheet, COL.coolingTotal, rowNum, r.coolingTotal);
    setCellValue(sheet, COL.heatingCfm, rowNum, r.heatingCfm);
    setCellValue(sheet, COL.heatingEat, rowNum, r.heatingEat);
    setCellValue(sheet, COL.heatingLat, rowNum, r.heatingLat);
    setCellValue(sheet, COL.heatingInput, rowNum, r.heatingTotalCapacity || r.heatingInput);
    setCellValue(sheet, COL.heatingOutput, rowNum, r.heatingOutput);
    setCellValue(sheet, COL.voltPh, rowNum, r.voltPh);
    setCellValue(sheet, COL.mca, rowNum, r.mca);
    setCellValue(sheet, COL.mocp, rowNum, r.mocp);
    setCellValue(sheet, COL.weight, rowNum, r.weight);
    setCellValue(sheet, COL.remarks, rowNum, r.remarks);
  }
  expandSheetRange(sheet, templateRow + rows.length - 1);
  return XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === 'GET' && url.pathname === '/api/catalog') {
      const f = {
        family: normalizeFamily(url.searchParams.get('family')),
        efficiency: normalizeEfficiency(url.searchParams.get('efficiency')),
        tonnage: url.searchParams.get('tonnage'),
        voltage: normalizeVoltage(url.searchParams.get('voltage')),
        heatType: normalizeHeatType(url.searchParams.get('heatType')),
        heatCapacity: normalizeHeatCapacity(url.searchParams.get('heatCapacity')),
      };
      return json({ items: await listCatalog(env, f) });
    }

    if (request.method === 'POST' && url.pathname === '/api/admin/import-catalog') {
      try {
        return await handleAdminImportCatalog(request, env);
      } catch (e) {
        return json({ error: e.message || 'Catalog import failed.' }, 500);
      }
    }

    if (request.method === 'POST' && url.pathname === '/api/preview-schedule') {
      try {
        const payload = await request.json();
        const units = Array.isArray(payload?.units) ? payload.units : [];
        const rows = await resolveScheduleRows(env, units);
        return json({ rows });
      } catch (e) {
        return json({ error: e.message }, 500);
      }
    }

    if (request.method === 'POST' && url.pathname === '/api/export-schedule') {
      try {
        const payload = await request.json();
        const units = Array.isArray(payload?.units) ? payload.units : [];
        const file = await createWorkbook(env, units);
        return new Response(file, {
          headers: {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition': 'attachment; filename="SSR-Schedule-Export.xlsx"',
          },
        });
      } catch (e) {
        return new Response(e.message || 'Export failed', { status: 500 });
      }
    }

    return env.ASSETS.fetch(request);
  },
};
