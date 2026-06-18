import * as XLSX from 'xlsx';

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json' }
  });
}

function normalizeText(value) {
  return String(value ?? '').trim();
}

function cleanBlank(value) {
  const text = normalizeText(value);
  return text ? text : null;
}

function asBlank(value) {
  return value == null ? '' : String(value);
}

function numberOrNull(value) {
  const n = Number(String(value ?? '').replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : null;
}

function normalizeVoltage(value) {
  const s = normalizeText(value).toLowerCase().replace(/\s+/g, '').replace(/-/g, '/');
  if (!s) return '';
  if (s.includes('460')) return '460/3';
  if (s.includes('208') || s.includes('230')) return '208/230/3';
  return normalizeText(value);
}

function normalizeUnitType(value) {
  const v = normalizeText(value).toLowerCase();
  if (!v) return '';
  if (['ac', 'packaged ac', 'gas pack', 'gas package', 'gaspack'].includes(v)) return 'Packaged AC';
  if (['heat pump', 'packaged heat pump', 'hp', 'packaged hp'].includes(v)) return 'Packaged Heat Pump';
  return normalizeText(value);
}

function normalizeHeatType(value) {
  const v = normalizeText(value).toLowerCase();
  if (!v || v === 'none' || v === 'no heat') return 'None';
  if (v.includes('electric')) return 'Electric Heat';
  if (v.includes('gas')) return 'Aluminum Gas Heat';
  return normalizeText(value);
}

function normalizeHeatCapacity(value) {
  return normalizeText(value);
}

function slugify(value) {
  return String(value ?? '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
}

function buildSelectionCode(unit) {
  const familyCode = normalizeUnitType(unit.family || unit.unitType).toLowerCase().includes('heat pump') ? 'HP' : 'AC';
  const tonnageCode = String(unit.tonnage ?? '').replace('.', 'P');
  const voltageCode = normalizeVoltage(unit.voltage) === '460/3' ? '460' : '208';
  const heatTypeKey = normalizeHeatType(unit.heatType);
  const normalizedHeatCapacity = normalizeHeatCapacity(unit.heatCapacity).replace(/\s+/g, '').replace(/[^0-9A-Za-z.-]/g, '');
  const heatCode = heatTypeKey === 'None'
    ? 'NOHEAT'
    : heatTypeKey === 'Electric Heat'
      ? `ELEC-${normalizedHeatCapacity}`
      : `GAS-${normalizedHeatCapacity}`;
  const reheatCode = unit.hotGasReheat ? 'HGRH' : 'NOHGRH';
  const econCode = unit.economizer === 'barometric' ? 'ECO-BARO' : unit.economizer === 'powered' ? 'ECO-PE' : 'NOECO';
  return `${familyCode}-${tonnageCode}-${voltageCode}-${heatCode}-${reheatCode}-${econCode}`;
}

function optionSummary(unit) {
  return normalizeText(unit.remarks) || [
    unit.hotGasReheat ? 'Hot Gas Reheat' : null,
    unit.economizer === 'barometric' ? 'Economizer w/ Barometric Relief' : null,
    unit.economizer === 'powered' ? 'Economizer w/ Powered Exhaust' : null
  ].filter(Boolean).join(', ');
}

async function ensureDocuments(env, modelId, modelNumber) {
  const existing = await env.DB.prepare('SELECT model_id FROM unit_documents WHERE model_id = ?').bind(modelId).first();
  if (existing) return;
  const slug = slugify(modelNumber);
  await env.DB.prepare(`INSERT INTO unit_documents (model_id, cutsheet_url, accessories_url, wiring_url, iom_url) VALUES (?, ?, ?, ?, ?)`)
    .bind(
      modelId,
      `https://selections.hhtrecho.com/cutsheets/${slug}.pdf`,
      `https://selections.hhtrecho.com/accessories/${slug}.pdf`,
      `https://selections.hhtrecho.com/wiring/${slug}.pdf`,
      `https://selections.hhtrecho.com/iom/${slug}.pdf`
    )
    .run();
}

function csvRowToModel(rec) {
  const modelNumber = normalizeText(rec.Model_number || rec.model_number);
  if (!modelNumber) return null;
  return {
    model_number: modelNumber,
    manufacturer: cleanBlank(rec.Manufacturer || rec.manufacturer),
    unit_type: normalizeUnitType(rec.Unit_Type || rec.unit_type),
    nominal_tonnage: numberOrNull(rec.Nominal_Tonnage || rec.nominal_tonnage),
    cfm: numberOrNull(rec.CFM || rec.cfm),
    hp: numberOrNull(rec.HP || rec.hp),
    esp: numberOrNull(rec.ESP || rec.esp),
    rpm: numberOrNull(rec.RPM || rec.rpm),
    cooling_eat_db: cleanBlank(rec.Cooling_EAT_DB),
    cooling_eat_wb: cleanBlank(rec.Cooling_EAT_WB),
    cooling_lat_db: cleanBlank(rec.Cooling_LAT_DB),
    cooling_lat_wb: cleanBlank(rec.Cooling_LAT_WB),
    cooling_total_capacity: numberOrNull(rec.Cooling_Total_Capacity),
    cooling_sensible_capacity: numberOrNull(rec.Cooling_Sensible_Capacity),
    eer: cleanBlank(rec.EER),
    seer_ieer: cleanBlank(rec['SEER/IEER'] || rec.SEER_IEER || rec.seer_ieer),
    heating_eat: cleanBlank(rec.Heating_EAT),
    heating_lat: cleanBlank(rec.Heating_LAT),
    heating_capacity: cleanBlank(rec.Heating_Capacity),
    heating_gas_input: cleanBlank(rec.Heating_Gas_Input),
    heatpump_total_capacity: cleanBlank(rec['Heatpump_Total Capacity'] || rec.Heatpump_Total_Capacity),
    heat_pump_hspf: cleanBlank(rec.Heat_Pump_HSPF),
    electric_heat_capacity: cleanBlank(rec.Electric_Heat_Capacity),
    voltage: normalizeVoltage(rec.Voltage || rec.voltage),
    mca: cleanBlank(rec.MCA || rec.mca),
    mocp: cleanBlank(rec.MOCP || rec.mocp)
  };
}

async function upsertCatalogRow(env, row) {
  const existing = await env.DB.prepare(`SELECT id FROM unit_models WHERE model_number = ? LIMIT 1`)
    .bind(row.model_number)
    .first();

  if (existing?.id) {
    await env.DB.prepare(`UPDATE unit_models SET manufacturer = ?, unit_type = ?, nominal_tonnage = ?, cfm = ?, hp = ?, esp = ?, rpm = ?, cooling_eat_db = ?, cooling_eat_wb = ?, cooling_lat_db = ?, cooling_lat_wb = ?, cooling_total_capacity = ?, cooling_sensible_capacity = ?, eer = ?, seer_ieer = ?, heating_eat = ?, heating_lat = ?, heating_capacity = ?, heating_gas_input = ?, heatpump_total_capacity = ?, heat_pump_hspf = ?, electric_heat_capacity = ?, voltage = ?, mca = ?, mocp = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?`)
      .bind(
        row.manufacturer,
        row.unit_type,
        row.nominal_tonnage,
        row.cfm,
        row.hp,
        row.esp,
        row.rpm,
        row.cooling_eat_db,
        row.cooling_eat_wb,
        row.cooling_lat_db,
        row.cooling_lat_wb,
        row.cooling_total_capacity,
        row.cooling_sensible_capacity,
        row.eer,
        row.seer_ieer,
        row.heating_eat,
        row.heating_lat,
        row.heating_capacity,
        row.heating_gas_input,
        row.heatpump_total_capacity,
        row.heat_pump_hspf,
        row.electric_heat_capacity,
        row.voltage,
        row.mca,
        row.mocp,
        existing.id
      )
      .run();
    await ensureDocuments(env, existing.id, row.model_number);
    return 'updated';
  }

  const insert = await env.DB.prepare(`INSERT INTO unit_models (model_number, manufacturer, unit_type, nominal_tonnage, cfm, hp, esp, rpm, cooling_eat_db, cooling_eat_wb, cooling_lat_db, cooling_lat_wb, cooling_total_capacity, cooling_sensible_capacity, eer, seer_ieer, heating_eat, heating_lat, heating_capacity, heating_gas_input, heatpump_total_capacity, heat_pump_hspf, electric_heat_capacity, voltage, mca, mocp) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`)
    .bind(
      row.model_number,
      row.manufacturer,
      row.unit_type,
      row.nominal_tonnage,
      row.cfm,
      row.hp,
      row.esp,
      row.rpm,
      row.cooling_eat_db,
      row.cooling_eat_wb,
      row.cooling_lat_db,
      row.cooling_lat_wb,
      row.cooling_total_capacity,
      row.cooling_sensible_capacity,
      row.eer,
      row.seer_ieer,
      row.heating_eat,
      row.heating_lat,
      row.heating_capacity,
      row.heating_gas_input,
      row.heatpump_total_capacity,
      row.heat_pump_hspf,
      row.electric_heat_capacity,
      row.voltage,
      row.mca,
      row.mocp
    )
    .run();

  const modelId = insert.meta?.last_row_id;
  if (modelId) await ensureDocuments(env, modelId, row.model_number);
  return 'inserted';
}

async function handleAdminImportCatalog(request, env) {
  const ct = request.headers.get('Content-Type') || '';
  if (!ct.toLowerCase().includes('multipart/form-data')) return json({ error: 'Expected multipart/form-data upload.' }, 400);
  const formData = await request.formData();
  const file = formData.get('file');
  if (!file || typeof file === 'string') return json({ error: 'Missing file field.' }, 400);

  const name = normalizeText(file.name).toLowerCase();
  if (!name.endsWith('.csv')) return json({ error: 'This fresh import path expects a CSV file.' }, 400);

  const text = await file.text();
  const wb = XLSX.read(text, { type: 'string' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const records = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  let rowsRead = 0;
  let inserted = 0;
  let updated = 0;
  const issues = [];

  for (const rec of records) {
    const row = csvRowToModel(rec);
    if (!row) continue;
    if (!row.unit_type) {
      issues.push(`${row.model_number}: Unit_Type is required.`);
      continue;
    }
    if (!Number.isFinite(row.nominal_tonnage)) {
      issues.push(`${row.model_number}: Nominal_Tonnage is required.`);
      continue;
    }
    if (!row.voltage) {
      issues.push(`${row.model_number}: Voltage is required.`);
      continue;
    }
    rowsRead += 1;
    try {
      const action = await upsertCatalogRow(env, row);
      if (action === 'inserted') inserted += 1;
      if (action === 'updated') updated += 1;
    } catch (error) {
      issues.push(`${row.model_number}: ${error.message}`);
    }
  }

  return json({ ok: true, rows_read: rowsRead, inserted, updated, issues });
}

async function findMatchingImportedRow(env, unit) {
  const requestedType = normalizeUnitType(unit.family || unit.unitType);
  const requestedVoltage = normalizeVoltage(unit.voltage);
  const requestedTonnage = Number(unit.tonnage);

  const exact = await env.DB.prepare(`SELECT um.*, ud.cutsheet_url, ud.accessories_url, ud.wiring_url, ud.iom_url FROM unit_models um LEFT JOIN unit_documents ud ON ud.model_id = um.id WHERE um.unit_type = ? AND um.nominal_tonnage = ? AND um.voltage = ? ORDER BY um.id LIMIT 1`)
    .bind(requestedType, requestedTonnage, requestedVoltage)
    .first();

  if (exact) return exact;

  const relaxed = await env.DB.prepare(`SELECT um.*, ud.cutsheet_url, ud.accessories_url, ud.wiring_url, ud.iom_url FROM unit_models um LEFT JOIN unit_documents ud ON ud.model_id = um.id WHERE um.nominal_tonnage = ? AND um.voltage = ? ORDER BY um.id LIMIT 1`)
    .bind(requestedTonnage, requestedVoltage)
    .first();

  return relaxed || null;
}

function previewFallback(unit) {
  const unitType = normalizeUnitType(unit.family || unit.unitType) || normalizeText(unit.family) || 'Packaged AC';
  return {
    tag: unit.tag,
    areaServed: unit.areaServed || '—',
    manufacturer: 'H&H Trecho',
    modelNumber: buildSelectionCode(unit),
    nominalTons: unit.tonnage,
    unitType,
    unitEer: '—',
    seerIeerr: '—',
    supplyCfm: '—',
    supplyEsp: '—',
    supplyBhp: '—',
    supplyHp: '—',
    supplyRpm: '—',
    coolingEat: '—',
    coolingLat: '—',
    coolingSensible: '—',
    coolingTotal: '—',
    heatingEat: '—',
    heatingLat: '—',
    heatingTotalCapacity: '—',
    heatingInput: unit.heatType === 'None' ? 'No heat' : `${unit.heatType} ${unit.heatCapacity}`,
    voltPh: unit.voltage,
    mca: '—',
    mocp: '—',
    filterType: '—',
    weight: '—',
    remarks: optionSummary(unit),
    matchFound: false,
    cutsheetUrl: '',
    accessoriesUrl: '',
    wiringUrl: '',
    iomUrl: ''
  };
}

function combineDbWb(db, wb) {
  const dbText = normalizeText(db);
  const wbText = normalizeText(wb);
  if (dbText && wbText) return `${dbText} / ${wbText}`;
  return dbText || wbText || '—';
}

function buildHeatingInput(unit, match) {
  if (normalizeText(match?.heating_gas_input)) return normalizeText(match.heating_gas_input);
  if (normalizeText(match?.electric_heat_capacity)) return `Electric Heat ${normalizeText(match.electric_heat_capacity)}`;
  if (unit.heatType === 'None') return 'No heat';
  return `${unit.heatType} ${unit.heatCapacity}`.trim();
}

function buildPreviewRow(unit, match) {
  if (!match) return previewFallback(unit);
  return {
    tag: unit.tag,
    areaServed: unit.areaServed || '—',
    manufacturer: match.manufacturer || 'Tempmaster',
    modelNumber: match.model_number || buildSelectionCode(unit),
    nominalTons: unit.tonnage,
    unitType: match.unit_type || normalizeUnitType(unit.family || unit.unitType),
    unitEer: asBlank(match.eer) || '—',
    seerIeerr: asBlank(match.seer_ieer) || '—',
    supplyCfm: asBlank(match.cfm) || '—',
    supplyEsp: asBlank(match.esp) || '—',
    supplyBhp: '—',
    supplyHp: asBlank(match.hp) || '—',
    supplyRpm: asBlank(match.rpm) || '—',
    coolingEat: combineDbWb(match.cooling_eat_db, match.cooling_eat_wb),
    coolingLat: combineDbWb(match.cooling_lat_db, match.cooling_lat_wb),
    coolingSensible: asBlank(match.cooling_sensible_capacity) || '—',
    coolingTotal: asBlank(match.cooling_total_capacity) || '—',
    heatingEat: asBlank(match.heating_eat) || '—',
    heatingLat: asBlank(match.heating_lat) || '—',
    heatingTotalCapacity: asBlank(match.heating_capacity || match.heatpump_total_capacity) || '—',
    heatingInput: buildHeatingInput(unit, match),
    voltPh: match.voltage || unit.voltage,
    mca: asBlank(match.mca) || '—',
    mocp: asBlank(match.mocp) || '—',
    filterType: '—',
    weight: '—',
    remarks: optionSummary(unit) || '—',
    matchFound: true,
    cutsheetUrl: match.cutsheet_url || '',
    accessoriesUrl: match.accessories_url || '',
    wiringUrl: match.wiring_url || '',
    iomUrl: match.iom_url || ''
  };
}

async function handlePreviewSchedule(request, env) {
  const payload = await request.json();
  const units = Array.isArray(payload?.units) ? payload.units : [];
  const rows = [];
  for (const unit of units) {
    const match = await findMatchingImportedRow(env, unit);
    rows.push(buildPreviewRow(unit, match));
  }
  return json({ ok: true, rows });
}

function setCell(ws, col, row, value) {
  const addr = XLSX.utils.encode_cell({ c: col - 1, r: row - 1 });
  if (value === undefined || value === null || value === '') {
    if (ws[addr]) delete ws[addr];
    return;
  }
  ws[addr] = { t: typeof value === 'number' ? 'n' : 's', v: value };
}

function updateSheetRange(ws, maxCol, maxRow) {
  ws['!ref'] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: maxCol - 1, r: maxRow - 1 } });
}

async function loadTemplateWorkbook(env) {
  const object = await env.TEMPLATES.get('SSR-Schedule-Example.xlsx');
  if (!object) throw new Error('Template workbook not found in R2 bucket.');
  const buffer = await object.arrayBuffer();
  return XLSX.read(buffer, { type: 'array', cellStyles: true, cellNF: true, cellDates: true });
}

async function createWorkbook(env, units) {
  const wb = await loadTemplateWorkbook(env);
  const sheetName = wb.SheetNames.includes('Table 1') ? 'Table 1' : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error('Template worksheet not found.');

  let currentRow = 4;
  for (const unit of units) {
    const match = await findMatchingImportedRow(env, unit);
    const row = buildPreviewRow(unit, match);

    setCell(ws, 2, currentRow, row.tag);
    setCell(ws, 3, currentRow, row.areaServed === '—' ? '' : row.areaServed);
    setCell(ws, 4, currentRow, row.manufacturer);
    setCell(ws, 5, currentRow, row.modelNumber);
    setCell(ws, 6, currentRow, Number(row.nominalTons));
    setCell(ws, 7, currentRow, row.unitType);
    setCell(ws, 8, currentRow, row.unitEer === '—' ? '' : row.unitEer);
    setCell(ws, 9, currentRow, row.seerIeerr === '—' ? '' : row.seerIeerr);
    setCell(ws, 10, currentRow, row.supplyCfm === '—' ? '' : Number(row.supplyCfm));
    setCell(ws, 11, currentRow, row.supplyEsp === '—' ? '' : row.supplyEsp);
    setCell(ws, 12, currentRow, row.supplyBhp === '—' ? '' : row.supplyBhp);
    setCell(ws, 13, currentRow, row.supplyHp === '—' ? '' : row.supplyHp);
    setCell(ws, 14, currentRow, row.supplyRpm === '—' ? '' : row.supplyRpm);
    setCell(ws, 15, currentRow, row.coolingEat === '—' ? '' : row.coolingEat);
    setCell(ws, 16, currentRow, row.coolingLat === '—' ? '' : row.coolingLat);
    setCell(ws, 17, currentRow, row.coolingSensible === '—' ? '' : row.coolingSensible);
    setCell(ws, 18, currentRow, row.coolingTotal === '—' ? '' : row.coolingTotal);
    setCell(ws, 19, currentRow, row.heatingEat === '—' ? '' : row.heatingEat);
    setCell(ws, 20, currentRow, row.heatingLat === '—' ? '' : row.heatingLat);
    setCell(ws, 21, currentRow, row.heatingTotalCapacity === '—' ? '' : row.heatingTotalCapacity);
    setCell(ws, 22, currentRow, row.heatingInput === 'No heat' ? '' : row.heatingInput);
    setCell(ws, 23, currentRow, row.voltPh);
    setCell(ws, 24, currentRow, row.mca === '—' ? '' : row.mca);
    setCell(ws, 25, currentRow, row.mocp === '—' ? '' : row.mocp);
    setCell(ws, 26, currentRow, row.filterType === '—' ? '' : row.filterType);
    setCell(ws, 27, currentRow, row.weight === '—' ? '' : row.weight);
    setCell(ws, 28, currentRow, row.remarks === '—' ? '' : row.remarks);

    currentRow += 1;
  }

  const existingRange = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  updateSheetRange(ws, Math.max(existingRange.e.c + 1, 28), Math.max(existingRange.e.r + 1, currentRow));
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellStyles: true });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === 'POST' && url.pathname === '/api/admin/import-catalog') {
      try {
        return await handleAdminImportCatalog(request, env);
      } catch (error) {
        return json({ error: error.message || 'Catalog import failed.' }, 500);
      }
    }

    if (request.method === 'POST' && url.pathname === '/api/preview-schedule') {
      try {
        return await handlePreviewSchedule(request, env);
      } catch (error) {
        return json({ error: error.message || 'Preview failed.' }, 500);
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
            'Content-Disposition': 'attachment; filename="SSR-Schedule-Export.xlsx"'
          }
        });
      } catch (error) {
        return new Response(error.message || 'Export failed', { status: 500 });
      }
    }

    return env.ASSETS.fetch(request);
  }
};
