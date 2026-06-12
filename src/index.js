import * as XLSX from 'xlsx';

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

function toInt(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? Math.trunc(n) : fallback;
}

function cleanBlank(value) {
  if (value === null || value === undefined) return null;
  const text = String(value).trim();
  return text === '' ? null : text;
}

function asBlank(value) {
  return value === null || value === undefined || value === '' ? '' : value;
}

function coerceArray(value) {
  return Array.isArray(value) ? value : [];
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

function normalizeHeatCapacityKey(value) {
  const text = normalizeText(value).replace(/,/g, '');
  if (!text) return '';
  const direct = Number(text);
  if (Number.isFinite(direct)) return String(direct);
  const match = text.match(/([\d,.]+)/);
  if (!match) return '';
  const extracted = Number(match[1].replace(/,/g, ''));
  return Number.isFinite(extracted) ? String(extracted) : '';
}

function normalizeTonnage(value) {
  return String(value ?? '').trim().replace(/\.0$/, '');
}

function normalizeNumericText(value) {
  const text = normalizeText(value).replace(/,/g, '');
  if (!text) return '';
  const n = Number(text);
  return Number.isFinite(n) ? String(n) : text;
}

function hasMeaningfulValue(value) {
  const text = normalizeNumericText(value);
  if (!text) return false;
  const n = Number(text);
  if (Number.isFinite(n)) return n !== 0;
  return true;
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
  const econCode =
    unit.economizer === 'barometric'
      ? 'ECO-BARO'
      : unit.economizer === 'powered'
      ? 'ECO-PE'
      : 'NOECO';
  return `${familyCode}-${efficiencyCode}-${tonnageCode}-${voltageCode}-${heatCode}-${reheatCode}-${econCode}`;
}

function optionSummary(unit, match) {
  return (
    normalizeText(unit.remarks) ||
    normalizeText(match?.remarks_default) ||
    [
      unit.hotGasReheat ? 'Hot Gas Reheat' : null,
      unit.economizer === 'barometric' ? 'Economizer w/ Barometric Relief' : null,
      unit.economizer === 'powered' ? 'Economizer w/ Powered Exhaust' : null,
      unit.curb ? 'Curb' : null,
    ]
      .filter(Boolean)
      .join(', ')
  );
}

function xlsxCellValue(sheet, col, row) {
  const addr = XLSX.utils.encode_cell({ c: col - 1, r: row - 1 });
  const cell = sheet[addr];
  if (!cell) return '';
  const v = cell.w !== undefined ? cell.w : cell.v;
  return normalizeText(v == null ? '' : String(v));
}

function xlsxRowCount(sheet) {
  const ref = sheet['!ref'];
  if (!ref) return 0;
  return XLSX.utils.decode_range(ref).e.r + 1;
}

function xlsxColCount(sheet) {
  const ref = sheet['!ref'];
  if (!ref) return 0;
  return XLSX.utils.decode_range(ref).e.c + 1;
}

function slugHeader(text) {
  return normalizeText(text)
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/#/g, ' number ')
    .replace(/\//g, ' ')
    .replace(/[.\-()]+/g, ' ')
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .replace(/_+/g, '_');
}

function fillMergedHeaderRow(values) {
  const filled = [];
  let current = '';
  for (const value of values) {
    const text = normalizeText(value);
    if (text) current = text;
    filled.push(current);
  }
  return filled;
}

function uniqueStrings(values) {
  const out = [];
  const seen = new Set();
  for (const value of values) {
    const text = normalizeText(value);
    if (!text) continue;
    const key = text.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(text);
  }
  return out;
}

function buildHeaderLayers(sheet, headerRows) {
  const maxCol = xlsxColCount(sheet);
  return headerRows.map((rowNumber) => {
    const raw = [];
    for (let col = 1; col <= maxCol; col++) raw.push(xlsxCellValue(sheet, col, rowNumber));
    return fillMergedHeaderRow(raw);
  });
}

function buildRawHeaderMap(sheet, headerRows) {
  const layers = buildHeaderLayers(sheet, headerRows);
  const maxCol = xlsxColCount(sheet);
  const byKey = {};
  const byCol = {};
  for (let col = 1; col <= maxCol; col++) {
    const parts = uniqueStrings(layers.map((layer) => layer[col - 1]));
    const key = slugHeader(parts.join(' '));
    if (!key) continue;
    byCol[col] = { key, parts };
    if (!(key in byKey)) byKey[key] = col;
  }
  return { byKey, byCol, headerRows };
}

function findColumnByAliases(rawHeaderMap, aliases) {
  for (const alias of aliases) {
    const key = slugHeader(alias);
    if (rawHeaderMap.byKey[key]) return rawHeaderMap.byKey[key];
  }
  const entries = Object.entries(rawHeaderMap.byKey);
  for (const alias of aliases) {
    const aliasKey = slugHeader(alias);
    const found = entries.find(([key]) => key.includes(aliasKey) || aliasKey.includes(key));
    if (found) return found[1];
  }
  return null;
}

// Template fingerprints + fixed column maps for known workbook formats.
// Column numbers are 1-based. Add new templates here as new formats arrive.
const TEMPLATE_DEFINITIONS = [
  {
    id: 'DS_COMMERCIAL_SOURCE_WIDE',
    description: 'DS Commercial source schedule (~60 cols, data row 7)',
    fingerprint: [
      { col: 1, row: 1, contains: 'DS Commercial Schedule' },
      { col: 5, row: 5, contains: 'Supply Air Blower' },
      { col: 14, row: 5, contains: 'Cooling' },
      { col: 34, row: 5, contains: 'Heating' },
    ],
    firstDataRow: 7,
    columns: {
      descriptor:            1,
      model_number:          2,
      brand:                 3,
      qty:                   4,
      airflow_cfm:           5,
      supply_fan_esp_in_wg:  8,
      supply_fan_hp:         7,
      supply_fan_rpm:        9,
      cooling_sensible_mbh:  16,   // Cooling section C14 + 2
      cooling_total_mbh:     17,   // Cooling section C14 + 3
      unit_eer:              22,
      seer_ieer:             23,
      gas_heat_input_mbh:    37,   // Heating section C34 + 3
      electric_heat_kw:      37,
      heatpump_capacity_mbh: 37,
      gas_heat_output_mbh:   38,   // Heating section C34 + 4
      voltage:               55,
      mca:                   56,
      mocp:                  58,
      weight_lbs:            60,   // Near end of row, after electrical
      remarks:               61,
    },
  },
  {
    id: 'DS_COMMERCIAL_V1',
    description: 'DS Commercial export template (~28 cols)',
    // Each fingerprint entry must match for the template to be selected.
    fingerprint: [
      { col: 2,  row: 1, contains: 'IDENTIFICATION' },
      { col: 18, row: 2, contains: 'SENSIBLE' },
      { col: 19, row: 2, contains: 'TOTAL' },
      { col: 23, row: 2, contains: 'INPUT' },
      { col: 25, row: 2, contains: 'VOLT' },
    ],
    firstDataRow: 4,
    columns: {
      descriptor:            2,
      model_number:          5,
      brand:                 4,
      qty:                   12,
      airflow_cfm:           10,
      supply_fan_esp_in_wg:  11,
      supply_fan_hp:         14,
      supply_fan_rpm:        15,
      cooling_sensible_mbh:  18,
      cooling_total_mbh:     19,
      heating_cfm:           20,
      // INPUT CAPACITY col handles gas MBH, electric kW, and HP MBH —
      // deriveHeatFields disambiguates by the actual values present
      gas_heat_input_mbh:    23,
      electric_heat_kw:      23,
      heatpump_capacity_mbh: 23,
      gas_heat_output_mbh:   24,
      unit_eer:              8,
      seer_ieer:             9,
      voltage:               25,
      mca:                   26,
      mocp:                  27,
      weight_lbs:            28,
      remarks:               29,
    },
  },
  {
    id: 'DS_COMMERCIAL_FULL_V1',
    description: 'DS Commercial full source template (~44 cols)',
    fingerprint: [
      { col: 2,  row: 1, contains: 'IDENTIFICATION' },
      { col: 21, row: 2, contains: 'SENSIBLE' },
      { col: 22, row: 2, contains: 'TOTAL' },
      { col: 26, row: 2, contains: 'CAPACITY' },
      { col: 39, row: 2, contains: 'VOLT' },
    ],
    firstDataRow: 4,
    columns: {
      descriptor:            2,
      model_number:          5,
      brand:                 4,
      qty:                   15,
      airflow_cfm:           10,
      supply_fan_esp_in_wg:  13,
      supply_fan_hp:         17,
      supply_fan_rpm:        18,
      cooling_sensible_mbh:  21,
      cooling_total_mbh:     22,
      heating_cfm:           23,
      gas_heat_input_mbh:    26,
      electric_heat_kw:      26,
      heatpump_capacity_mbh: 26,
      gas_heat_output_mbh:   27,
      unit_eer:              8,
      seer_ieer:             9,
      voltage:               39,
      mca:                   40,
      mocp:                  41,
      weight_lbs:            43,
      remarks:               44,
    },
  },
];

function detectTemplate(sheet) {
  for (const tmpl of TEMPLATE_DEFINITIONS) {
    const matches = tmpl.fingerprint.every(({ col, row, contains }) =>
      xlsxCellValue(sheet, col, row).toUpperCase().includes(contains.toUpperCase())
    );
    if (matches) return tmpl;
  }
  return null;
}

function buildColumnMap(sheet) {
  const tmpl = detectTemplate(sheet);

  if (tmpl) {
    return {
      columns: { ...tmpl.columns },
      rawHeaderMap: { headerRows: [tmpl.firstDataRow - 1], byKey: {}, byCol: {} },
      missing: [],
      template_id: tmpl.id,
      firstDataRow: tmpl.firstDataRow,
    };
  }

  // ── Fallback: legacy dynamic alias detection for unrecognized formats ──────
  const headerRows = detectBestHeaderRows_legacy(sheet);
  const rawHeaderMap = buildRawHeaderMap(sheet, headerRows);
  const fields = {
    descriptor:            ['tag', 'tag number'],
    model_number:          ['model number'],
    brand:                 ['brand'],
    qty:                   ['qty', 'quantity'],
    airflow_cfm:           ['supply air blower airflow cfm', 'airflow cfm', 'supply cfm', 'cfm'],
    supply_fan_hp:         ['supply air blower hp', 'blower hp', 'fan hp', 'hp ea'],
    supply_fan_esp_in_wg:  ['supply air blower esp iwg', 'esp iwg', 'esp in wg', 'esp'],
    supply_fan_rpm:        ['supply air blower blwr rpm', 'blwr rpm', 'blower rpm', 'rpm'],
    cooling_total_mbh:     ['cooling coil total capacity mbh', 'cooling total mbh', 'total capacity mbh', 'total mbh', 'cooling total'],
    cooling_sensible_mbh:  ['cooling coil sensible capacity mbh', 'cooling sensible mbh', 'sensible capacity mbh', 'sensible mbh'],
    unit_eer:              ['cooling eer', 'unit eer', 'eer'],
    seer_ieer:             ['cooling seer ieer', 'seer ieer', 'ieer', 'seer'],
    gas_heat_input_mbh:    ['heating coil input capacity mbtuh', 'input capacity mbtuh', 'gas heat mbh', 'heating input mbh', 'gas input'],
    gas_heat_output_mbh:   ['heating coil output capacity', 'output capacity', 'gas heat out', 'heating output mbh'],
    electric_heat_kw:      ['electric heater kw', 'electric heat kw', 'electric kw', 'heater kw', 'elec kw'],
    heatpump_capacity_mbh: ['heat pump ratings mbh', 'heat pump capacity mbh', 'heat pump mbh', 'hp mbh'],
    voltage:               ['electrical voltage', 'voltage', 'volt ph', 'volt'],
    mca:                   ['electrical mca', 'mca', 'min circuit amps'],
    mocp:                  ['electrical max fuse', 'max fuse', 'mocp', 'max ocpd'],
    weight_lbs:            ['operating weight lbs', 'weight lbs', 'weight', 'oper wt lbs', 'wt lbs'],
    remarks:               ['remarks'],
  };
  const columns = {};
  for (const [field, aliases] of Object.entries(fields)) {
    columns[field] = findColumnByAliases(rawHeaderMap, aliases);
  }
  const required = ['descriptor', 'model_number', 'brand', 'qty', 'voltage', 'mca'];
  const missing = required.filter((f) => !columns[f]);
  return {
    columns,
    rawHeaderMap,
    missing,
    template_id: 'UNKNOWN_DYNAMIC_FALLBACK',
    firstDataRow: Math.max(...headerRows) + 1,
  };
}

// Renamed to avoid conflict — called only by fallback path above
function detectBestHeaderRows_legacy(sheet) {
  const candidates = [
    [3, 4, 5], [4, 5, 6], [5, 6, 7], [6, 7, 8],
    [7, 8, 9], [8, 9, 10], [9, 10, 11], [10, 11, 12],
  ];
  let best = null;
  for (const headerRows of candidates) {
    const rawHeaderMap = buildRawHeaderMap(sheet, headerRows);
    const probes = {
      descriptor:   ['tag', 'tag number'],
      model_number: ['model number'],
      brand:        ['brand'],
      qty:          ['qty', 'quantity'],
      voltage:      ['electrical voltage', 'voltage', 'volt ph'],
      mca:          ['electrical mca', 'mca'],
      weight_lbs:   ['operating weight lbs', 'weight lbs', 'weight'],
    };
    let score = 0;
    for (const aliases of Object.values(probes)) {
      if (findColumnByAliases(rawHeaderMap, aliases)) score++;
    }
    if (!best || score > best.score) best = { headerRows, score };
    if (best.score === 7) break;
  }
  return best?.headerRows || [6, 7, 8];
}


function getCell(sheet, columnMap, field, rowNumber) {
  const col = columnMap.columns[field];
  return col ? xlsxCellValue(sheet, col, rowNumber) : '';
}

function parseDescriptorBasics(descriptor) {
  const text = normalizeText(descriptor);
  const tonnageMatch = text.match(/^(\d+(?:\.\d+)?)\s*-?\s*ton/i);
  const voltageMatch = text.match(/\b(2083|208-3|208\/3|208-3-60|4603|460-3|460\/3|460-3-60)\b/i);
  const tonnage = tonnageMatch ? normalizeTonnage(tonnageMatch[1]) : '';
  const voltage = normalizeVoltage(voltageMatch ? voltageMatch[1] : '');
  const familyLabel = /\bHP\b|\bHeat Pump\b/i.test(text) ? 'Heat Pump' : 'AC';
  return {
    family_key: familyLabel === 'Heat Pump' ? 'hp' : 'ac',
    family_label: familyLabel,
    efficiency_key: 'Standard',
    efficiency_label: 'Standard',
    tonnage_key: tonnage,
    tonnage_value: tonnage ? Number(tonnage) : 0,
    voltage_key: voltage,
    voltage_label: voltage,
  };
}

function deriveHeatFields(gasInputMbh, gasOutputMbh, electricKw) {
  if (hasMeaningfulValue(gasInputMbh) || hasMeaningfulValue(gasOutputMbh)) {
    const cap = normalizeHeatCapacityKey(gasInputMbh) || normalizeHeatCapacityKey(gasOutputMbh);
    return {
      aux_heat_type_key: 'gas',
      aux_heat_type_label: 'Aluminum Gas Heat',
      aux_heat_capacity_key: cap,
      aux_heat_capacity_label: cap,
    };
  }
  if (hasMeaningfulValue(electricKw)) {
    const cap = normalizeHeatCapacityKey(electricKw);
    return {
      aux_heat_type_key: 'electric',
      aux_heat_type_label: 'Electric Heat',
      aux_heat_capacity_key: cap,
      aux_heat_capacity_label: cap,
    };
  }
  return {
    aux_heat_type_key: 'none',
    aux_heat_type_label: 'None',
    aux_heat_capacity_key: '',
    aux_heat_capacity_label: '',
  };
}

async function loadWorkbookFromR2(env, filename) {
  const obj = await env.TEMPLATES.get(filename);
  if (!obj) throw new Error(`Workbook not found in R2: ${filename}`);
  const buffer = await obj.arrayBuffer();
  return XLSX.read(buffer, { type: 'array', cellText: true, cellDates: true });
}

async function handleUploadTemplate(request, env) {
  const ct = request.headers.get('Content-Type') || '';
  if (!ct.toLowerCase().includes('multipart/form-data'))
    return json({ error: 'Expected multipart/form-data upload.' }, 400);
  const formData = await request.formData();
  const file = formData.get('file');
  if (!file || typeof file === 'string') return json({ error: 'Missing file field.' }, 400);
  const key = normalizeText(file.name) || `upload-${Date.now()}.xlsx`;
  const buffer = await file.arrayBuffer();
  await env.TEMPLATES.put(key, buffer, {
    httpMetadata: {
      contentType: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
  });
  return json({ ok: true, key });
}

async function insertBatch(env, payload) {
  const r = await env.DB.prepare(
    `INSERT INTO import_batches (source_filename, source_sheet, vendor, product_line, notes) VALUES (?, ?, ?, ?, ?)`
  )
    .bind(
      payload.source_filename,
      payload.source_sheet || 'Schedule',
      payload.vendor || null,
      payload.product_line || null,
      payload.notes || null
    )
    .run();
  return r.meta.last_row_id;
}

async function insertStagingRow(env, batchId, d) {
  const r = await env.DB.prepare(
    `INSERT INTO staging_schedule_rows (
      batch_id, source_row_number, source_descriptor, raw_model_number, raw_brand, raw_qty,
      raw_airflow_cfm, raw_supply_fan_hp, raw_supply_fan_esp_in_wg, raw_supply_fan_rpm,
      raw_cooling_total_mbh, raw_cooling_sensible_mbh, raw_unit_eer, raw_seer_ieer,
      raw_refrigerant, raw_heating_input_mbh, raw_heating_output_mbh, raw_voltage,
      raw_mca, raw_mocp, raw_weight_lbs, raw_remarks,
      family_key, family_label, efficiency_key, efficiency_label,
      tonnage_key, tonnage_value, voltage_key, voltage_label,
      aux_heat_type_key, aux_heat_type_label, aux_heat_capacity_key, aux_heat_capacity_label,
      parse_status, parse_notes
    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`
  )
    .bind(
      batchId,
      d.source_row_number,
      d.source_descriptor,
      d.raw_model_number,
      d.raw_brand,
      d.raw_qty,
      d.raw_airflow_cfm,
      d.raw_supply_fan_hp,
      d.raw_supply_fan_esp_in_wg,
      d.raw_supply_fan_rpm,
      d.raw_cooling_total_mbh,
      d.raw_cooling_sensible_mbh,
      d.raw_unit_eer,
      d.raw_seer_ieer,
      d.raw_refrigerant,
      d.raw_heating_input_mbh,
      d.raw_heating_output_mbh,
      d.raw_voltage,
      d.raw_mca,
      d.raw_mocp,
      d.raw_weight_lbs,
      d.raw_remarks,
      d.family_key,
      d.family_label,
      d.efficiency_key,
      d.efficiency_label,
      d.tonnage_key,
      d.tonnage_value,
      d.voltage_key,
      d.voltage_label,
      d.aux_heat_type_key,
      d.aux_heat_type_label,
      d.aux_heat_capacity_key,
      d.aux_heat_capacity_label,
      d.parse_status,
      d.parse_notes
    )
    .run();
  return r.meta.last_row_id;
}

async function insertImportModelResult(env, batchId, stagingRowId, modelNumber, unitModelId, action, reason = null) {
  await env.DB.prepare(
    `INSERT INTO import_model_results (batch_id, staging_row_id, model_number, unit_model_id, action, reason) VALUES (?, ?, ?, ?, ?, ?)`
  )
    .bind(batchId, stagingRowId, modelNumber || null, unitModelId || null, action, reason)
    .run();
}

async function upsertUnitModelV2(env, batchId, d) {
  const modelNumber = d.raw_model_number || '';
  const efficiencyKey = normalizeEfficiency(d.efficiency_key || d.efficiency_label || 'Standard');
  const existing = await env.DB.prepare(`SELECT * FROM unit_models_v2 WHERE model_number = ?`)
    .bind(modelNumber)
    .first();
  if (!existing) {
    const ins = await env.DB.prepare(
      `INSERT INTO unit_models_v2 (
        model_number, family_key, family_label, tonnage_key, tonnage_value,
        voltage_key, voltage_label, aux_heat_type_key, aux_heat_type_label,
        aux_heat_capacity_key, aux_heat_capacity_label, efficiency_key, efficiency_label, source_batch_id
      ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)`
    )
      .bind(
        modelNumber,
        d.family_key,
        d.family_label,
        d.tonnage_key,
        d.tonnage_value,
        d.voltage_key,
        d.voltage_label,
        d.aux_heat_type_key,
        d.aux_heat_type_label,
        d.aux_heat_capacity_key,
        d.aux_heat_capacity_label,
        efficiencyKey,
        d.efficiency_label,
        batchId
      )
      .run();
    return { action: 'inserted', unitModelId: ins.meta.last_row_id };
  }
  const changed = [
    String(existing.family_key || '') !== String(d.family_key || ''),
    String(existing.family_label || '') !== String(d.family_label || ''),
    String(existing.tonnage_key || '') !== String(d.tonnage_key || ''),
    Number(existing.tonnage_value ?? 0) !== Number(d.tonnage_value ?? 0),
    String(existing.voltage_key || '') !== String(d.voltage_key || ''),
    String(existing.voltage_label || '') !== String(d.voltage_label || ''),
    String(existing.aux_heat_type_key || '') !== String(d.aux_heat_type_key || ''),
    String(existing.aux_heat_type_label || '') !== String(d.aux_heat_type_label || ''),
    String(existing.aux_heat_capacity_key || '') !== String(d.aux_heat_capacity_key || ''),
    String(existing.aux_heat_capacity_label || '') !== String(d.aux_heat_capacity_label || ''),
    String(existing.efficiency_key || '') !== String(efficiencyKey),
    String(existing.efficiency_label || '') !== String(d.efficiency_label || ''),
  ].some(Boolean);
  if (!changed) return { action: 'unchanged', unitModelId: existing.id };
  await env.DB.prepare(
    `UPDATE unit_models_v2
      SET family_key=?, family_label=?, tonnage_key=?, tonnage_value=?,
          voltage_key=?, voltage_label=?, aux_heat_type_key=?, aux_heat_type_label=?,
          aux_heat_capacity_key=?, aux_heat_capacity_label=?,
          efficiency_key=?, efficiency_label=?, source_batch_id=?, updated_at=CURRENT_TIMESTAMP
      WHERE id=?`
  )
    .bind(
      d.family_key, d.family_label, d.tonnage_key, d.tonnage_value,
      d.voltage_key, d.voltage_label, d.aux_heat_type_key, d.aux_heat_type_label,
      d.aux_heat_capacity_key, d.aux_heat_capacity_label,
      efficiencyKey, d.efficiency_label, batchId, existing.id
    )
    .run();
  return { action: 'updated', unitModelId: existing.id };
}

async function stageDsCommercialWorkbook(env, payload) {
  const workbook = await loadWorkbookFromR2(env, payload.source_filename);
  const sheetName = payload.source_sheet || workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) throw new Error(`Worksheet not found. Available: ${workbook.SheetNames.join(', ')}`);
  const batchId = await insertBatch(env, payload);
  const columnMap = buildColumnMap(sheet);
  if (columnMap.missing.length)
    throw new Error(`Could not identify required columns: ${columnMap.missing.join(', ')}`);
  const firstDataRow = columnMap.firstDataRow;
  const totalRows = xlsxRowCount(sheet);
  const stagedRows = [];
  for (let rowNumber = firstDataRow; rowNumber <= totalRows; rowNumber++) {
    const descriptor = getCell(sheet, columnMap, 'descriptor', rowNumber);
    const modelNumber = getCell(sheet, columnMap, 'model_number', rowNumber);
    const brand = getCell(sheet, columnMap, 'brand', rowNumber);
    const qty = getCell(sheet, columnMap, 'qty', rowNumber);
    if (!descriptor || /^tag\b/i.test(descriptor)) continue;
    if (/^model number$/i.test(modelNumber)) continue;
    if (!/^\d+(?:\.\d+)?\s*-?\s*ton/i.test(descriptor)) continue;
    if (!modelNumber) continue;
    const rawGasInput = getCell(sheet, columnMap, 'gas_heat_input_mbh', rowNumber);
    const rawGasOutput = getCell(sheet, columnMap, 'gas_heat_output_mbh', rowNumber);
    const rawElecKw = getCell(sheet, columnMap, 'electric_heat_kw', rowNumber);
    const rawHpMbh = getCell(sheet, columnMap, 'heatpump_capacity_mbh', rowNumber);
    const rawVoltage = getCell(sheet, columnMap, 'voltage', rowNumber);
    const descFields = parseDescriptorBasics(descriptor);
    const heatFields = deriveHeatFields(rawGasInput, rawGasOutput, rawElecKw);

    // FIX: if voltage_label came out blank from the descriptor, use the raw_voltage cell value
    if (!descFields.voltage_label && rawVoltage) {
      const normalized = normalizeVoltage(rawVoltage);
      descFields.voltage_key = normalized;
      descFields.voltage_label = normalized;
    }

    const rowData = {
      source_row_number: rowNumber,
      source_descriptor: descriptor,
      raw_model_number: modelNumber,
      raw_brand: brand,
      raw_qty: qty,
      raw_airflow_cfm: getCell(sheet, columnMap, 'airflow_cfm', rowNumber),
      raw_supply_fan_hp: getCell(sheet, columnMap, 'supply_fan_hp', rowNumber),
      raw_supply_fan_esp_in_wg: getCell(sheet, columnMap, 'supply_fan_esp_in_wg', rowNumber),
      raw_supply_fan_rpm: getCell(sheet, columnMap, 'supply_fan_rpm', rowNumber),
      raw_cooling_total_mbh: getCell(sheet, columnMap, 'cooling_total_mbh', rowNumber),
      raw_cooling_sensible_mbh: getCell(sheet, columnMap, 'cooling_sensible_mbh', rowNumber),
      raw_unit_eer: getCell(sheet, columnMap, 'unit_eer', rowNumber),
      raw_seer_ieer: getCell(sheet, columnMap, 'seer_ieer', rowNumber),
      raw_refrigerant: getCell(sheet, columnMap, 'refrigerant', rowNumber),
      raw_heating_input_mbh: hasMeaningfulValue(rawHpMbh) ? rawHpMbh : (rawGasInput || rawElecKw),
      raw_heating_output_mbh: hasMeaningfulValue(rawHpMbh) ? '' : (rawGasOutput || rawGasInput),
      raw_voltage: rawVoltage,
      raw_mca: getCell(sheet, columnMap, 'mca', rowNumber),
      raw_mocp: getCell(sheet, columnMap, 'mocp', rowNumber),
      raw_weight_lbs: getCell(sheet, columnMap, 'weight_lbs', rowNumber),
      raw_remarks: getCell(sheet, columnMap, 'remarks', rowNumber),
      ...descFields,
      ...heatFields,
      parse_status: 'parsed',
      parse_notes: null,
    };
    if (hasMeaningfulValue(rawHpMbh)) {
      rowData.family_key = 'hp';
      rowData.family_label = 'Heat Pump';
      if (
        !hasMeaningfulValue(rawGasInput) &&
        !hasMeaningfulValue(rawGasOutput) &&
        !hasMeaningfulValue(rawElecKw)
      ) {
        rowData.aux_heat_type_key = 'none';
        rowData.aux_heat_type_label = 'None';
        rowData.aux_heat_capacity_key = '';
        rowData.aux_heat_capacity_label = '';
      }
    }
    stagedRows.push(rowData);
  }
  const dupCounts = new Map();
  for (const row of stagedRows)
    dupCounts.set(row.raw_model_number, (dupCounts.get(row.raw_model_number) || 0) + 1);
  const seen = new Set();
  for (const row of stagedRows) {
    const stagingRowId = await insertStagingRow(env, batchId, row);
    row.id = stagingRowId;
    const count = dupCounts.get(row.raw_model_number) || 0;
    if (count > 1 && seen.has(row.raw_model_number)) {
      await insertImportModelResult(
        env, batchId, stagingRowId, row.raw_model_number, null,
        'duplicate_in_batch', 'Duplicate model number within import batch'
      );
      continue;
    }
    seen.add(row.raw_model_number);
    const upsert = await upsertUnitModelV2(env, batchId, row);
    await insertImportModelResult(env, batchId, stagingRowId, row.raw_model_number, upsert.unitModelId, upsert.action, null);
  }
  return batchId;
}

async function getBatch(env, batchId) {
  return (await env.DB.prepare(`SELECT * FROM import_batches WHERE id=?`).bind(batchId).first()) || null;
}

async function resolveActiveBatchId(env, batchId) {
  const id = toInt(batchId, NaN);
  if (Number.isFinite(id) && id > 0) {
    const batch = await getBatch(env, id);
    if (batch) return batch.id;
  }
  const latest = await env.DB.prepare(
    `SELECT id FROM import_batches ORDER BY id DESC LIMIT 1`
  ).first();
  return latest?.id ?? null;
}

async function getBatchSummary(env, batchId) {
  const totals = await env.DB.prepare(
    `SELECT COUNT(*) AS rows_staged,
      SUM(CASE WHEN parse_status='parsed' THEN 1 ELSE 0 END) AS rows_parsed,
      SUM(CASE WHEN parse_status!='parsed' THEN 1 ELSE 0 END) AS rows_failed,
      SUM(CASE WHEN TRIM(COALESCE(parse_notes,''))!='' THEN 1 ELSE 0 END) AS rows_with_warnings,
      COUNT(DISTINCT NULLIF(TRIM(raw_model_number),'')) AS unique_models_in_batch
    FROM staging_schedule_rows WHERE batch_id=?`
  ).bind(batchId).first();
  const actions = await env.DB.prepare(
    `SELECT
      COALESCE(SUM(CASE WHEN action='inserted' THEN 1 ELSE 0 END),0) AS catalog_inserts,
      COALESCE(SUM(CASE WHEN action='updated' THEN 1 ELSE 0 END),0) AS catalog_updates,
      COALESCE(SUM(CASE WHEN action='unchanged' THEN 1 ELSE 0 END),0) AS catalog_unchanged,
      COALESCE(SUM(CASE WHEN action='duplicate_in_batch' THEN 1 ELSE 0 END),0) AS duplicates_in_batch
    FROM import_model_results WHERE batch_id=?`
  ).bind(batchId).first();
  return {
    rows_read: toInt(totals?.rows_staged),
    rows_staged: toInt(totals?.rows_staged),
    rows_parsed: toInt(totals?.rows_parsed),
    rows_failed: toInt(totals?.rows_failed),
    rows_with_warnings: toInt(totals?.rows_with_warnings),
    unique_models_in_batch: toInt(totals?.unique_models_in_batch),
    catalog_inserts: toInt(actions?.catalog_inserts),
    catalog_updates: toInt(actions?.catalog_updates),
    catalog_unchanged: toInt(actions?.catalog_unchanged),
    duplicates_in_batch: toInt(actions?.duplicates_in_batch),
  };
}

async function getBatchIssues(env, batchId) {
  const duplicates = await env.DB.prepare(
    `SELECT raw_model_number AS model_number, COUNT(*) AS duplicate_count,
      GROUP_CONCAT(source_row_number) AS source_row_numbers
    FROM staging_schedule_rows
    WHERE batch_id=? AND NULLIF(TRIM(raw_model_number),'') IS NOT NULL
    GROUP BY raw_model_number HAVING COUNT(*)>1
    ORDER BY duplicate_count DESC, raw_model_number`
  ).bind(batchId).all();
  const parseIssues = await env.DB.prepare(
    `SELECT source_row_number, raw_model_number, parse_status, parse_notes
    FROM staging_schedule_rows
    WHERE batch_id=? AND (parse_status!='parsed' OR TRIM(COALESCE(parse_notes,''))!='')
    ORDER BY source_row_number`
  ).bind(batchId).all();
  const issues = [];
  for (const row of coerceArray(duplicates.results)) {
    issues.push({
      type: 'duplicate_model_in_batch',
      model_number: row.model_number,
      count: toInt(row.duplicate_count),
      source_row_numbers: String(row.source_row_numbers || '')
        .split(',')
        .map((v) => toInt(v))
        .filter(Boolean),
    });
  }
  const warningRows = coerceArray(parseIssues.results).map((row) => ({
    source_row_number: row.source_row_number,
    model_number: row.raw_model_number,
    parse_status: row.parse_status,
    parse_notes: cleanBlank(row.parse_notes),
  }));
  if (warningRows.length) issues.push({ type: 'parse_warnings', rows: warningRows });
  return issues;
}

function buildBatchLinks(batchId) {
  return {
    batch: `/api/import-batches/${batchId}`,
    staging_rows: `/api/import-batches/${batchId}/staging-rows`,
    catalog_results: `/api/import-batches/${batchId}/catalog-results`,
  };
}

async function handleGetImportBatch(env, batchId) {
  const batch = await getBatch(env, batchId);
  if (!batch) return json({ error: 'Import batch not found.' }, 404);
  const summary = await getBatchSummary(env, batchId);
  const issues = await getBatchIssues(env, batchId);
  return json({ ok: true, batch, summary, issues, links: buildBatchLinks(batchId) });
}

async function handleGetImportBatchRows(env, batchId) {
  const batch = await getBatch(env, batchId);
  if (!batch) return json({ error: 'Import batch not found.' }, 404);
  const rows = await env.DB.prepare(
    `SELECT * FROM staging_schedule_rows WHERE batch_id=? ORDER BY source_row_number, id`
  ).bind(batchId).all();
  return json({ ok: true, batch, rows: coerceArray(rows.results) });
}

async function handleGetImportBatchCatalogResults(env, batchId) {
  const batch = await getBatch(env, batchId);
  if (!batch) return json({ error: 'Import batch not found.' }, 404);
  const rows = await env.DB.prepare(
    `SELECT imr.id, imr.batch_id, imr.staging_row_id, ssr.source_row_number, ssr.source_descriptor,
      ssr.raw_model_number, imr.model_number, imr.unit_model_id, imr.action, imr.reason,
      um.family_key, um.tonnage_key, um.voltage_key, um.aux_heat_type_key,
      um.aux_heat_capacity_key, um.efficiency_key, imr.created_at
    FROM import_model_results imr
    JOIN staging_schedule_rows ssr ON ssr.id=imr.staging_row_id
    LEFT JOIN unit_models_v2 um ON um.id=imr.unit_model_id
    WHERE imr.batch_id=? ORDER BY ssr.source_row_number, imr.id`
  ).bind(batchId).all();
  return json({ ok: true, batch, results: coerceArray(rows.results) });
}

async function listCatalog(env, filters) {
  let sql = `SELECT m.*, d.cutsheet_url, d.accessories_url, d.wiring_url, d.iom_url
    FROM unit_models m
    LEFT JOIN unit_documents d ON d.model_id=m.id WHERE 1=1`;
  const binds = [];
  if (filters.family) { sql += ' AND m.family=?'; binds.push(filters.family); }
  if (filters.efficiency) { sql += ' AND m.efficiency=?'; binds.push(filters.efficiency); }
  if (filters.tonnage != null) { sql += ' AND m.tonnage=?'; binds.push(filters.tonnage); }
  if (filters.voltage) { sql += ' AND m.voltage=?'; binds.push(filters.voltage); }
  if (filters.heatType) { sql += ' AND m.heat_type=?'; binds.push(filters.heatType); }
  if (normalizeHeatType(filters.heatType) !== 'None') {
    sql += ' AND m.heat_capacity=?'; binds.push(filters.heatCapacity);
  } else {
    sql += " AND COALESCE(m.heat_capacity,'')=''";
  }
  sql += ' ORDER BY m.model_number';
  return (await env.DB.prepare(sql).bind(...binds).all()).results;
}

async function findMatchingImportedRowInScope(env, unit, batchId = null) {
  const batchClause = batchId ? 'batch_id=? AND ' : '';
  const batchBind = batchId ? [batchId] : [];

  const requestedModel = normalizeText(unit.modelNumber ?? unit.selectedModelNumber ?? '');
  if (requestedModel) {
    const byModel = await env.DB.prepare(
      `SELECT * FROM staging_schedule_rows
       WHERE ${batchClause}raw_model_number=? AND parse_status='parsed'
       ORDER BY id DESC LIMIT 1`
    )
      .bind(...batchBind, requestedModel)
      .first();
    if (byModel) return byModel;
  }

  const family = normalizeFamily(unit.family);
  const tonnage = String(unit.tonnage ?? '');
  const voltage = normalizeVoltage(unit.voltage);
  const heatType = normalizeHeatType(unit.heatType);
  const heatCapKey = heatType === 'None' ? '' : normalizeHeatCapacityKey(unit.heatCapacity);
  const voltageClause = `(voltage_label=? OR TRIM(COALESCE(voltage_label,''))='')`;

  const exact = await env.DB.prepare(
    `SELECT * FROM staging_schedule_rows
     WHERE ${batchClause}family_label=?
       AND tonnage_value=CAST(? AS REAL)
       AND ${voltageClause}
       AND aux_heat_type_label=?
       AND aux_heat_capacity_key=?
       AND parse_status='parsed'
     ORDER BY id DESC LIMIT 1`
  )
    .bind(...batchBind, family, tonnage, voltage, heatType, heatCapKey)
    .first();
  if (exact) return exact;

  const byHeatType = await env.DB.prepare(
    `SELECT * FROM staging_schedule_rows
     WHERE ${batchClause}family_label=?
       AND tonnage_value=CAST(? AS REAL)
       AND ${voltageClause}
       AND aux_heat_type_label=?
       AND parse_status='parsed'
     ORDER BY id DESC LIMIT 1`
  )
    .bind(...batchBind, family, tonnage, voltage, heatType)
    .first();
  if (byHeatType) return byHeatType;

  const relaxed = await env.DB.prepare(
    `SELECT * FROM staging_schedule_rows
     WHERE ${batchClause}family_label=?
       AND tonnage_value=CAST(? AS REAL)
       AND ${voltageClause}
       AND parse_status='parsed'
     ORDER BY id DESC LIMIT 1`
  )
    .bind(...batchBind, family, tonnage, voltage)
    .first();
  if (relaxed) return relaxed;

  const byFamilyTonnage = await env.DB.prepare(
    `SELECT * FROM staging_schedule_rows
     WHERE ${batchClause}family_label=?
       AND tonnage_value=CAST(? AS REAL)
       AND parse_status='parsed'
     ORDER BY id DESC LIMIT 1`
  )
    .bind(...batchBind, family, tonnage)
    .first();
  return byFamilyTonnage ?? null;
}

async function findMatchingImportedRow(env, unit, batchId = null) {
  const preferredBatchId = await resolveActiveBatchId(env, batchId);
  const scopes = preferredBatchId ? [preferredBatchId, null] : [null];

  for (const scopeBatchId of scopes) {
    const match = await findMatchingImportedRowInScope(env, unit, scopeBatchId);
    if (match) return match;
  }
  return null;
}

function buildResolvedScheduleRow(unit, match, index = 0) {
  const unitType = asBlank(match?.family_label ?? unit.family);
  const isHP = normalizeFamily(unitType) === 'Heat Pump';

  const tonnageFromDb = match !== null && match.tonnage_value != null && Number(match.tonnage_value) > 0
  ? match.tonnage_value
  : null;
  

  // FIX: for non-HP units, prefer imported heating input over UI heatCapacity field
  const importedHeatingInput = asBlank(match?.raw_heating_input_mbh);
  const importedHeatingOutput = asBlank(match?.raw_heating_output_mbh);

  return {
    tag: normalizeText(unit.tag) || `RTU-${index + 1}`,
    areaServed: asBlank(unit.areaServed),
    manufacturer: asBlank(match?.raw_brand) || 'Tempmaster',
    modelNumber: asBlank(match?.raw_model_number) || buildSelectionCode(unit),
    nominalTons: tonnageFromDb ?? unit.tonnage,
    unitType,
    unitEer: asBlank(match?.raw_unit_eer),
    seerIeerr: asBlank(match?.raw_seer_ieer),
    supplyCfm: asBlank(match?.raw_airflow_cfm),
    supplyEsp: asBlank(match?.raw_supply_fan_esp_in_wg),
    supplyQty: asBlank(match?.raw_qty) || 1,
    supplyBhp: '',
    supplyHp: asBlank(match?.raw_supply_fan_hp),
    supplyRpm: asBlank(match?.raw_supply_fan_rpm),
    coolingEat: '',
    coolingLat: '',
    coolingSensible: asBlank(match?.raw_cooling_sensible_mbh),
    coolingTotal: asBlank(match?.raw_cooling_total_mbh),
    heatingCfm: asBlank(match?.raw_airflow_cfm),
    heatingEat: '',
    heatingLat: '',
    // FIX: non-HP units — use imported gas/electric input; fall back to UI heatCapacity
    heatingInput: isHP
      ? importedHeatingInput
      : importedHeatingInput || unit.heatCapacity,
    heatingTotalCapacity: isHP
      ? importedHeatingOutput
      : importedHeatingInput,
    heatingOutput: importedHeatingOutput,
    voltPh: asBlank(match?.raw_voltage) || unit.voltage,
    mca: asBlank(match?.raw_mca),
    mocp: asBlank(match?.raw_mocp),
    weight: asBlank(match?.raw_weight_lbs),
    remarks: asBlank(match?.raw_remarks) || optionSummary(unit, match),
    selectionCode: buildSelectionCode(unit),
    matchFound: Boolean(match),
    cutSheetUrl: '',
    accessoriesUrl: '',
    wiringUrl: '',
    iomUrl: '',
  };
}

async function resolveScheduleRows(env, units, batchId = null) {
  const rows = [];
  for (let i = 0; i < units.length; i++) {
    const unit = units[i];
    const match = await findMatchingImportedRow(env, unit, batchId);
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
    cell.t = 's'; cell.v = ''; cell.w = '';
  } else if (typeof value === 'number') {
    cell.t = 'n'; cell.v = value; cell.w = String(value);
  } else {
    cell.t = 's'; cell.v = String(value); cell.w = String(value);
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

async function createWorkbook(env, units, batchId = null) {
  const obj = await env.TEMPLATES.get('SSR-Schedule-Example.xlsx');
  if (!obj) throw new Error('Template workbook not found in R2 bucket.');
  const buf = await obj.arrayBuffer();
  const workbook = XLSX.read(buf, { type: 'array', cellStyles: true, cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const templateRow = 4;
  const rows = await resolveScheduleRows(env, units, batchId);
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
    setCellValue(sheet, COL.heatingInput, rowNum, r.heatingInput);
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

    if (request.method === 'POST' && url.pathname === '/api/upload-template') {
      try { return await handleUploadTemplate(request, env); }
      catch (e) { return json({ error: e.message }, 500); }
    }

    if (request.method === 'POST' && url.pathname === '/api/import-schedule') {
      try {
        const payload = await request.json();
        if (!payload?.source_filename) return json({ error: 'source_filename is required.' }, 400);
        const batchId = await stageDsCommercialWorkbook(env, payload);
        const batch = await getBatch(env, batchId);
        const summary = await getBatchSummary(env, batchId);
        const issues = await getBatchIssues(env, batchId);
        return json({ ok: true, batch, summary, issues, links: buildBatchLinks(batchId) });
      } catch (e) { return json({ error: e.message }, 500); }
    }

    if (request.method === 'POST' && url.pathname === '/api/preview-schedule') {
      try {
        const payload = await request.json();
        const units = Array.isArray(payload?.units) ? payload.units : [];
        const batchId = payload?.batch_id ?? payload?.batchId ?? null;
        const rows = await resolveScheduleRows(env, units, batchId);
        // FIX: was json(rows) — UI reads data.rows so must be wrapped object
        return json({ rows });
      } catch (e) {
        console.error('preview-schedule error', e);
        return json({ error: e.message }, 500);
      }
    }

    const batchMatch = url.pathname.match(/^\/api\/import-batches\/(\d+)$/);
    if (request.method === 'GET' && batchMatch)
      return handleGetImportBatch(env, Number(batchMatch[1]));

    const batchRowsMatch = url.pathname.match(/^\/api\/import-batches\/(\d+)\/staging-rows$/);
    if (request.method === 'GET' && batchRowsMatch)
      return handleGetImportBatchRows(env, Number(batchRowsMatch[1]));

    const batchCatalogMatch = url.pathname.match(/^\/api\/import-batches\/(\d+)\/catalog-results$/);
    if (request.method === 'GET' && batchCatalogMatch)
      return handleGetImportBatchCatalogResults(env, Number(batchCatalogMatch[1]));

    if (request.method === 'POST' && url.pathname === '/api/export-schedule') {
      try {
        const payload = await request.json();
        const units = Array.isArray(payload?.units) ? payload.units : [];
        const batchId = payload?.batch_id ?? payload?.batchId ?? null;
        const file = await createWorkbook(env, units, batchId);
        return new Response(file, {
          headers: {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition': 'attachment; filename="SSR-Schedule-Export.xlsx"',
          },
        });
      } catch (e) { return new Response(e.message || 'Export failed', { status: 500 }); }
    }
    if (request.method === 'GET' && url.pathname === '/api/debug-template') {
      try {
        const filename = url.searchParams.get('file');
        if (!filename) return json({ error: 'file param required — e.g. /api/debug-template?file=your-upload.xlsx' }, 400);
        const obj = await env.TEMPLATES.get(filename);
        if (!obj) return json({ error: `File not found in R2: ${filename}` }, 404);
        const buf = await obj.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array', cellText: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const colMap = buildColumnMap(sheet);
        // Dump first 5 rows so you can see what the fingerprint is checking
        const cellPreview = [];
        for (let row = 1; row <= 5; row++) {
          const rowData = {};
          for (let col = 1; col <= Math.min(xlsxColCount(sheet), 50); col++) {
            const v = xlsxCellValue(sheet, col, row);
            if (v) rowData[`C${col}`] = v;
          }
          cellPreview.push({ row, cells: rowData });
        }
        return json({
          template_id: colMap.template_id,
          columns: colMap.columns,
          firstDataRow: colMap.firstDataRow,
          missing: colMap.missing,
          cell_preview: cellPreview,
        });
      } catch (e) {
        return json({ error: e.message }, 500);
      }
    }
    
    return env.ASSETS.fetch(request);
  },
};
