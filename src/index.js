import * as XLSX from 'xlsx';

// ─────────────────────────────────────────────
// Utility helpers
// ─────────────────────────────────────────────

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

// ─────────────────────────────────────────────
// Domain normalization
// ─────────────────────────────────────────────

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
  const compact = normalizeText(value).toLowerCase().replace(/\s+/g, '').replace(/v/g, '');
  if (['208/3', '208-3', '2083', '208/230/3', '2082303', '208-3-60'].includes(compact.replace('/230', ''))) return '208/230/3';
  if (['208/230/3', '2082303', '208-3-60'].includes(compact)) return '208/230/3';
  if (['460/3', '460-3', '4603', '460-3-60'].includes(compact)) return '460/3';
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

function normalizeTonnage(value) {
  return String(value ?? '').trim().replace('.0', '');
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

// ─────────────────────────────────────────────
// Schedule helpers
// ─────────────────────────────────────────────

function buildSelectionCode(unit) {
  const familyCode = normalizeFamily(unit.family) === 'Heat Pump' ? 'HP' : 'AC';
  const efficiencyCode = normalizeEfficiency(unit.efficiency) === 'High' ? 'HI' : 'STD';
  const tonnageCode = String(unit.tonnage).replace('.', 'P');
  const voltageCode = normalizeVoltage(unit.voltage) === '460/3' ? '460' : '208';
  const heatTypeKey = normalizeHeatType(unit.heatType);
  const normalizedHeatCapacity = normalizeHeatCapacity(unit.heatCapacity).replace(/\s+/g, '');

  const heatCode = heatTypeKey === 'None'
    ? 'NOHEAT'
    : heatTypeKey === 'Electric Heat'
      ? `ELEC-${normalizedHeatCapacity}`
      : `GAS-${normalizedHeatCapacity}`;

  const reheatCode = unit.hotGasReheat ? 'HGRH' : 'NOHGRH';
  const econCode = unit.economizer === 'barometric'
    ? 'ECO-BARO'
    : unit.economizer === 'powered'
      ? 'ECO-PE'
      : 'NOECO';

  return `${familyCode}-${efficiencyCode}-${tonnageCode}-${voltageCode}-${heatCode}-${reheatCode}-${econCode}`;
}

function optionSummary(unit, match) {
  return normalizeText(unit.remarks) || normalizeText(match?.remarks_default) || [
    unit.hotGasReheat ? 'Hot Gas Reheat' : null,
    unit.economizer === 'barometric' ? 'Economizer w/ Barometric Relief' : null,
    unit.economizer === 'powered' ? 'Economizer w/ Powered Exhaust' : null,
    unit.curb ? 'Curb' : null,
  ].filter(Boolean).join(', ');
}

// ─────────────────────────────────────────────
// SheetJS workbook helpers
// ─────────────────────────────────────────────

/**
 * Load a workbook from R2 using SheetJS.
 * SheetJS works natively in Cloudflare Workers — no Node.js dependencies.
 */
async function loadWorkbookFromR2(env, sourceFilename) {
  const r2Object = await env.TEMPLATES.get(sourceFilename);
  if (!r2Object) throw new Error(`Workbook not found in R2: "${sourceFilename}"`);

  const arrayBuffer = await r2Object.arrayBuffer();
  if (!arrayBuffer || arrayBuffer.byteLength < 512) {
    throw new Error(`R2 object "${sourceFilename}" is empty or corrupt (${arrayBuffer?.byteLength ?? 0} bytes).`);
  }

  return XLSX.read(arrayBuffer, { type: 'array', cellText: true, cellDates: false, raw: false });
}

/**
 * Get a worksheet by name, falling back to the first sheet.
 * Returns { worksheet, name } where worksheet is a SheetJS sheet object.
 */
function getWorksheet(workbook, sheetName) {
  const name = workbook.SheetNames.includes(sheetName)
    ? sheetName
    : workbook.SheetNames[0];

  if (!name) {
    throw new Error('Workbook contains no sheets.');
  }

  return { worksheet: workbook.Sheets[name], name };
}

/**
 * Convert a SheetJS sheet to a 2D array of trimmed strings.
 * Row index is 0-based. Column index is 0-based.
 * Empty cells become ''.
 */
function sheetToRows(worksheet) {
  return XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
    raw: false,
    blankrows: true,
  }).map((row) => row.map((cell) => normalizeText(cell)));
}

/**
 * Read a single cell from the rows array by 0-based row/col.
 */
function getCell(rows, rowIndex, colIndex) {
  return normalizeText(rows[rowIndex]?.[colIndex] ?? '');
}

// ─────────────────────────────────────────────
// Header detection (same deterministic anchor approach, adapted for SheetJS rows)
// ─────────────────────────────────────────────

function slugHeader(text) {
  return normalizeText(text)
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/#/g, ' number ')
    .replace(/\//g, ' ')
    .replace(/[.\/\(\)\-]+/g, ' ')
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .replace(/_+/g, '_');
}

function fillMergedHeaderRow(cells) {
  const filled = [];
  let current = '';
  for (const cell of cells) {
    const text = normalizeText(cell);
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

/**
 * Scan the first 20 rows to find the anchor row containing both "Tag" and "Model Number".
 * Returns 0-based row index, or null if not found.
 */
function findAnchorRowIndex(rows) {
  const maxRow = Math.min(rows.length, 20);
  for (let r = 0; r < maxRow; r += 1) {
    const row = rows[r];
    let hasTag = false;
    let hasModelNumber = false;
    for (const cell of row) {
      const v = normalizeText(cell).toLowerCase();
      if (v === 'tag' || v === 'tag number' || v === 'tag #') hasTag = true;
      if (v === 'model number' || v === 'model no' || v === 'model #') hasModelNumber = true;
    }
    if (hasTag && hasModelNumber) return r;
  }
  return null;
}

function buildRawHeaderMap(rows, headerRowIndices) {
  const maxCol = Math.max(...headerRowIndices.map((r) => rows[r]?.length ?? 0));
  const layers = headerRowIndices.map((r) => fillMergedHeaderRow(rows[r] ?? []));

  const byKey = {};
  const byCol = {};

  for (let col = 0; col < maxCol; col += 1) {
    const parts = uniqueStrings(layers.map((layer) => layer[col] ?? ''));
    const key = slugHeader(parts.join(' '));
    if (!key) continue;
    byCol[col] = { key, parts };
    if (!(key in byKey)) byKey[key] = col;
  }

  return { byKey, byCol };
}

function findColumnByAliases(rawHeaderMap, aliases) {
  // Pass 1: exact slug match
  for (const alias of aliases) {
    const key = slugHeader(alias);
    if (rawHeaderMap.byKey[key] !== undefined) return rawHeaderMap.byKey[key];
  }

  const entries = Object.entries(rawHeaderMap.byKey);

  // Pass 2: alias contained in header slug
  for (const alias of aliases) {
    const aliasKey = slugHeader(alias);
    const found = entries.find(([key]) => key.includes(aliasKey));
    if (found) return found[1];
  }

  // Pass 3: header slug contained in alias
  for (const alias of aliases) {
    const aliasKey = slugHeader(alias);
    const found = entries.find(([key]) => aliasKey.includes(key) && key.length >= 3);
    if (found) return found[1];
  }

  return null;
}

// ─────────────────────────────────────────────
// Field alias definitions (unchanged from prior refactor)
// ─────────────────────────────────────────────

const FIELD_ALIASES = {
  descriptor: ['tag number', 'tag #', 'tag'],
  model_number: ['model number', 'model no', 'model #'],
  brand: ['brand', 'manufacturer'],
  qty: ['qty', 'quantity'],
  airflow_cfm: ['supply air blower airflow cfm', 'airflow cfm', 'supply cfm', 'cfm'],
  supply_fan_hp: ['supply air blower hp', 'blower hp', 'fan hp'],
  supply_fan_esp_in_wg: ['supply air blower esp iwg', 'esp in wg', 'esp iwg', 'esp'],
  supply_fan_rpm: ['supply air blower blwr rpm', 'blwr rpm', 'blower rpm', 'fan rpm', 'rpm'],
  cooling_total_mbh: ['cooling capacity mbh total', 'cooling total mbh', 'cooling capacity total', 'total mbh'],
  cooling_sensible_mbh: ['cooling capacity mbh sens', 'cooling sensible mbh', 'cooling capacity sens', 'sensible mbh'],
  unit_eer: ['cooling eer', 'unit eer', 'eer'],
  seer_ieer: ['cooling seer ieer', 'cooling seerieer', 'seer ieer', 'seer ieerr', 'ieer', 'seer'],
  refrigerant: ['cooling refrigerant type', 'cooling refrigerant', 'refrigerant type', 'refrigerant', 'refrig'],
  gas_heat_input_mbh: ['heating gas heat input mbh', 'heating gas heat mbh', 'heating gas input mbh', 'gas heat input', 'gas input mbh', 'gas heat mbh'],
  gas_heat_output_mbh: ['heating gas heat output mbh', 'heating gas output mbh', 'gas heat output', 'gas output mbh', 'gas heat out'],
  electric_heat_kw: ['electric heater kw', 'electric heat kw', 'elec heat kw', 'heater kw', 'electric kw', 'elec kw'],
  heatpump_capacity_mbh: ['heat pump ratings mbh', 'heat pump capacity mbh', 'heat pump mbh', 'hp ratings mbh', 'hp capacity mbh', 'hp mbh'],
  voltage: ['electrical voltage', 'elec voltage', 'volt ph', 'voltage', 'volt'],
  mca: ['electrical mca', 'elec mca', 'min circuit amps', 'min circ amps', 'mca'],
  mocp: ['electrical max fuse', 'elec max fuse', 'electrical ocpd', 'max fuse', 'max ocpd', 'mocp', 'ocpd'],
  weight_lbs: ['electrical operating weight lbs', 'operating weight lbs', 'oper wt lbs', 'oper weight lbs', 'electrical weight lbs', 'weight lbs', 'weight'],
  remarks: ['remarks', 'notes'],
};

const REQUIRED_FIELDS = ['descriptor', 'model_number', 'brand', 'qty', 'voltage', 'mca', 'mocp', 'weight_lbs'];

function buildColumnMap(rows) {
  const anchorRowIndex = findAnchorRowIndex(rows);

  if (anchorRowIndex === null) {
    return {
      columns: {},
      anchorRowIndex: null,
      firstDataRowIndex: null,
      missing: REQUIRED_FIELDS,
      warnings: ['Could not locate header row (expected a row containing "Tag" and "Model Number" in the first 20 rows).'],
    };
  }

  const headerRowIndices = anchorRowIndex >= 2
    ? [anchorRowIndex - 2, anchorRowIndex - 1, anchorRowIndex]
    : anchorRowIndex === 1
      ? [anchorRowIndex - 1, anchorRowIndex]
      : [anchorRowIndex];

  const rawHeaderMap = buildRawHeaderMap(rows, headerRowIndices);

  const columns = {};
  for (const [field, aliases] of Object.entries(FIELD_ALIASES)) {
    columns[field] = findColumnByAliases(rawHeaderMap, aliases);
  }

  // Collision detection
  const criticalPairs = [
    ['mca', 'mocp'],
    ['mca', 'refrigerant'],
    ['mocp', 'refrigerant'],
    ['voltage', 'mca'],
    ['voltage', 'mocp'],
  ];
  const warnings = [];
  for (const [a, b] of criticalPairs) {
    if (columns[a] !== null && columns[a] !== undefined &&
        columns[b] !== null && columns[b] !== undefined &&
        columns[a] === columns[b]) {
      warnings.push(`Column collision: "${a}" and "${b}" both resolved to column ${columns[a]} — "${b}" cleared.`);
      columns[b] = null;
    }
  }

  const firstDataRowIndex = anchorRowIndex + 1;
  const sampleWarnings = validateSampleDataRow(rows[firstDataRowIndex] ?? [], columns, firstDataRowIndex + 1);
  warnings.push(...sampleWarnings);

  const missing = REQUIRED_FIELDS.filter((f) => columns[f] === null || columns[f] === undefined);

  console.log(`[buildColumnMap] anchorRowIndex=${anchorRowIndex} firstDataRowIndex=${firstDataRowIndex}`);
  console.log('[buildColumnMap] columns:', JSON.stringify(columns));
  if (warnings.length) console.warn('[buildColumnMap] warnings:', warnings);

  return { columns, anchorRowIndex, firstDataRowIndex, missing, warnings };
}

function validateSampleDataRow(rowCells, columns, rowNumber) {
  const warnings = [];
  const read = (field) => {
    const col = columns[field];
    return col !== null && col !== undefined ? normalizeText(rowCells[col] ?? '') : '';
  };

  const mcaVal = read('mca');
  const mocpVal = read('mocp');
  const refrigerantVal = read('refrigerant');
  const voltageVal = read('voltage');

  if (mcaVal && !/^\d+(\.\d+)?$/.test(mcaVal.trim())) {
    warnings.push(`Row ${rowNumber}: mca value "${mcaVal}" is not numeric — column may be misidentified.`);
  }
  if (mocpVal && /^R\d{3}/i.test(mocpVal.trim())) {
    warnings.push(`Row ${rowNumber}: mocp value "${mocpVal}" looks like a refrigerant code — column likely misidentified.`);
  }
  if (refrigerantVal && /^\d+(\.\d+)?$/.test(refrigerantVal.trim())) {
    warnings.push(`Row ${rowNumber}: refrigerant value "${refrigerantVal}" is purely numeric — column may be misidentified.`);
  }
  if (voltageVal && !/208|460|230|575/i.test(voltageVal)) {
    warnings.push(`Row ${rowNumber}: voltage value "${voltageVal}" does not look like a voltage — column may be misidentified.`);
  }

  return warnings;
}

function getMappedCell(rowCells, columnMap, fieldName) {
  const col = columnMap.columns[fieldName];
  return col !== null && col !== undefined ? normalizeText(rowCells[col] ?? '') : '';
}

// ─────────────────────────────────────────────
// Descriptor and heat field parsing (unchanged)
// ─────────────────────────────────────────────

function parseDescriptorBasics(descriptor) {
  const text = normalizeText(descriptor);
  const tonnageMatch = text.match(/^(\d+(?:\.\d+)?)\s*-?\s*Ton/i);
  const voltageMatch = text.match(/\b(2083|208-3|208\/3|208-3-60|4603|460-3|460\/3|460-3-60)\b/i);

  const tonnage = tonnageMatch ? normalizeTonnage(tonnageMatch[1]) : '';
  const voltage = normalizeVoltage(voltageMatch ? voltageMatch[1] : '');
  const familyLabel = /\bHP\b|\bHeat Pump\b/i.test(text) ? 'Heat Pump' : 'AC';
  const familyKey = familyLabel === 'Heat Pump' ? 'hp' : 'ac';
  const efficiencyLabel = 'Standard';
  const efficiencyKey = normalizeEfficiency(efficiencyLabel);

  return {
    family_key: familyKey,
    family_label: familyLabel,
    efficiency_key: efficiencyKey,
    efficiency_label: efficiencyLabel,
    tonnage_key: tonnage || '',
    tonnage_value: tonnage ? Number(tonnage) : 0,
    voltage_key: voltage || '',
    voltage_label: voltage || '',
  };
}

function deriveHeatFieldsFromRow(rowData) {
  const gasInputMbh = normalizeNumericText(rowData.raw_gas_heat_input_mbh);
  const gasOutputMbh = normalizeNumericText(rowData.raw_gas_heat_output_mbh);
  const electricKw = normalizeNumericText(rowData.raw_electric_heat_kw);

  if (hasMeaningfulValue(gasInputMbh) || hasMeaningfulValue(gasOutputMbh)) {
    const gasCapacity = gasInputMbh || gasOutputMbh;
    return {
      aux_heat_type_key: 'gas',
      aux_heat_type_label: 'Aluminum Gas Heat',
      aux_heat_capacity_key: gasCapacity,
      aux_heat_capacity_label: gasCapacity,
    };
  }

  if (hasMeaningfulValue(electricKw)) {
    return {
      aux_heat_type_key: 'electric',
      aux_heat_type_label: 'Electric Heat',
      aux_heat_capacity_key: electricKw,
      aux_heat_capacity_label: electricKw,
    };
  }

  return {
    aux_heat_type_key: 'none',
    aux_heat_type_label: 'None',
    aux_heat_capacity_key: '',
    aux_heat_capacity_label: '',
  };
}

// ─────────────────────────────────────────────
// Upload handler
// ─────────────────────────────────────────────

async function handleUploadTemplate(request, env) {
  const contentType = request.headers.get('Content-Type') || '';
  if (!contentType.toLowerCase().includes('multipart/form-data')) {
    return json({ error: 'Expected multipart/form-data upload.' }, 400);
  }

  const formData = await request.formData();
  const file = formData.get('file');

  if (!file || typeof file === 'string') {
    return json({ error: 'Missing file field.' }, 400);
  }

  const arrayBuffer = await file.arrayBuffer();
  const bytes = new Uint8Array(arrayBuffer);

  if (bytes.byteLength < 512) {
    return json({ error: `File appears empty or too small (${bytes.byteLength} bytes).` }, 400);
  }

  // Verify XLSX/ZIP magic bytes: PK\x03\x04
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b || bytes[2] !== 0x03 || bytes[3] !== 0x04) {
    return json({ error: 'File does not appear to be a valid XLSX workbook.' }, 400);
  }

  const key = normalizeText(file.name) || `upload-${Date.now()}.xlsx`;

  await env.TEMPLATES.put(key, bytes, {
    httpMetadata: {
      contentType: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
  });

  return json({ ok: true, key });
}

// ─────────────────────────────────────────────
// Database helpers (unchanged)
// ─────────────────────────────────────────────

async function insertBatch(env, payload) {
  const result = await env.DB.prepare(`
    INSERT INTO import_batches (source_filename, source_sheet, vendor, product_line, notes)
    VALUES (?, ?, ?, ?, ?)
  `)
    .bind(
      payload.source_filename,
      payload.source_sheet || 'Schedule',
      payload.vendor || null,
      payload.product_line || null,
      payload.notes || null,
    )
    .run();

  return result.meta.last_row_id;
}

async function insertStagingRow(env, batchId, rowData) {
  const result = await env.DB.prepare(`
    INSERT INTO staging_schedule_rows (
      batch_id, source_row_number, source_descriptor, raw_model_number, raw_brand, raw_qty,
      raw_airflow_cfm, raw_supply_fan_hp, raw_supply_fan_esp_in_wg, raw_supply_fan_rpm,
      raw_cooling_total_mbh, raw_cooling_sensible_mbh, raw_unit_eer, raw_seer_ieer,
      raw_refrigerant, raw_heating_input_mbh, raw_heating_output_mbh, raw_voltage,
      raw_mca, raw_mocp, raw_weight_lbs, raw_remarks,
      family_key, family_label, efficiency_key, efficiency_label,
      tonnage_key, tonnage_value, voltage_key, voltage_label,
      aux_heat_type_key, aux_heat_type_label, aux_heat_capacity_key, aux_heat_capacity_label,
      parse_status, parse_notes
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).bind(
    batchId,
    rowData.source_row_number,
    rowData.source_descriptor,
    rowData.raw_model_number,
    rowData.raw_brand,
    rowData.raw_qty,
    rowData.raw_airflow_cfm,
    rowData.raw_supply_fan_hp,
    rowData.raw_supply_fan_esp_in_wg,
    rowData.raw_supply_fan_rpm,
    rowData.raw_cooling_total_mbh,
    rowData.raw_cooling_sensible_mbh,
    rowData.raw_unit_eer,
    rowData.raw_seer_ieer,
    rowData.raw_refrigerant,
    rowData.raw_heating_input_mbh,
    rowData.raw_heating_output_mbh,
    rowData.raw_voltage,
    rowData.raw_mca,
    rowData.raw_mocp,
    rowData.raw_weight_lbs,
    rowData.raw_remarks,
    rowData.family_key,
    rowData.family_label,
    rowData.efficiency_key,
    rowData.efficiency_label,
    rowData.tonnage_key,
    rowData.tonnage_value,
    rowData.voltage_key,
    rowData.voltage_label,
    rowData.aux_heat_type_key,
    rowData.aux_heat_type_label,
    rowData.aux_heat_capacity_key,
    rowData.aux_heat_capacity_label,
    rowData.parse_status,
    rowData.parse_notes,
  ).run();

  return result.meta.last_row_id;
}

async function insertImportModelResult(env, batchId, stagingRowId, modelNumber, unitModelId, action, reason = null) {
  await env.DB.prepare(`
    INSERT INTO import_model_results (batch_id, staging_row_id, model_number, unit_model_id, action, reason)
    VALUES (?, ?, ?, ?, ?, ?)
  `).bind(batchId, stagingRowId, modelNumber || null, unitModelId || null, action, reason).run();
}

async function upsertUnitModelV2(env, batchId, stagedRow) {
  const modelNumber = stagedRow.raw_model_number || '';
  const familyKey = stagedRow.family_key || '';
  const familyLabel = stagedRow.family_label || '';
  const tonnageKey = stagedRow.tonnage_key || '';
  const tonnageValue = stagedRow.tonnage_value ?? 0;
  const voltageKey = stagedRow.voltage_key || '';
  const voltageLabel = stagedRow.voltage_label || '';
  const auxHeatTypeKey = stagedRow.aux_heat_type_key || 'None';
  const auxHeatTypeLabel = stagedRow.aux_heat_type_label || 'None';
  const auxHeatCapacityKey = stagedRow.aux_heat_capacity_key || '';
  const auxHeatCapacityLabel = stagedRow.aux_heat_capacity_label || '';
  const efficiencyLabel = stagedRow.efficiency_label || 'Standard';
  const efficiencyKey = normalizeEfficiency(stagedRow.efficiency_key || stagedRow.efficiency_label || 'Standard');

  const existing = await env.DB.prepare(`SELECT * FROM unit_models_v2 WHERE model_number = ?`).bind(modelNumber).first();

  if (!existing) {
    const inserted = await env.DB.prepare(`
      INSERT INTO unit_models_v2 (
        model_number, family_key, family_label, tonnage_key, tonnage_value,
        voltage_key, voltage_label, aux_heat_type_key, aux_heat_type_label,
        aux_heat_capacity_key, aux_heat_capacity_label, efficiency_key, efficiency_label,
        source_batch_id
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      modelNumber, familyKey, familyLabel, tonnageKey, tonnageValue,
      voltageKey, voltageLabel, auxHeatTypeKey, auxHeatTypeLabel,
      auxHeatCapacityKey, auxHeatCapacityLabel, efficiencyKey, efficiencyLabel, batchId
    ).run();

    return { action: 'inserted', unitModelId: inserted.meta.last_row_id };
  }

  const changed = [
    String(existing.family_key || '') !== String(familyKey),
    String(existing.family_label || '') !== String(familyLabel),
    String(existing.tonnage_key || '') !== String(tonnageKey),
    Number(existing.tonnage_value ?? 0) !== Number(tonnageValue),
    String(existing.voltage_key || '') !== String(voltageKey),
    String(existing.voltage_label || '') !== String(voltageLabel),
    String(existing.aux_heat_type_key || '') !== String(auxHeatTypeKey),
    String(existing.aux_heat_type_label || '') !== String(auxHeatTypeLabel),
    String(existing.aux_heat_capacity_key || '') !== String(auxHeatCapacityKey),
    String(existing.aux_heat_capacity_label || '') !== String(auxHeatCapacityLabel),
    String(existing.efficiency_key || '') !== String(efficiencyKey),
    String(existing.efficiency_label || '') !== String(efficiencyLabel),
  ].some(Boolean);

  if (!changed) return { action: 'unchanged', unitModelId: existing.id };

  await env.DB.prepare(`
    UPDATE unit_models_v2
    SET family_key = ?, family_label = ?, tonnage_key = ?, tonnage_value = ?,
        voltage_key = ?, voltage_label = ?, aux_heat_type_key = ?, aux_heat_type_label = ?,
        aux_heat_capacity_key = ?, aux_heat_capacity_label = ?,
        efficiency_key = ?, efficiency_label = ?, source_batch_id = ?,
        updated_at = CURRENT_TIMESTAMP
    WHERE id = ?
  `).bind(
    familyKey, familyLabel, tonnageKey, tonnageValue,
    voltageKey, voltageLabel, auxHeatTypeKey, auxHeatTypeLabel,
    auxHeatCapacityKey, auxHeatCapacityLabel,
    efficiencyKey, efficiencyLabel, batchId, existing.id
  ).run();

  return { action: 'updated', unitModelId: existing.id };
}

// ─────────────────────────────────────────────
// Main import pipeline
// ─────────────────────────────────────────────

function parseWorkbookRow(rowCells, rowNumber, columnMap) {
  const descriptor = getMappedCell(rowCells, columnMap, 'descriptor');
  const modelNumber = getMappedCell(rowCells, columnMap, 'model_number');
  const brand = getMappedCell(rowCells, columnMap, 'brand');
  const qty = getMappedCell(rowCells, columnMap, 'qty');

  if (!descriptor || /^tag\b/i.test(descriptor)) return null;
  if (/^model number$/i.test(modelNumber)) return null;
  if (!/^\d+(?:\.\d+)?\s*-?\s*ton/i.test(descriptor)) return null;
  if (!modelNumber) return null;

  const rawGasHeatInputMbh = getMappedCell(rowCells, columnMap, 'gas_heat_input_mbh');
  const rawGasHeatOutputMbh = getMappedCell(rowCells, columnMap, 'gas_heat_output_mbh');
  const rawElectricHeatKw = getMappedCell(rowCells, columnMap, 'electric_heat_kw');
  const rawHeatPumpCapacityMbh = getMappedCell(rowCells, columnMap, 'heatpump_capacity_mbh');

  const descriptorFields = parseDescriptorBasics(descriptor);
  const derivedHeatFields = deriveHeatFieldsFromRow({
    raw_gas_heat_input_mbh: rawGasHeatInputMbh,
    raw_gas_heat_output_mbh: rawGasHeatOutputMbh,
    raw_electric_heat_kw: rawElectricHeatKw,
  });

  const rowData = {
    source_row_number: rowNumber,
    source_descriptor: descriptor,
    raw_model_number: modelNumber,
    raw_brand: brand,
    raw_qty: qty,
    raw_airflow_cfm: getMappedCell(rowCells, columnMap, 'airflow_cfm'),
    raw_supply_fan_hp: getMappedCell(rowCells, columnMap, 'supply_fan_hp'),
    raw_supply_fan_esp_in_wg: getMappedCell(rowCells, columnMap, 'supply_fan_esp_in_wg'),
    raw_supply_fan_rpm: getMappedCell(rowCells, columnMap, 'supply_fan_rpm'),
    raw_cooling_total_mbh: getMappedCell(rowCells, columnMap, 'cooling_total_mbh'),
    raw_cooling_sensible_mbh: getMappedCell(rowCells, columnMap, 'cooling_sensible_mbh'),
    raw_unit_eer: getMappedCell(rowCells, columnMap, 'unit_eer'),
    raw_seer_ieer: getMappedCell(rowCells, columnMap, 'seer_ieer'),
    raw_refrigerant: getMappedCell(rowCells, columnMap, 'refrigerant'),
    raw_heating_input_mbh: rawGasHeatInputMbh || rawElectricHeatKw,
    raw_heating_output_mbh: rawHeatPumpCapacityMbh || rawGasHeatOutputMbh,
    raw_voltage: getMappedCell(rowCells, columnMap, 'voltage'),
    raw_mca: getMappedCell(rowCells, columnMap, 'mca'),
    raw_mocp: getMappedCell(rowCells, columnMap, 'mocp'),
    raw_weight_lbs: getMappedCell(rowCells, columnMap, 'weight_lbs'),
    raw_remarks: getMappedCell(rowCells, columnMap, 'remarks'),
    ...descriptorFields,
    ...derivedHeatFields,
    parse_status: 'parsed',
    parse_notes: null,
  };

  if (hasMeaningfulValue(rawHeatPumpCapacityMbh)) {
    rowData.family_key = 'hp';
    rowData.family_label = 'Heat Pump';
    if (
      !hasMeaningfulValue(rawGasHeatInputMbh) &&
      !hasMeaningfulValue(rawGasHeatOutputMbh) &&
      !hasMeaningfulValue(rawElectricHeatKw)
    ) {
      rowData.aux_heat_type_key = 'none';
      rowData.aux_heat_type_label = 'None';
      rowData.aux_heat_capacity_key = '';
      rowData.aux_heat_capacity_label = '';
    }
  }

  return rowData;
}

async function stageDsCommercialWorkbook(env, payload) {
  const workbook = await loadWorkbookFromR2(env, payload.source_filename);
  const { worksheet, name: sheetName } = getWorksheet(workbook, payload.source_sheet || 'Schedule');

  console.log(`[stageDsCommercialWorkbook] using sheet "${sheetName}"`);

  const rows = sheetToRows(worksheet);
  const columnMap = buildColumnMap(rows);

  if (columnMap.missing.length) {
    throw new Error(
      `Could not identify required columns: ${columnMap.missing.join(', ')}. ` +
      (columnMap.warnings.length ? `Warnings: ${columnMap.warnings.join(' | ')}` : '')
    );
  }

  if (columnMap.warnings.length) {
    console.warn('[stageDsCommercialWorkbook] schema warnings:', columnMap.warnings);
  }

  const batchId = await insertBatch(env, payload);
  const stagedRows = [];

  for (let r = columnMap.firstDataRowIndex; r < rows.length; r += 1) {
    const parsed = parseWorkbookRow(rows[r], r + 1, columnMap); // r+1 = 1-based row number
    if (parsed) stagedRows.push(parsed);
  }

  const duplicateCounts = new Map();
  for (const row of stagedRows) {
    duplicateCounts.set(row.raw_model_number, (duplicateCounts.get(row.raw_model_number) || 0) + 1);
  }

  const seenModels = new Set();

  for (const row of stagedRows) {
    const stagingRowId = await insertStagingRow(env, batchId, row);
    row.id = stagingRowId;

    const occurrences = duplicateCounts.get(row.raw_model_number) || 0;
    if (occurrences > 1 && seenModels.has(row.raw_model_number)) {
      await insertImportModelResult(env, batchId, stagingRowId, row.raw_model_number, null, 'duplicate_in_batch', 'Duplicate model number within import batch');
      continue;
    }

    seenModels.add(row.raw_model_number);
    const upsert = await upsertUnitModelV2(env, batchId, row);
    await insertImportModelResult(env, batchId, stagingRowId, row.raw_model_number, upsert.unitModelId, upsert.action, null);
  }

  return batchId;
}

// ─────────────────────────────────────────────
// Batch query helpers (unchanged)
// ─────────────────────────────────────────────

async function getBatch(env, batchId) {
  return (await env.DB.prepare(`SELECT * FROM import_batches WHERE id = ?`).bind(batchId).first()) || null;
}

async function getBatchSummary(env, batchId) {
  const totals = await env.DB.prepare(`
    SELECT
      COUNT(*) AS rows_staged,
      SUM(CASE WHEN parse_status = 'parsed' THEN 1 ELSE 0 END) AS rows_parsed,
      SUM(CASE WHEN parse_status != 'parsed' THEN 1 ELSE 0 END) AS rows_failed,
      SUM(CASE WHEN TRIM(COALESCE(parse_notes, '')) != '' THEN 1 ELSE 0 END) AS rows_with_warnings,
      COUNT(DISTINCT NULLIF(TRIM(raw_model_number), '')) AS unique_models_in_batch
    FROM staging_schedule_rows WHERE batch_id = ?
  `).bind(batchId).first();

  const actions = await env.DB.prepare(`
    SELECT
      COALESCE(SUM(CASE WHEN action = 'inserted' THEN 1 ELSE 0 END), 0) AS catalog_inserts,
      COALESCE(SUM(CASE WHEN action = 'updated' THEN 1 ELSE 0 END), 0) AS catalog_updates,
      COALESCE(SUM(CASE WHEN action = 'unchanged' THEN 1 ELSE 0 END), 0) AS catalog_unchanged,
      COALESCE(SUM(CASE WHEN action = 'duplicate_in_batch' THEN 1 ELSE 0 END), 0) AS duplicates_in_batch
    FROM import_model_results WHERE batch_id = ?
  `).bind(batchId).first();

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

function summarizeIssueRows(rows) {
  return rows.map((row) => ({
    source_row_number: row.source_row_number,
    model_number: row.raw_model_number,
    parse_status: row.parse_status,
    parse_notes: cleanBlank(row.parse_notes),
  }));
}

async function getBatchIssues(env, batchId) {
  const duplicates = await env.DB.prepare(`
    SELECT raw_model_number AS model_number, COUNT(*) AS duplicate_count,
           GROUP_CONCAT(source_row_number) AS source_row_numbers
    FROM staging_schedule_rows
    WHERE batch_id = ? AND NULLIF(TRIM(raw_model_number), '') IS NOT NULL
    GROUP BY raw_model_number HAVING COUNT(*) > 1
    ORDER BY duplicate_count DESC, raw_model_number
  `).bind(batchId).all();

  const parseIssues = await env.DB.prepare(`
    SELECT source_row_number, raw_model_number, parse_status, parse_notes
    FROM staging_schedule_rows
    WHERE batch_id = ? AND (parse_status != 'parsed' OR TRIM(COALESCE(parse_notes, '')) != '')
    ORDER BY source_row_number
  `).bind(batchId).all();

  const issues = [];
  for (const row of coerceArray(duplicates.results)) {
    issues.push({
      type: 'duplicate_model_in_batch',
      model_number: row.model_number,
      count: toInt(row.duplicate_count),
      source_row_numbers: String(row.source_row_numbers || '').split(',').map((v) => toInt(v)).filter(Boolean),
    });
  }

  const warningRows = summarizeIssueRows(coerceArray(parseIssues.results));
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

// ─────────────────────────────────────────────
// Batch GET handlers (unchanged)
// ─────────────────────────────────────────────

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
  const rows = await env.DB.prepare(`
    SELECT * FROM staging_schedule_rows WHERE batch_id = ? ORDER BY source_row_number, id
  `).bind(batchId).all();
  return json({ ok: true, batch, rows: coerceArray(rows.results) });
}

async function handleGetImportBatchCatalogResults(env, batchId) {
  const batch = await getBatch(env, batchId);
  if (!batch) return json({ error: 'Import batch not found.' }, 404);
  const rows = await env.DB.prepare(`
    SELECT imr.id, imr.batch_id, imr.staging_row_id,
           ssr.source_row_number, ssr.source_descriptor, ssr.raw_model_number,
           imr.model_number, imr.unit_model_id, imr.action, imr.reason,
           um.family_key, um.tonnage_key, um.voltage_key,
           um.aux_heat_type_key, um.aux_heat_capacity_key, um.efficiency_key,
           imr.created_at
    FROM import_model_results imr
    JOIN staging_schedule_rows ssr ON ssr.id = imr.staging_row_id
    LEFT JOIN unit_models_v2 um ON um.id = imr.unit_model_id
    WHERE imr.batch_id = ?
    ORDER BY ssr.source_row_number, imr.id
  `).bind(batchId).all();
  return json({ ok: true, batch, results: coerceArray(rows.results) });
}

// ─────────────────────────────────────────────
// Catalog query (unchanged)
// ─────────────────────────────────────────────

async function listCatalog(env, filters = {}) {
  let sql = `
    SELECT m.*, d.cutsheet_url, d.accessories_url, d.wiring_url, d.iom_url
    FROM unit_models m
    LEFT JOIN unit_documents d ON d.model_id = m.id
    WHERE 1 = 1
  `;
  const binds = [];

  if (filters.family) { sql += ' AND m.family = ?'; binds.push(filters.family); }
  if (filters.efficiency) { sql += ' AND m.efficiency = ?'; binds.push(filters.efficiency); }
  if (filters.tonnage !== '' && filters.tonnage !== null && filters.tonnage !== undefined) {
    sql += ' AND m.tonnage = ?'; binds.push(filters.tonnage);
  }
  if (filters.voltage) { sql += ' AND m.voltage = ?'; binds.push(filters.voltage); }
  if (filters.heatType) { sql += ' AND m.heat_type = ?'; binds.push(filters.heatType); }
  if (normalizeHeatType(filters.heatType) !== 'None') {
    sql += ' AND m.heat_capacity = ?'; binds.push(filters.heatCapacity || '');
  } else {
    sql += ` AND COALESCE(m.heat_capacity, '') = ''`;
  }

  sql += ' ORDER BY m.model_number';
  const result = await env.DB.prepare(sql).bind(...binds).all();
  return result.results || [];
}

// ─────────────────────────────────────────────
// Schedule resolution (unchanged)
// ─────────────────────────────────────────────

async function findMatchingImportedRow(env, unit) {
  const requestedModelNumber = normalizeText(unit.modelNumber || unit.selectedModelNumber);

  if (requestedModelNumber) {
    const byModel = await env.DB.prepare(`
      SELECT * FROM staging_schedule_rows
      WHERE raw_model_number = ? AND parse_status = 'parsed'
      ORDER BY id DESC LIMIT 1
    `).bind(requestedModelNumber).first();
    if (byModel) return byModel;
  }

  const family = normalizeFamily(unit.family);
  const tonnage = String(unit.tonnage);
  const voltage = normalizeVoltage(unit.voltage);
  const heatType = normalizeHeatType(unit.heatType);
  const heatCapacity = normalizeHeatCapacity(unit.heatCapacity || '');

  const exact = await env.DB.prepare(`
    SELECT * FROM staging_schedule_rows
    WHERE family_label = ? AND CAST(tonnage_value AS TEXT) = ?
      AND voltage_label = ? AND aux_heat_type_label = ? AND aux_heat_capacity_key = ?
      AND parse_status = 'parsed'
    ORDER BY id DESC LIMIT 1
  `).bind(family, tonnage, voltage, heatType, heatType === 'None' ? '' : heatCapacity).first();

  if (exact) return exact;

  return await env.DB.prepare(`
    SELECT * FROM staging_schedule_rows
    WHERE family_label = ? AND CAST(tonnage_value AS TEXT) = ?
      AND voltage_label = ? AND parse_status = 'parsed'
    ORDER BY id DESC LIMIT 1
  `).bind(family, tonnage, voltage).first() || null;
}

function buildResolvedScheduleRow(unit, match, index = 0) {
  const unitType = asBlank(match?.family_label || unit.family);
  const isHeatPump = normalizeFamily(unitType) === 'Heat Pump';

  return {
    tag: normalizeText(unit.tag) || `RTU-${index + 1}`,
    areaServed: asBlank(unit.areaServed),
    manufacturer: asBlank(match?.raw_brand || 'Tempmaster'),
    modelNumber: asBlank(match?.raw_model_number || buildSelectionCode(unit)),
    nominalTons: asBlank(match?.tonnage_value ?? unit.tonnage),
    unitType,
    unitEer: asBlank(match?.raw_unit_eer),
    seerIeerr: asBlank(match?.raw_seer_ieer),
    supplyCfm: asBlank(match?.raw_airflow_cfm),
    supplyEsp: asBlank(match?.raw_supply_fan_esp_in_wg),
    supplyQty: asBlank(match?.raw_qty || 1),
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
    heatingInput: isHeatPump ? '' : asBlank(match?.raw_heating_input_mbh || unit.heatCapacity),
    heatingTotalCapacity: isHeatPump
      ? asBlank(match?.raw_heating_output_mbh || match?.raw_heating_input_mbh)
      : '',
    heatingOutput: asBlank(match?.raw_heating_output_mbh),
    voltPh: asBlank(match?.raw_voltage || unit.voltage),
    mca: asBlank(match?.raw_mca),
    mocp: asBlank(match?.raw_mocp),
    weight: asBlank(match?.raw_weight_lbs),
    remarks: asBlank(match?.raw_remarks || optionSummary(unit, match)),
    selectionCode: buildSelectionCode(unit),
    matchFound: Boolean(match),
    cutSheetUrl: '',
    accessoriesUrl: '',
    wiringUrl: '',
    iomUrl: '',
  };
}

async function resolveScheduleRows(env, units) {
  const rows = [];
  for (let index = 0; index < units.length; index += 1) {
    const unit = units[index];
    const match = await findMatchingImportedRow(env, unit);
    rows.push(buildResolvedScheduleRow(unit, match, index));
  }
  return rows;
}

// ─────────────────────────────────────────────
// Export workbook — built fresh with SheetJS
// Column order matches the original COL map exactly.
// Basic styling is applied via cell styles (SheetJS community supports
// column widths, number formats, and basic fills/borders).
// ─────────────────────────────────────────────

const SCHEDULE_COLUMNS = [
  { field: 'tag',               header: 'Tag',                    width: 12 },
  { field: 'areaServed',        header: 'Area Served',            width: 20 },
  { field: 'manufacturer',      header: 'Manufacturer',           width: 16 },
  { field: 'modelNumber',       header: 'Model Number',           width: 28 },
  { field: 'nominalTons',       header: 'Nominal Tons',           width: 12 },
  { field: 'unitType',          header: 'Unit Type',              width: 14 },
  { field: 'unitEer',           header: 'EER',                    width: 10 },
  { field: 'seerIeerr',         header: 'SEER / IEER',            width: 12 },
  { field: 'supplyCfm',         header: 'Supply CFM',             width: 12 },
  { field: 'supplyEsp',         header: 'ESP (in. w.g.)',         width: 14 },
  { field: 'supplyQty',         header: 'Fan Qty',                width: 10 },
  { field: 'supplyBhp',         header: 'BHP',                    width: 10 },
  { field: 'supplyHp',          header: 'Fan HP',                 width: 10 },
  { field: 'supplyRpm',         header: 'Fan RPM',                width: 10 },
  { field: 'coolingEat',        header: 'Cooling EAT',            width: 12 },
  { field: 'coolingLat',        header: 'Cooling LAT',            width: 12 },
  { field: 'coolingSensible',   header: 'Cooling Sensible (MBH)', width: 22 },
  { field: 'coolingTotal',      header: 'Cooling Total (MBH)',    width: 20 },
  { field: 'heatingCfm',        header: 'Heating CFM',            width: 12 },
  { field: 'heatingEat',        header: 'Heating EAT',            width: 12 },
  { field: 'heatingLat',        header: 'Heating LAT',            width: 12 },
  { field: 'heatingInput',      header: 'Heating Total Capacity', width: 22 },
  { field: 'heatingOutput',     header: 'Heating Output (MBH)',   width: 20 },
  { field: 'voltPh',            header: 'Volt / Ph',              width: 12 },
  { field: 'mca',               header: 'MCA',                    width: 10 },
  { field: 'mocp',              header: 'MOCP',                   width: 10 },
  { field: 'weight',            header: 'Weight (lbs)',           width: 14 },
  { field: 'remarks',           header: 'Remarks',                width: 36 },
];

async function createWorkbook(env, units) {
  const resolvedRows = await resolveScheduleRows(env, units);

  const wb = XLSX.utils.book_new();

  // Build array-of-arrays: first row = headers, remaining = data
  const headerRow = SCHEDULE_COLUMNS.map((col) => col.header);
  const dataRows = resolvedRows.map((row) =>
    SCHEDULE_COLUMNS.map((col) => {
      const val = row[col.field];
      // Attempt numeric coercion for known numeric fields so Excel treats them as numbers
      if (val !== '' && ['nominalTons', 'supplyCfm', 'supplyHp', 'supplyRpm',
          'coolingSensible', 'coolingTotal', 'heatingInput', 'heatingOutput',
          'mca', 'mocp', 'weight', 'supplyEsp', 'unitEer'].includes(col.field)) {
        const n = Number(val);
        if (Number.isFinite(n)) return n;
      }
      return val ?? '';
    })
  );

  const ws = XLSX.utils.aoa_to_sheet([headerRow, ...dataRows]);

  // Column widths
  ws['!cols'] = SCHEDULE_COLUMNS.map((col) => ({ wch: col.width }));

  // Freeze the header row
  ws['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft' };

  XLSX.utils.book_append_sheet(wb, ws, 'Schedule');

  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
}

// ─────────────────────────────────────────────
// Router (unchanged)
// ─────────────────────────────────────────────

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === 'GET' && url.pathname === '/api/catalog') {
      const filters = {
        family: normalizeFamily(url.searchParams.get('family')),
        efficiency: normalizeEfficiency(url.searchParams.get('efficiency')),
        tonnage: url.searchParams.get('tonnage'),
        voltage: normalizeVoltage(url.searchParams.get('voltage')),
        heatType: normalizeHeatType(url.searchParams.get('heatType')),
        heatCapacity: normalizeHeatCapacity(url.searchParams.get('heatCapacity')),
      };
      const results = await listCatalog(env, filters);
      return json({ items: results });
    }

    if (request.method === 'POST' && url.pathname === '/api/upload-template') {
      try {
        return await handleUploadTemplate(request, env);
      } catch (error) {
        return json({ error: error.message || 'Upload failed.' }, 500);
      }
    }

    if (request.method === 'POST' && url.pathname === '/api/import-schedule') {
      try {
        const payload = await request.json();
        if (!payload?.source_filename) {
          return json({ error: 'source_filename is required.' }, 400);
        }
        const batchId = await stageDsCommercialWorkbook(env, payload);
        const batch = await getBatch(env, batchId);
        const summary = await getBatchSummary(env, batchId);
        const issues = await getBatchIssues(env, batchId);
        return json({ ok: true, batch, summary, issues, links: buildBatchLinks(batchId) });
      } catch (error) {
        return json({ error: error.message || 'Import failed.' }, 500);
      }
    }

    if (request.method === 'POST' && url.pathname === '/api/preview-schedule') {
      try {
        const payload = await request.json();
        const units = Array.isArray(payload?.units) ? payload.units : [];
        const rows = await resolveScheduleRows(env, units);
        return json({ rows });
      } catch (error) {
        console.error('preview-schedule error', error);
        return json({ error: error.message || 'Preview failed' }, 500);
      }
    }

    const batchMatch = url.pathname.match(/^\/api\/import-batches\/(\d+)$/);
    if (request.method === 'GET' && batchMatch) {
      return handleGetImportBatch(env, Number(batchMatch[1]));
    }

    const batchRowsMatch = url.pathname.match(/^\/api\/import-batches\/(\d+)\/staging-rows$/);
    if (request.method === 'GET' && batchRowsMatch) {
      return handleGetImportBatchRows(env, Number(batchRowsMatch[1]));
    }

    const batchCatalogMatch = url.pathname.match(/^\/api\/import-batches\/(\d+)\/catalog-results$/);
    if (request.method === 'GET' && batchCatalogMatch) {
      return handleGetImportBatchCatalogResults(env, Number(batchCatalogMatch[1]));
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
      } catch (error) {
        return new Response(error.message || 'Export failed', { status: 500 });
      }
    }

    return env.ASSETS.fetch(request);
  },
};