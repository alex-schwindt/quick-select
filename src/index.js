import ExcelJS from 'exceljs';

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

function joinSlash(a, b) {
  const parts = [asBlank(a), asBlank(b)].filter((v) => v !== '');
  return parts.length ? parts.join('/') : '';
}

// ─────────────────────────────────────────────
// Workbook cell reading
// ─────────────────────────────────────────────

function getCellValue(row, index) {
  const cell = row.getCell(index);
  const value = cell?.value;

  if (value && typeof value === 'object') {
    if ('text' in value) return normalizeText(value.text);
    if ('result' in value) return normalizeText(value.result);
    if ('richText' in value && Array.isArray(value.richText)) {
      return normalizeText(value.richText.map((part) => part.text || '').join(''));
    }
  }

  return normalizeText(value);
}

// ─────────────────────────────────────────────
// Header detection — two-phase approach:
//   Phase 1: scan every row for a "Tag / Model Number" anchor.
//   Phase 2: build the full column map from that anchor row band.
// This replaces the scoring/fuzzy approach with a deterministic anchor search.
// ─────────────────────────────────────────────

function slugHeader(text) {
  return normalizeText(text)
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/#/g, ' number ')
    .replace(/\//g, ' ')
    .replace(/[\.\/\(\)\-]+/g, ' ')
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

/**
 * Scan rows 1–20 to find an anchor row whose cells contain "Tag" (or "Tag Number")
 * AND "Model Number". This is the bottom row of the header band. The band is then
 * composed of up to 3 rows ending at that anchor row.
 *
 * Returns the anchor row number, or null if not found.
 */
function findAnchorRow(worksheet) {
  const maxRow = Math.min(worksheet.rowCount || 20, 20);

  for (let r = 1; r <= maxRow; r += 1) {
    const row = worksheet.getRow(r);
    const maxCol = worksheet.columnCount || 0;
    let hasTag = false;
    let hasModelNumber = false;

    for (let col = 1; col <= maxCol; col += 1) {
      const raw = getCellValue(row, col).toLowerCase().trim();
      if (raw === 'tag' || raw === 'tag number' || raw === 'tag #') hasTag = true;
      if (raw === 'model number' || raw === 'model no' || raw === 'model #') hasModelNumber = true;
    }

    if (hasTag && hasModelNumber) {
      return r;
    }
  }

  return null;
}

/**
 * Build a column lookup map from a band of header rows.
 * Returns { byKey: { slug -> colIndex }, byCol: { colIndex -> { key, parts } } }
 */
function buildRawHeaderMap(worksheet, headerRows) {
  const maxCol = worksheet.columnCount || 0;
  const layers = headerRows.map((rowNumber) => {
    const row = worksheet.getRow(rowNumber);
    const raw = [];
    for (let col = 1; col <= maxCol; col += 1) {
      raw.push(getCellValue(row, col));
    }
    return fillMergedHeaderRow(raw);
  });

  const byKey = {};
  const byCol = {};

  for (let col = 1; col <= maxCol; col += 1) {
    const parts = uniqueStrings(layers.map((layer) => layer[col - 1]));
    const key = slugHeader(parts.join(' '));
    if (!key) continue;
    byCol[col] = { key, parts };
    if (!(key in byKey)) byKey[key] = col;
  }

  return { byKey, byCol, headerRows };
}

/**
 * Find a column by trying a prioritized list of slug aliases.
 * Priority order:
 *   1. Exact slug match
 *   2. Alias slug is contained within a header slug (alias is more specific)
 *   3. Header slug is contained within an alias slug (header is more specific)
 *
 * criticalField=true means we log a warning instead of silently returning null.
 */
function findColumnByAliases(rawHeaderMap, aliases, fieldName = '') {
  // Pass 1: exact match
  for (const alias of aliases) {
    const key = slugHeader(alias);
    if (rawHeaderMap.byKey[key]) return rawHeaderMap.byKey[key];
  }

  // Pass 2: alias contained in header slug (e.g., alias="mca" matches header="electrical_mca")
  const entries = Object.entries(rawHeaderMap.byKey);
  for (const alias of aliases) {
    const aliasKey = slugHeader(alias);
    const found = entries.find(([key]) => key.includes(aliasKey));
    if (found) return found[1];
  }

  // Pass 3: header contained in alias (broader alias matches narrower header)
  for (const alias of aliases) {
    const aliasKey = slugHeader(alias);
    const found = entries.find(([key]) => aliasKey.includes(key) && key.length >= 3);
    if (found) return found[1];
  }

  return null;
}

// ─────────────────────────────────────────────
// Field definitions — all aliases in priority order, most specific first.
// The electrical section aliases are designed so that mocp (max fuse) and mca
// can never resolve to the same column or to the refrigerant column.
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
  // Refrigerant: long compound alias first so it never collides with mocp
  refrigerant: ['cooling refrigerant type', 'cooling refrigerant', 'refrigerant type', 'refrigerant', 'refrig'],
  gas_heat_input_mbh: ['heating gas heat input mbh', 'heating gas heat mbh', 'heating gas input mbh', 'gas heat input', 'gas input mbh', 'gas heat mbh'],
  gas_heat_output_mbh: ['heating gas heat output mbh', 'heating gas output mbh', 'gas heat output', 'gas output mbh', 'gas heat out'],
  electric_heat_kw: ['electric heater kw', 'electric heat kw', 'elec heat kw', 'heater kw', 'electric kw', 'elec kw'],
  heatpump_capacity_mbh: ['heat pump ratings mbh', 'heat pump capacity mbh', 'heat pump mbh', 'hp ratings mbh', 'hp capacity mbh', 'hp mbh'],
  // Electrical fields: compound aliases (with "electrical" prefix) are listed first
  // so a tightly nested merged-header workbook hits the exact compound slug.
  voltage: ['electrical voltage', 'elec voltage', 'volt ph', 'voltage', 'volt'],
  mca: ['electrical mca', 'elec mca', 'min circuit amps', 'min circ amps', 'mca'],
  // mocp must never match "refrigerant" — the aliases here are all electrical-domain terms
  mocp: ['electrical max fuse', 'elec max fuse', 'electrical ocpd', 'max fuse', 'max ocpd', 'mocp', 'ocpd'],
  weight_lbs: ['electrical operating weight lbs', 'operating weight lbs', 'oper wt lbs', 'oper weight lbs', 'electrical weight lbs', 'weight lbs', 'weight'],
  remarks: ['remarks', 'notes'],
};

const REQUIRED_FIELDS = ['descriptor', 'model_number', 'brand', 'qty', 'voltage', 'mca', 'mocp', 'weight_lbs'];

/**
 * Build the column map for a worksheet.
 * Uses anchor-row detection (deterministic) instead of scoring (probabilistic).
 *
 * Returns { columns, rawHeaderMap, headerRows, missing, warnings }
 */
function buildColumnMap(worksheet) {
  const anchorRow = findAnchorRow(worksheet);
  if (anchorRow === null) {
    // Could not find the anchor row at all — return empty map with all fields missing
    return {
      columns: {},
      rawHeaderMap: { byKey: {}, byCol: {}, headerRows: [] },
      headerRows: [],
      missing: REQUIRED_FIELDS,
      warnings: ['Could not locate header row (expected a row containing "Tag" and "Model Number" in the first 20 rows).'],
    };
  }

  // Use up to 3 rows ending at the anchor row (to capture multi-row merged group headers)
  const headerRows = anchorRow >= 3
    ? [anchorRow - 2, anchorRow - 1, anchorRow]
    : anchorRow === 2
      ? [anchorRow - 1, anchorRow]
      : [anchorRow];

  const rawHeaderMap = buildRawHeaderMap(worksheet, headerRows);

  const columns = {};
  for (const [field, aliases] of Object.entries(FIELD_ALIASES)) {
    columns[field] = findColumnByAliases(rawHeaderMap, aliases, field);
  }

  // Collision detection: if two critical fields mapped to the same column, something
  // went wrong. Log a warning and nullify the less-specific one.
  const criticalPairs = [
    ['mca', 'mocp'],
    ['mca', 'refrigerant'],
    ['mocp', 'refrigerant'],
    ['voltage', 'mca'],
    ['voltage', 'mocp'],
  ];
  const warnings = [];
  for (const [a, b] of criticalPairs) {
    if (columns[a] && columns[b] && columns[a] === columns[b]) {
      warnings.push(`Column collision: "${a}" and "${b}" both resolved to column ${columns[a]} — "${b}" will be cleared.`);
      console.warn(`[buildColumnMap] collision: ${a}=${columns[a]} === ${b}=${columns[b]}, clearing ${b}`);
      columns[b] = null;
    }
  }

  // Value-sanity check: sample the first data row after the anchor and verify
  // that mca/mocp/weight_lbs look numeric and refrigerant looks like a refrigerant code.
  const firstDataRowNum = anchorRow + 1;
  const firstDataRow = worksheet.getRow(firstDataRowNum);
  const sampleWarnings = validateSampleRow(firstDataRow, columns, firstDataRowNum);
  warnings.push(...sampleWarnings);

  const missing = REQUIRED_FIELDS.filter((field) => !columns[field]);

  console.log(`[buildColumnMap] anchorRow=${anchorRow} headerRows=[${headerRows}]`);
  console.log('[buildColumnMap] columns:', JSON.stringify(columns));
  if (warnings.length) console.warn('[buildColumnMap] warnings:', warnings);

  return { columns, rawHeaderMap, headerRows, missing, warnings };
}

/**
 * Inspect the first data row to detect cross-field contamination.
 * Returns an array of warning strings (empty = clean).
 */
function validateSampleRow(row, columns, rowNumber) {
  const warnings = [];

  const read = (field) => (columns[field] ? getCellValue(row, columns[field]) : '');

  const voltageVal = read('voltage');
  const mcaVal = read('mca');
  const mocpVal = read('mocp');
  const refrigerantVal = read('refrigerant');
  const weightVal = read('weight_lbs');

  // MCA should be numeric and plausibly between 1 and 1000
  if (mcaVal && !/^\d+(\.\d+)?$/.test(mcaVal.trim())) {
    warnings.push(`Row ${rowNumber}: mca value "${mcaVal}" is not numeric — column may be misidentified.`);
  }

  // MOCP should be numeric OR a known refrigerant pattern should NOT appear there
  if (mocpVal && /^R\d{3}/i.test(mocpVal.trim())) {
    warnings.push(`Row ${rowNumber}: mocp value "${mocpVal}" looks like a refrigerant code — column is likely misidentified. Check header alias for "mocp".`);
  }

  // Refrigerant should look like R454B, R410A, etc. — if it's numeric that's suspicious
  if (refrigerantVal && /^\d+(\.\d+)?$/.test(refrigerantVal.trim())) {
    warnings.push(`Row ${rowNumber}: refrigerant value "${refrigerantVal}" is purely numeric — column may be misidentified.`);
  }

  // Voltage should look like a voltage
  if (voltageVal && !/208|460|230|575/i.test(voltageVal)) {
    warnings.push(`Row ${rowNumber}: voltage value "${voltageVal}" does not look like a voltage — column may be misidentified.`);
  }

  return warnings;
}

function getMappedCellValue(row, columnMap, fieldName) {
  const col = columnMap.columns[fieldName];
  return col ? getCellValue(row, col) : '';
}

// ─────────────────────────────────────────────
// Descriptor and heat field parsing
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
// R2 / workbook loading
// ─────────────────────────────────────────────

async function loadWorkbookFromSource(env, sourceFilename) {
  const r2Object = await env.TEMPLATES.get(sourceFilename);
  if (!r2Object) throw new Error(`Workbook not found in R2: "${sourceFilename}"`);

  const arrayBuffer = await r2Object.arrayBuffer();
  if (!arrayBuffer || arrayBuffer.byteLength < 512) {
    throw new Error(`R2 object "${sourceFilename}" is empty or corrupt (${arrayBuffer?.byteLength ?? 0} bytes).`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  return workbook;
}

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

  const key = normalizeText(file.name) || `upload-${Date.now()}.xlsx`;
  const buffer = await file.arrayBuffer();
  await env.TEMPLATES.put(key, buffer, {
    httpMetadata: {
      contentType: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
  });

  return json({ ok: true, key });
}

// ─────────────────────────────────────────────
// Database helpers
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
  const efficiencyKey = normalizeEfficiency(
    stagedRow.efficiency_key || stagedRow.efficiency_label || 'Standard'
  );

  const existing = await env.DB.prepare(`
    SELECT * FROM unit_models_v2 WHERE model_number = ?
  `).bind(modelNumber).first();

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

  if (!changed) {
    return { action: 'unchanged', unitModelId: existing.id };
  }

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

/**
 * Parse a single data row into a staging row object.
 * All field reads go through getMappedCellValue so there is one
 * canonical place to add per-field transformations.
 */
function parseWorkbookRow(row, rowNumber, columnMap) {
  const descriptor = getMappedCellValue(row, columnMap, 'descriptor');
  const modelNumber = getMappedCellValue(row, columnMap, 'model_number');
  const brand = getMappedCellValue(row, columnMap, 'brand');
  const qty = getMappedCellValue(row, columnMap, 'qty');

  // Skip header rows and rows without the required descriptor/model pattern
  if (!descriptor || /^tag\b/i.test(descriptor)) return null;
  if (/^model number$/i.test(modelNumber)) return null;
  if (!/^\d+(?:\.\d+)?\s*-?\s*ton/i.test(descriptor)) return null;
  if (!modelNumber) return null;

  const rawGasHeatInputMbh = getMappedCellValue(row, columnMap, 'gas_heat_input_mbh');
  const rawGasHeatOutputMbh = getMappedCellValue(row, columnMap, 'gas_heat_output_mbh');
  const rawElectricHeatKw = getMappedCellValue(row, columnMap, 'electric_heat_kw');
  const rawHeatPumpCapacityMbh = getMappedCellValue(row, columnMap, 'heatpump_capacity_mbh');

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
    raw_airflow_cfm: getMappedCellValue(row, columnMap, 'airflow_cfm'),
    raw_supply_fan_hp: getMappedCellValue(row, columnMap, 'supply_fan_hp'),
    raw_supply_fan_esp_in_wg: getMappedCellValue(row, columnMap, 'supply_fan_esp_in_wg'),
    raw_supply_fan_rpm: getMappedCellValue(row, columnMap, 'supply_fan_rpm'),
    raw_cooling_total_mbh: getMappedCellValue(row, columnMap, 'cooling_total_mbh'),
    raw_cooling_sensible_mbh: getMappedCellValue(row, columnMap, 'cooling_sensible_mbh'),
    raw_unit_eer: getMappedCellValue(row, columnMap, 'unit_eer'),
    raw_seer_ieer: getMappedCellValue(row, columnMap, 'seer_ieer'),
    raw_refrigerant: getMappedCellValue(row, columnMap, 'refrigerant'),
    raw_heating_input_mbh: rawGasHeatInputMbh || rawElectricHeatKw,
    raw_heating_output_mbh: rawHeatPumpCapacityMbh || rawGasHeatOutputMbh,
    raw_voltage: getMappedCellValue(row, columnMap, 'voltage'),
    raw_mca: getMappedCellValue(row, columnMap, 'mca'),
    raw_mocp: getMappedCellValue(row, columnMap, 'mocp'),
    raw_weight_lbs: getMappedCellValue(row, columnMap, 'weight_lbs'),
    raw_remarks: getMappedCellValue(row, columnMap, 'remarks'),
    ...descriptorFields,
    ...derivedHeatFields,
    parse_status: 'parsed',
    parse_notes: null,
  };

  // Heat pump override: if HP capacity is present, set family and clear aux heat
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
  const workbook = await loadWorkbookFromSource(env, payload.source_filename);
  const worksheet =
    workbook.getWorksheet(payload.source_sheet || 'Schedule') || workbook.worksheets?.[0];

  if (!worksheet) {
    throw new Error(
      `Worksheet not found. Available sheets: ${workbook.worksheets.map((ws) => ws.name).join(', ')}`
    );
  }

  const columnMap = buildColumnMap(worksheet);

  if (columnMap.missing.length) {
    throw new Error(
      `Could not identify required columns from workbook headers: ${columnMap.missing.join(', ')}. ` +
      (columnMap.warnings.length ? `Warnings: ${columnMap.warnings.join(' | ')}` : '')
    );
  }

  if (columnMap.warnings.length) {
    console.warn('[stageDsCommercialWorkbook] schema warnings:', columnMap.warnings);
  }

  const batchId = await insertBatch(env, payload);
  const stagedRows = [];

  worksheet.eachRow((row, rowNumber) => {
    const parsed = parseWorkbookRow(row, rowNumber, columnMap);
    if (parsed) stagedRows.push(parsed);
  });

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
      await insertImportModelResult(
        env, batchId, stagingRowId, row.raw_model_number,
        null, 'duplicate_in_batch', 'Duplicate model number within import batch'
      );
      continue;
    }

    seenModels.add(row.raw_model_number);
    const upsert = await upsertUnitModelV2(env, batchId, row);
    await insertImportModelResult(env, batchId, stagingRowId, row.raw_model_number, upsert.unitModelId, upsert.action, null);
  }

  return batchId;
}

// ─────────────────────────────────────────────
// Batch query helpers
// ─────────────────────────────────────────────

async function getBatch(env, batchId) {
  const result = await env.DB.prepare(`SELECT * FROM import_batches WHERE id = ?`).bind(batchId).first();
  return result || null;
}

async function getBatchSummary(env, batchId) {
  const totals = await env.DB.prepare(`
    SELECT
      COUNT(*) AS rows_staged,
      SUM(CASE WHEN parse_status = 'parsed' THEN 1 ELSE 0 END) AS rows_parsed,
      SUM(CASE WHEN parse_status != 'parsed' THEN 1 ELSE 0 END) AS rows_failed,
      SUM(CASE WHEN TRIM(COALESCE(parse_notes, '')) != '' THEN 1 ELSE 0 END) AS rows_with_warnings,
      COUNT(DISTINCT NULLIF(TRIM(raw_model_number), '')) AS unique_models_in_batch
    FROM staging_schedule_rows
    WHERE batch_id = ?
  `).bind(batchId).first();

  const actions = await env.DB.prepare(`
    SELECT
      COALESCE(SUM(CASE WHEN action = 'inserted' THEN 1 ELSE 0 END), 0) AS catalog_inserts,
      COALESCE(SUM(CASE WHEN action = 'updated' THEN 1 ELSE 0 END), 0) AS catalog_updates,
      COALESCE(SUM(CASE WHEN action = 'unchanged' THEN 1 ELSE 0 END), 0) AS catalog_unchanged,
      COALESCE(SUM(CASE WHEN action = 'duplicate_in_batch' THEN 1 ELSE 0 END), 0) AS duplicates_in_batch
    FROM import_model_results
    WHERE batch_id = ?
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
    SELECT
      raw_model_number AS model_number,
      COUNT(*) AS duplicate_count,
      GROUP_CONCAT(source_row_number) AS source_row_numbers
    FROM staging_schedule_rows
    WHERE batch_id = ? AND NULLIF(TRIM(raw_model_number), '') IS NOT NULL
    GROUP BY raw_model_number
    HAVING COUNT(*) > 1
    ORDER BY duplicate_count DESC, raw_model_number
  `).bind(batchId).all();

  const parseIssues = await env.DB.prepare(`
    SELECT source_row_number, raw_model_number, parse_status, parse_notes
    FROM staging_schedule_rows
    WHERE batch_id = ?
      AND (parse_status != 'parsed' OR TRIM(COALESCE(parse_notes, '')) != '')
    ORDER BY source_row_number
  `).bind(batchId).all();

  const issues = [];
  for (const row of coerceArray(duplicates.results)) {
    issues.push({
      type: 'duplicate_model_in_batch',
      model_number: row.model_number,
      count: toInt(row.duplicate_count),
      source_row_numbers: String(row.source_row_numbers || '')
        .split(',')
        .map((value) => toInt(value))
        .filter(Boolean),
    });
  }

  const warningRows = summarizeIssueRows(coerceArray(parseIssues.results));
  if (warningRows.length) {
    issues.push({ type: 'parse_warnings', rows: warningRows });
  }

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
// Batch GET handlers
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
    SELECT *
    FROM staging_schedule_rows
    WHERE batch_id = ?
    ORDER BY source_row_number, id
  `).bind(batchId).all();
  return json({ ok: true, batch, rows: coerceArray(rows.results) });
}

async function handleGetImportBatchCatalogResults(env, batchId) {
  const batch = await getBatch(env, batchId);
  if (!batch) return json({ error: 'Import batch not found.' }, 404);
  const rows = await env.DB.prepare(`
    SELECT
      imr.id, imr.batch_id, imr.staging_row_id,
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
// Catalog query
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
// Schedule resolution (preview + export)
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
// Export workbook
// ─────────────────────────────────────────────

const COL = {
  tag: 2, areaServed: 3, manufacturer: 4, modelNumber: 5, nominalTons: 6,
  unitType: 7, unitEer: 8, seerIeerr: 9, supplyCfm: 10, supplyEsp: 11,
  supplyQty: 12, supplyBhp: 13, supplyHp: 14, supplyRpm: 15,
  coolingEat: 16, coolingLat: 17, coolingSensible: 18, coolingTotal: 19,
  heatingCfm: 20, heatingEat: 21, heatingLat: 22, heatingInput: 23,
  heatingOutput: 24, voltPh: 25, mca: 26, mocp: 27, weight: 28, remarks: 29,
};

async function getTemplateWorkbook(env) {
  const object = await env.TEMPLATES.get('SSR-Schedule-Example.xlsx');
  if (!object) throw new Error('Template workbook not found in R2 bucket.');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await object.arrayBuffer());
  return workbook;
}

async function createWorkbook(env, units) {
  const workbook = await getTemplateWorkbook(env);
  const worksheet = workbook.getWorksheet('Table 1') || workbook.worksheets[0];
  const templateRowNumber = 4;
  const baseRow = worksheet.getRow(templateRowNumber);
  const rows = await resolveScheduleRows(env, units);

  for (let index = 0; index < rows.length; index += 1) {
    const scheduleRow = rows[index];
    const rowNumber = templateRowNumber + index;
    if (rowNumber > templateRowNumber) worksheet.duplicateRow(templateRowNumber, 1, true);
    const row = worksheet.getRow(rowNumber);
    row.height = baseRow.height;

    row.getCell(COL.tag).value = scheduleRow.tag;
    row.getCell(COL.areaServed).value = scheduleRow.areaServed;
    row.getCell(COL.manufacturer).value = scheduleRow.manufacturer;
    row.getCell(COL.modelNumber).value = scheduleRow.modelNumber;
    row.getCell(COL.nominalTons).value = scheduleRow.nominalTons;
    row.getCell(COL.unitType).value = scheduleRow.unitType;
    row.getCell(COL.unitEer).value = scheduleRow.unitEer;
    row.getCell(COL.seerIeerr).value = scheduleRow.seerIeerr;
    row.getCell(COL.supplyCfm).value = scheduleRow.supplyCfm;
    row.getCell(COL.supplyEsp).value = scheduleRow.supplyEsp;
    row.getCell(COL.supplyQty).value = scheduleRow.supplyQty;
    row.getCell(COL.supplyBhp).value = scheduleRow.supplyBhp;
    row.getCell(COL.supplyHp).value = scheduleRow.supplyHp;
    row.getCell(COL.supplyRpm).value = scheduleRow.supplyRpm;
    row.getCell(COL.coolingEat).value = scheduleRow.coolingEat;
    row.getCell(COL.coolingLat).value = scheduleRow.coolingLat;
    row.getCell(COL.coolingSensible).value = scheduleRow.coolingSensible;
    row.getCell(COL.coolingTotal).value = scheduleRow.coolingTotal;
    row.getCell(COL.heatingCfm).value = scheduleRow.heatingCfm;
    row.getCell(COL.heatingEat).value = scheduleRow.heatingEat;
    row.getCell(COL.heatingLat).value = scheduleRow.heatingLat;
    row.getCell(COL.heatingInput).value = scheduleRow.heatingTotalCapacity || scheduleRow.heatingInput;
    row.getCell(COL.heatingOutput).value = scheduleRow.heatingOutput;
    row.getCell(COL.voltPh).value = scheduleRow.voltPh;
    row.getCell(COL.mca).value = scheduleRow.mca;
    row.getCell(COL.mocp).value = scheduleRow.mocp;
    row.getCell(COL.weight).value = scheduleRow.weight;
    row.getCell(COL.remarks).value = scheduleRow.remarks;
    row.commit();
  }

  return workbook.xlsx.writeBuffer();
}

// ─────────────────────────────────────────────
// Router
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
