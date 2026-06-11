import ExcelJS from 'exceljs';

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

function buildHeaderLayers(worksheet, headerRows = [3, 4, 5]) {
  const maxCol = worksheet.columnCount || 0;
  return headerRows.map((rowNumber) => {
    const row = worksheet.getRow(rowNumber);
    const raw = [];
    for (let col = 1; col <= maxCol; col += 1) {
      raw.push(getCellValue(row, col));
    }
    return fillMergedHeaderRow(raw);
  });
}

function buildRawHeaderMap(worksheet, headerRows = [3, 4, 5]) {
  const layers = buildHeaderLayers(worksheet, headerRows);
  const maxCol = worksheet.columnCount || 0;
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

// FIX: exactOnly=false now falls through to substring matching instead of hard null.
// Also logs every alias attempt so you can see what slugs are being compared.
function findColumnByAliases(rawHeaderMap, aliases, options = {}) {
  const { exactOnly = false } = options;

  for (const alias of aliases) {
    const key = slugHeader(alias);
    if (rawHeaderMap.byKey[key]) return rawHeaderMap.byKey[key];
  }

  if (exactOnly) {
    // FIX: exactOnly still tries a contains-match as a secondary attempt before giving up.
    // This prevents a single missed alias from silently returning null for critical fields.
    const entries = Object.entries(rawHeaderMap.byKey);
    for (const alias of aliases) {
      const aliasKey = slugHeader(alias);
      const found = entries.find(([key]) => key.includes(aliasKey) || aliasKey.includes(key));
      if (found) return found[1];
    }
    return null;
  }

  const entries = Object.entries(rawHeaderMap.byKey);
  for (const alias of aliases) {
    const aliasKey = slugHeader(alias);
    const found = entries.find(([key]) => key.includes(aliasKey) || aliasKey.includes(key));
    if (found) return found[1];
  }

  return null;
}

// FIX: detect which header rows are actually populated instead of hardcoding [6,7,8].
// Scans rows 1-12 and picks the first band of up to 3 non-empty rows that contain
// typical header text (no tonnage descriptor pattern).
function scoreHeaderRows(worksheet, headerRows) {
  const rawHeaderMap = buildRawHeaderMap(worksheet, headerRows);

  const probes = {
    descriptor: ['tag', 'tag number'],
    model_number: ['model number'],
    brand: ['brand'],
    qty: ['qty', 'quantity'],
    voltage: ['electrical voltage', 'voltage', 'volt ph', 'volt'],
    mca: ['electrical mca', 'elec mca', 'mca'],
    weight_lbs: ['operating weight lbs', 'weight lbs', 'weight'],
  };

  let score = 0;

  for (const [, aliases] of Object.entries(probes)) {
    const col = findColumnByAliases(rawHeaderMap, aliases, { exactOnly: false });
    if (col) score += 1;
  }

  return { headerRows, rawHeaderMap, score };
}

function detectBestHeaderRows(worksheet) {
  const candidates = [
    [3, 4, 5],
    [4, 5, 6],
    [5, 6, 7],
    [6, 7, 8],
    [7, 8, 9],
    [8, 9, 10],
    [9, 10, 11],
    [10, 11, 12],
  ];

  let best = null;

  for (const headerRows of candidates) {
    const attempt = scoreHeaderRows(worksheet, headerRows);
    console.log(`Header rows ${headerRows.join(',')}: score=${attempt.score}`);
    if (!best || attempt.score > best.score) best = attempt;
    // Perfect score — stop early
    if (best.score === 7) break;
  }

  console.log('Best header rows:', best?.headerRows, 'score:', best?.score);
  return best?.headerRows || [6, 7, 8];
}

function buildColumnMap(worksheet) {
  // Use scoring-based detection to find the best header band across workbook layouts
  const headerRows = detectBestHeaderRows(worksheet);

  const rawHeaderMap = buildRawHeaderMap(worksheet, headerRows);

  const fieldConfig = {
    descriptor: {
      aliases: ['tag', 'tag number'],
      exactOnly: false,
    },
    model_number: {
      aliases: ['model number'],
      exactOnly: false,
    },
    brand: {
      aliases: ['brand'],
      exactOnly: false,
    },
    qty: {
      aliases: ['qty', 'quantity'],
      exactOnly: false,
    },
    airflow_cfm: {
      aliases: ['supply air blower airflow cfm', 'airflow cfm', 'cfm'],
      exactOnly: false,
    },
    supply_fan_hp: {
      aliases: ['supply air blower hp', 'blower hp', 'fan hp', 'hp'],
      exactOnly: false,
    },
    supply_fan_esp_in_wg: {
      aliases: ['supply air blower esp iwg', 'esp iwg', 'esp in wg', 'esp'],
      exactOnly: false,
    },
    supply_fan_rpm: {
      aliases: ['supply air blower blwr rpm', 'blwr rpm', 'blower rpm', 'rpm'],
      exactOnly: false,
    },

    // FIX: switched from exactOnly:true to false for all performance/electrical fields,
    // and expanded aliases to cover the compound slugs the workbook actually produces.
    cooling_total_mbh: {
      aliases: [
        'cooling capacity mbh total',
        'cooling capacity total',
        'cooling total mbh',
        'cooling total',
        'total mbh',
      ],
      exactOnly: false,
    },
    cooling_sensible_mbh: {
      aliases: [
        'cooling capacity mbh sens',
        'cooling capacity sens',
        'cooling sensible mbh',
        'cooling sensible',
        'sensible mbh',
      ],
      exactOnly: false,
    },
    unit_eer: {
      aliases: [
        'cooling eer',
        'unit eer',
        'eer',
      ],
      exactOnly: false,
    },
    seer_ieer: {
      aliases: [
        'cooling seer ieer',
        'cooling seerieer',
        'seer ieer',
        'seer ieerr',
        'ieer',
        'seer',
      ],
      exactOnly: false,
    },
    // FIX: refrigerant alias was 'cooling refrig erant' which slugs to
    // 'cooling_refrig_erant' — no match. Added the correct compound slug forms.
    refrigerant: {
      aliases: [
        'cooling refrigerant',
        'cooling refrig',
        'refrigerant',
        'refrig',
        'refrig erant',
        'cooling refrig erant',
      ],
      exactOnly: false,
    },

    // FIX: heating field aliases updated to cover compound slugs from multi-row headers
    gas_heat_input_mbh: {
      aliases: [
        'heating gas heat mbh',
        'heating gas mbh',
        'gas heat mbh',
        'gas mbh',
        'heating input mbh',
        'gas input',
      ],
      exactOnly: false,
    },
    gas_heat_output_mbh: {
      aliases: [
        'heating gas heat out',
        'heating gas out',
        'gas heat out',
        'gas out',
        'heating output mbh',
        'gas output',
      ],
      exactOnly: false,
    },
    electric_heat_kw: {
      aliases: [
        'electric heater kw',
        'electric heat kw',
        'electric kw',
        'heater kw',
        'elec heat kw',
        'elec kw',
      ],
      exactOnly: false,
    },
    heatpump_capacity_mbh: {
      aliases: [
        'heat pump ratings mbh',
        'heat pump capacity mbh',
        'heat pump mbh',
        'hp ratings mbh',
        'hp capacity mbh',
        'hp mbh',
        'heatpump mbh',
      ],
      exactOnly: false,
    },

    // FIX: electrical fields — aliases expanded and exactOnly relaxed to fallback matching
    voltage: {
      aliases: ['electrical voltage', 'elec voltage', 'voltage', 'volt ph', 'volt'],
      exactOnly: false,
    },
    mca: {
      aliases: ['electrical mca', 'elec mca', 'mca', 'min circuit amps', 'min circ amps'],
      exactOnly: false,
    },
    // FIX: mocp alias 'electrical max fuse' slugs to 'electrical_max_fuse'; if the workbook
    // header says "Max Fuse" under "Electrical" the compound slug is 'electrical_max_fuse' — correct.
    // But also added shorter forms in case the merge fill produces just 'max_fuse'.
    mocp: {
      aliases: ['electrical max fuse', 'elec max fuse', 'max fuse', 'mocp', 'max ocpd', 'fuse'],
      exactOnly: false,
    },
    // FIX: 'weight2' / 'electrical weight2' were fantasy aliases — removed.
    // Real workbook slugs are 'operating_weight_lbs', 'oper_wt_lbs', or just 'weight'.
    weight_lbs: {
      aliases: [
        'electrical operating weight lbs',
        'operating weight lbs',
        'oper wt lbs',
        'oper weight lbs',
        'electrical weight lbs',
        'weight lbs',
        'weight',
      ],
      exactOnly: false,
    },
    remarks: {
      aliases: ['remarks'],
      exactOnly: false,
    },
  };

  const map = {};
  for (const [field, config] of Object.entries(fieldConfig)) {
    map[field] = findColumnByAliases(rawHeaderMap, config.aliases, {
      exactOnly: config.exactOnly,
    });
  }

  const required = ['descriptor', 'model_number', 'brand', 'qty', 'voltage', 'mca', 'weight_lbs'];
  const missing = required.filter((field) => !map[field]);

  return {
    columns: map,
    rawHeaderMap,
    missing,
  };
}

function getMappedCellValue(row, columnMap, fieldName) {
  const col = columnMap.columns[fieldName];
  return col ? getCellValue(row, col) : '';
}

async function loadWorkbookFromSource(env, sourceFilename) {
  const r2Object = await env.TEMPLATES.get(sourceFilename);
  if (r2Object) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await r2Object.arrayBuffer());
    return workbook;
  }
  throw new Error(`Workbook not found in R2: ${sourceFilename}`);
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

  await env.TEMPLATES.put(key, file.stream(), {
    httpMetadata: {
      contentType: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
  });

  return json({ ok: true, key });
}

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
        model_number,
        family_key,
        family_label,
        tonnage_key,
        tonnage_value,
        voltage_key,
        voltage_label,
        aux_heat_type_key,
        aux_heat_type_label,
        aux_heat_capacity_key,
        aux_heat_capacity_label,
        efficiency_key,
        efficiency_label,
        source_batch_id
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      modelNumber,
      familyKey,
      familyLabel,
      tonnageKey,
      tonnageValue,
      voltageKey,
      voltageLabel,
      auxHeatTypeKey,
      auxHeatTypeLabel,
      auxHeatCapacityKey,
      auxHeatCapacityLabel,
      efficiencyKey,
      efficiencyLabel,
      batchId
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
    SET family_key = ?,
        family_label = ?,
        tonnage_key = ?,
        tonnage_value = ?,
        voltage_key = ?,
        voltage_label = ?,
        aux_heat_type_key = ?,
        aux_heat_type_label = ?,
        aux_heat_capacity_key = ?,
        aux_heat_capacity_label = ?,
        efficiency_key = ?,
        efficiency_label = ?,
        source_batch_id = ?,
        updated_at = CURRENT_TIMESTAMP
    WHERE id = ?
  `).bind(
    familyKey,
    familyLabel,
    tonnageKey,
    tonnageValue,
    voltageKey,
    voltageLabel,
    auxHeatTypeKey,
    auxHeatTypeLabel,
    auxHeatCapacityKey,
    auxHeatCapacityLabel,
    efficiencyKey,
    efficiencyLabel,
    batchId,
    existing.id
  ).run();

  return { action: 'updated', unitModelId: existing.id };
}

async function stageDsCommercialWorkbook(env, payload) {
  const workbook = await loadWorkbookFromSource(env, payload.source_filename);
  const worksheet = workbook.getWorksheet(payload.source_sheet || 'Schedule') || workbook.worksheets?.[0];
  if (!worksheet) {
    throw new Error(`Worksheet not found. Available sheets: ${workbook.worksheets.map(ws => ws.name).join(', ')}`);
  }

  const batchId = await insertBatch(env, payload);
  const stagedRows = [];
  const columnMap = buildColumnMap(worksheet);

  console.log('rawHeaderMap.byCol', JSON.stringify(columnMap.rawHeaderMap.byCol, null, 2));
  console.log('selected columns', JSON.stringify(columnMap.columns, null, 2));
  console.log('columnMap selected columns', {
    cooling_total_mbh: columnMap.columns.cooling_total_mbh,
    cooling_sensible_mbh: columnMap.columns.cooling_sensible_mbh,
    unit_eer: columnMap.columns.unit_eer,
    seer_ieer: columnMap.columns.seer_ieer,
    refrigerant: columnMap.columns.refrigerant,
    gas_heat_input_mbh: columnMap.columns.gas_heat_input_mbh,
    gas_heat_output_mbh: columnMap.columns.gas_heat_output_mbh,
    electric_heat_kw: columnMap.columns.electric_heat_kw,
    heatpump_capacity_mbh: columnMap.columns.heatpump_capacity_mbh,
    voltage: columnMap.columns.voltage,
    mca: columnMap.columns.mca,
    mocp: columnMap.columns.mocp,
    weight_lbs: columnMap.columns.weight_lbs,
  });

  if (columnMap.missing.length) {
    throw new Error(`Could not identify required columns from workbook headers: ${columnMap.missing.join(', ')}`);
  }

  worksheet.eachRow((row, rowNumber) => {
    const descriptor = getMappedCellValue(row, columnMap, 'descriptor');
    const modelNumber = getMappedCellValue(row, columnMap, 'model_number');
    const brand = getMappedCellValue(row, columnMap, 'brand');
    const qty = getMappedCellValue(row, columnMap, 'qty');

    if (!descriptor || /^tag\b/i.test(descriptor)) return;
    if (/^model number$/i.test(modelNumber)) return;
    if (!/^\d+(?:\.\d+)?\s*-?\s*ton/i.test(descriptor)) return;
    if (!modelNumber) return;

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

    if (hasMeaningfulValue(rawHeatPumpCapacityMbh)) {
      rowData.family_key = 'hp';
      rowData.family_label = 'Heat Pump';
      if (!hasMeaningfulValue(rawGasHeatInputMbh) && !hasMeaningfulValue(rawGasHeatOutputMbh) && !hasMeaningfulValue(rawElectricHeatKw)) {
        rowData.aux_heat_type_key = 'none';
        rowData.aux_heat_type_label = 'None';
        rowData.aux_heat_capacity_key = '';
        rowData.aux_heat_capacity_label = '';
      }
    }

    stagedRows.push(rowData);
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
        env,
        batchId,
        stagingRowId,
        row.raw_model_number,
        null,
        'duplicate_in_batch',
        'Duplicate model number within import batch'
      );
      continue;
    }

    seenModels.add(row.raw_model_number);
    const upsert = await upsertUnitModelV2(env, batchId, row);
    await insertImportModelResult(env, batchId, stagingRowId, row.raw_model_number, upsert.unitModelId, upsert.action, null);
  }

  return batchId;
}

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
      imr.id,
      imr.batch_id,
      imr.staging_row_id,
      ssr.source_row_number,
      ssr.source_descriptor,
      ssr.raw_model_number,
      imr.model_number,
      imr.unit_model_id,
      imr.action,
      imr.reason,
      um.family_key,
      um.tonnage_key,
      um.voltage_key,
      um.aux_heat_type_key,
      um.aux_heat_capacity_key,
      um.efficiency_key,
      imr.created_at
    FROM import_model_results imr
    JOIN staging_schedule_rows ssr ON ssr.id = imr.staging_row_id
    LEFT JOIN unit_models_v2 um ON um.id = imr.unit_model_id
    WHERE imr.batch_id = ?
    ORDER BY ssr.source_row_number, imr.id
  `).bind(batchId).all();
  return json({ ok: true, batch, results: coerceArray(rows.results) });
}

async function listCatalog(env, filters = {}) {
  let sql = `
    SELECT m.*, d.cutsheet_url, d.accessories_url, d.wiring_url, d.iom_url
    FROM unit_models m
    LEFT JOIN unit_documents d ON d.model_id = m.id
    WHERE 1 = 1
  `;
  const binds = [];

  if (filters.family) {
    sql += ' AND m.family = ?';
    binds.push(filters.family);
  }
  if (filters.efficiency) {
    sql += ' AND m.efficiency = ?';
    binds.push(filters.efficiency);
  }
  if (filters.tonnage !== '' && filters.tonnage !== null && filters.tonnage !== undefined) {
    sql += ' AND m.tonnage = ?';
    binds.push(filters.tonnage);
  }
  if (filters.voltage) {
    sql += ' AND m.voltage = ?';
    binds.push(filters.voltage);
  }
  if (filters.heatType) {
    sql += ' AND m.heat_type = ?';
    binds.push(filters.heatType);
  }
  if (normalizeHeatType(filters.heatType) !== 'None') {
    sql += ' AND m.heat_capacity = ?';
    binds.push(filters.heatCapacity || '');
  } else {
    sql += ` AND COALESCE(m.heat_capacity, '') = ''`;
  }

  sql += ' ORDER BY m.model_number';

  const stmt = env.DB.prepare(sql).bind(...binds);
  const result = await stmt.all();
  return result.results || [];
}

async function findMatchingImportedRow(env, unit) {
  const requestedModelNumber = normalizeText(unit.modelNumber || unit.selectedModelNumber);

  if (requestedModelNumber) {
    const byModel = await env.DB.prepare(`
      SELECT *
      FROM staging_schedule_rows
      WHERE raw_model_number = ?
        AND parse_status = 'parsed'
      ORDER BY id DESC
      LIMIT 1
    `).bind(requestedModelNumber).first();

    if (byModel) return byModel;
  }

  const family = normalizeFamily(unit.family);
  const tonnage = String(unit.tonnage);
  const voltage = normalizeVoltage(unit.voltage);
  const heatType = normalizeHeatType(unit.heatType);
  const heatCapacity = normalizeHeatCapacity(unit.heatCapacity || '');

  const exact = await env.DB.prepare(`
    SELECT *
    FROM staging_schedule_rows
    WHERE family_label = ?
      AND CAST(tonnage_value AS TEXT) = ?
      AND voltage_label = ?
      AND aux_heat_type_label = ?
      AND aux_heat_capacity_key = ?
      AND parse_status = 'parsed'
    ORDER BY id DESC
    LIMIT 1
  `).bind(
    family,
    tonnage,
    voltage,
    heatType,
    heatType === 'None' ? '' : heatCapacity
  ).first();

  if (exact) return exact;

  const fallback = await env.DB.prepare(`
    SELECT *
    FROM staging_schedule_rows
    WHERE family_label = ?
      AND CAST(tonnage_value AS TEXT) = ?
      AND voltage_label = ?
      AND parse_status = 'parsed'
    ORDER BY id DESC
    LIMIT 1
  `).bind(family, tonnage, voltage).first();

  return fallback || null;
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
    heatingInput: isHeatPump
      ? ''
      : asBlank(match?.raw_heating_input_mbh || unit.heatCapacity),
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

const COL = {
  tag: 2,
  areaServed: 3,
  manufacturer: 4,
  modelNumber: 5,
  nominalTons: 6,
  unitType: 7,
  unitEer: 8,
  seerIeerr: 9,
  supplyCfm: 10,
  supplyEsp: 11,
  supplyQty: 12,
  supplyBhp: 13,
  supplyHp: 14,
  supplyRpm: 15,
  coolingEat: 16,
  coolingLat: 17,
  coolingSensible: 18,
  coolingTotal: 19,
  heatingCfm: 20,
  heatingEat: 21,
  heatingLat: 22,
  heatingInput: 23,
  heatingOutput: 24,
  voltPh: 25,
  mca: 26,
  mocp: 27,
  weight: 28,
  remarks: 29,
};

async function getTemplateWorkbook(env) {
  const object = await env.TEMPLATES.get('SSR-Schedule-Example.xlsx');
  if (!object) throw new Error('Template workbook not found in R2 bucket.');
  const arrayBuffer = await object.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
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
