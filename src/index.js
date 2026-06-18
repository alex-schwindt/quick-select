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

function normalizeFamily(value) {
  const v = normalizeText(value).toLowerCase();
  if (['heat pump', 'heatpump', 'hp'].includes(v)) return 'Heat Pump';
  if (['ac', 'gas pack', 'gas pkg', 'gas package', 'gaspack'].includes(v)) return 'AC';
  return normalizeText(value);
}

function normalizeEfficiency(value) {
  const v = normalizeText(value).toLowerCase();
  if (['high', 'high efficiency', 'high-efficiency', 'hi'].includes(v)) return 'High';
  return 'Standard';
}

function normalizeVoltage(value) {
  const s = normalizeText(value).toLowerCase().replace(/\s+/g, '').replace(/-/g, '/');
  if (!s) return '';
  if (s.includes('460')) return '460/3';
  if (s.includes('208') || s.includes('230')) return '208/230/3';
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
  const text = normalizeText(value);
  if (!text) return '';
  const num = Number(text.replace(/[^0-9.]/g, ''));
  if (!Number.isFinite(num)) return text;
  if (/kw/i.test(text)) return `${num % 1 === 0 ? String(num.toFixed(0)) : String(num)} kW`;
  if (/mbh/i.test(text)) return `${num % 1 === 0 ? String(num.toFixed(0)) : String(num)} MBH`;
  return text;
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
  return Number.isFinite(n) ? n !== 0 : true;
}

function slugify(value) {
  return String(value ?? '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/(^-|-$)/g, '');
}

function buildSelectionCode(unit) {
  const familyCode = normalizeFamily(unit.family) === 'Heat Pump' ? 'HP' : 'AC';
  const efficiencyCode = normalizeEfficiency(unit.efficiency) === 'High' ? 'HI' : 'STD';
  const tonnageCode = String(unit.tonnage).replace('.', 'P');
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
  return `${familyCode}-${efficiencyCode}-${tonnageCode}-${voltageCode}-${heatCode}-${reheatCode}-${econCode}`;
}

function optionSummary(unit, match) {
  return normalizeText(unit.remarks) || normalizeText(match?.remarks_default) || [
    unit.hotGasReheat ? 'Hot Gas Reheat' : null,
    unit.economizer === 'barometric' ? 'Economizer w/ Barometric Relief' : null,
    unit.economizer === 'powered' ? 'Economizer w/ Powered Exhaust' : null
  ].filter(Boolean).join(', ');
}

function numberOrNull(value) {
  const n = Number(String(value ?? '').replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : null;
}

function xlsxCellValue(sheet, col, row) {
  const addr = XLSX.utils.encode_cell({ c: col - 1, r: row - 1 });
  const cell = sheet[addr];
  if (!cell) return '';
  const v = cell.w !== undefined ? cell.w : cell.v;
  return normalizeText(v);
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

function sheetToMatrix(sheet) {
  const rows = xlsxRowCount(sheet);
  const cols = xlsxColCount(sheet);
  const out = [];
  for (let r = 1; r <= rows; r++) {
    const row = [];
    for (let c = 1; c <= cols; c++) row.push(xlsxCellValue(sheet, c, r));
    out.push(row);
  }
  return out;
}

function looksLikeTypicalCatalogRow(row) {
  const descriptor = normalizeText(row[0]);
  const modelNumber = normalizeText(row[1]);
  return /\b\d+(?:\.\d+)?-?ton\b/i.test(descriptor) && /\b(?:gas|heat pump|hp)\b/i.test(descriptor) && /^[A-Z0-9]{8,}$/i.test(modelNumber);
}

function inferVoltageFromRow(row) {
  for (const cell of row) {
    const text = normalizeText(cell);
    const voltage = normalizeVoltage(text);
    if (voltage) {
      if (/^(208\/230\/3|460\/3)$/.test(voltage)) return voltage;
    }
  }
  return '';
}

function inferCatalogRowFromTypicalLayout(row) {
  const descriptor = normalizeText(row[0]);
  const modelNumber = normalizeText(row[1]);
  if (!descriptor || !modelNumber) return null;

  const tonnageMatch = descriptor.match(/(\d+(?:\.\d+)?)\s*-?\s*Ton/i);
  const gasMatch = descriptor.match(/(\d+(?:\.\d+)?)\s*MBH/i);
  const kwMatch = descriptor.match(/(\d+(?:\.\d+)?)\s*kW/i);
  const isHeatPump = /\bheat pump\b|\bhp\b/i.test(descriptor);
  const isGas = /\bgas\b/i.test(descriptor);
  const isElectric = /\belectric\b/i.test(descriptor);
  const voltage = inferVoltageFromRow(row) || normalizeVoltage(descriptor);
  const coolingCfm = numberOrNull(row[4]);
  const eer = cleanBlank(row[7]);
  const ieer = cleanBlank(row[8]);
  const coolingTotal = cleanBlank(row[20]);
  const refrigerant = cleanBlank(row[14]);
  const mca = cleanBlank(row[26]);
  const mocp = cleanBlank(row[27]);
  const weight = cleanBlank(row[29]);

  const tonnageValue = tonnageMatch ? Number(tonnageMatch[1]) : null;
  if (!tonnageValue) return null;

  let family = 'AC';
  if (isHeatPump) family = 'Heat Pump';

  let heatType = 'None';
  let heatCapacity = '';
  if (isGas && gasMatch) {
    heatType = 'Aluminum Gas Heat';
    heatCapacity = `${Number(gasMatch[1]) % 1 === 0 ? Number(gasMatch[1]).toFixed(0) : gasMatch[1]} MBH`;
  } else if (isElectric && kwMatch) {
    heatType = 'Electric Heat';
    heatCapacity = `${Number(kwMatch[1]) % 1 === 0 ? Number(kwMatch[1]).toFixed(0) : kwMatch[1]} kW`;
  }

  return {
    family,
    efficiency: 'Standard',
    tonnage: tonnageValue,
    voltage,
    heat_type: heatType,
    heat_capacity: heatCapacity,
    model_number: modelNumber,
    model_code: modelNumber,
    unit_type: family === 'Heat Pump' ? 'Packaged Heat Pump' : 'Packaged AC',
    unit_eer: eer,
    seer_ieer: ieer,
    cooling_cfm: coolingCfm,
    cooling_total_capacity_mbh: coolingTotal,
    heating_capacity_mbtu: heatCapacity,
    refrigerant_type: refrigerant,
    mca,
    mocp,
    operating_weight_lbs: weight,
    source_descriptor: descriptor
  };
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

async function upsertCatalogRow(env, row) {
  const existing = await env.DB.prepare(`SELECT id FROM unit_models WHERE family = ? AND efficiency = ? AND tonnage = CAST(? AS REAL) AND voltage = ? AND heat_type = ? AND IFNULL(heat_capacity, '') = IFNULL(?, '') AND model_number = ? LIMIT 1`)
    .bind(row.family, row.efficiency, row.tonnage, row.voltage, row.heat_type, row.heat_capacity, row.model_number)
    .first();

  if (existing?.id) {
    await env.DB.prepare(`UPDATE unit_models SET model_code = ?, unit_type = ?, unit_eer = ?, seer_ieer = ?, cooling_cfm = ?, cooling_total_capacity_mbh = ?, heating_capacity_mbtu = ?, refrigerant_type = ?, mca = ?, mocp = ?, operating_weight_lbs = ? WHERE id = ?`)
      .bind(
        cleanBlank(row.model_code),
        cleanBlank(row.unit_type),
        cleanBlank(row.unit_eer),
        cleanBlank(row.seer_ieer),
        row.cooling_cfm,
        cleanBlank(row.cooling_total_capacity_mbh),
        cleanBlank(row.heating_capacity_mbtu),
        cleanBlank(row.refrigerant_type),
        cleanBlank(row.mca),
        cleanBlank(row.mocp),
        cleanBlank(row.operating_weight_lbs),
        existing.id
      )
      .run();
    await ensureDocuments(env, existing.id, row.model_number);
    return 'updated';
  }

  const insert = await env.DB.prepare(`INSERT INTO unit_models (family, efficiency, tonnage, voltage, heat_type, heat_capacity, model_code, model_number, unit_type, unit_eer, seer_ieer, cooling_cfm, cooling_total_capacity_mbh, heating_capacity_mbtu, refrigerant_type, mca, mocp, operating_weight_lbs) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`)
    .bind(
      row.family,
      row.efficiency,
      row.tonnage,
      row.voltage,
      row.heat_type,
      cleanBlank(row.heat_capacity) || '',
      cleanBlank(row.model_code),
      row.model_number,
      cleanBlank(row.unit_type),
      cleanBlank(row.unit_eer),
      cleanBlank(row.seer_ieer),
      row.cooling_cfm,
      cleanBlank(row.cooling_total_capacity_mbh),
      cleanBlank(row.heating_capacity_mbtu),
      cleanBlank(row.refrigerant_type),
      cleanBlank(row.mca),
      cleanBlank(row.mocp),
      cleanBlank(row.operating_weight_lbs)
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
  const issues = [];
  let rowsRead = 0;
  let inserted = 0;
  let updated = 0;

  if (name.endsWith('.csv')) {
    const text = await file.text();
    const wb = XLSX.read(text, { type: 'string' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const records = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    for (const rec of records) {
      const modelNumber = normalizeText(rec.model_number || rec['Model Number']);
      if (!modelNumber) continue;
      rowsRead += 1;
      const row = {
        family: normalizeFamily(rec.family || rec['Family']),
        efficiency: normalizeEfficiency(rec.efficiency || rec['Efficiency']),
        tonnage: Number(rec.tonnage || rec['Tonnage']),
        voltage: normalizeVoltage(rec.voltage || rec['Voltage']),
        heat_type: normalizeHeatType(rec.heat_type || rec['Heat Type']),
        heat_capacity: normalizeHeatCapacity(rec.heat_capacity || rec['Heat Capacity']),
        model_number: modelNumber,
        model_code: normalizeText(rec.model_code || rec['Model Code'] || modelNumber),
        unit_type: normalizeText(rec.unit_type || rec['Unit Type']),
        unit_eer: cleanBlank(rec.unit_eer || rec['Unit EER']),
        seer_ieer: cleanBlank(rec.seer_ieer || rec['SEER/IEER']),
        cooling_cfm: numberOrNull(rec.cooling_cfm || rec['Cooling CFM']),
        cooling_total_capacity_mbh: cleanBlank(rec.cooling_total_capacity_mbh || rec['Cooling Total Capacity MBH']),
        heating_capacity_mbtu: cleanBlank(rec.heating_capacity_mbtu || rec['Heating Capacity MBTU']),
        refrigerant_type: cleanBlank(rec.refrigerant_type || rec['Refrigerant Type']),
        mca: cleanBlank(rec.mca || rec['MCA']),
        mocp: cleanBlank(rec.mocp || rec['MOCP']),
        operating_weight_lbs: cleanBlank(rec.operating_weight_lbs || rec['Operating Weight Lbs'])
      };
      const action = await upsertCatalogRow(env, row);
      if (action === 'inserted') inserted += 1;
      if (action === 'updated') updated += 1;
    }
    return json({ ok: true, rows_read: rowsRead, inserted, updated, issues });
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellText: true, cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) return json({ error: 'No worksheet found in workbook.' }, 400);
  const matrix = sheetToMatrix(sheet);

  const parsedRows = [];
  for (let i = 0; i < matrix.length; i += 1) {
    const row = matrix[i];
    if (!looksLikeTypicalCatalogRow(row)) continue;
    const parsed = inferCatalogRowFromTypicalLayout(row);
    if (!parsed) {
      issues.push(`Unable to parse row ${i + 1}.`);
      continue;
    }
    if (!parsed.voltage) issues.push(`Row ${i + 1} missing recognizable voltage for model ${parsed.model_number}.`);
    parsedRows.push({ ...parsed, source_row_number: i + 1 });
  }

  for (const row of parsedRows) {
    rowsRead += 1;
    try {
      const action = await upsertCatalogRow(env, row);
      if (action === 'inserted') inserted += 1;
      if (action === 'updated') updated += 1;
    } catch (error) {
      issues.push(`Row ${row.source_row_number} (${row.model_number}): ${error.message}`);
    }
  }

  return json({ ok: true, rows_read: rowsRead, inserted, updated, issues });
}

async function findMatchingImportedRow(env, unit) {
  const family = normalizeFamily(unit.family);
  const efficiency = normalizeEfficiency(unit.efficiency);
  const voltage = normalizeVoltage(unit.voltage);
  const heatType = normalizeHeatType(unit.heatType);
  const heatCapacity = normalizeHeatCapacity(unit.heatCapacity);
  const tonnage = Number(unit.tonnage);

  const exact = await env.DB.prepare(`SELECT um.*, ud.cutsheet_url, ud.accessories_url, ud.wiring_url, ud.iom_url FROM unit_models um LEFT JOIN unit_documents ud ON ud.model_id = um.id WHERE um.family = ? AND um.efficiency = ? AND um.tonnage = CAST(? AS REAL) AND um.voltage = ? AND um.heat_type = ? AND IFNULL(um.heat_capacity, '') = IFNULL(?, '') ORDER BY um.id LIMIT 1`)
    .bind(family, efficiency, tonnage, voltage, heatType, heatCapacity)
    .first();
  if (exact) return exact;

  const relaxed = await env.DB.prepare(`SELECT um.*, ud.cutsheet_url, ud.accessories_url, ud.wiring_url, ud.iom_url FROM unit_models um LEFT JOIN unit_documents ud ON ud.model_id = um.id WHERE um.family = ? AND um.tonnage = CAST(? AS REAL) AND um.voltage = ? ORDER BY um.id LIMIT 1`)
    .bind(family, tonnage, voltage)
    .first();
  return relaxed || null;
}

function previewFallback(unit) {
  return {
    tag: unit.tag,
    areaServed: unit.areaServed || '—',
    manufacturer: 'H&H Trecho',
    modelNumber: buildSelectionCode(unit),
    nominalTons: unit.tonnage,
    unitType: unit.family,
    supplyCfm: '—',
    coolingTotal: '—',
    heatingTotalCapacity: '—',
    heatingInput: unit.heatType === 'None' ? 'No heat' : `${unit.heatType} ${unit.heatCapacity}`,
    voltPh: unit.voltage,
    mca: '—',
    mocp: '—',
    weight: '—',
    remarks: optionSummary(unit),
    matchFound: false,
    cutsheetUrl: '',
    accessoriesUrl: '',
    wiringUrl: '',
    iomUrl: ''
  };
}

function buildPreviewRow(unit, match) {
  if (!match) return previewFallback(unit);
  return {
    tag: unit.tag,
    areaServed: unit.areaServed || '—',
    manufacturer: 'Tempmaster',
    modelNumber: match.model_number || buildSelectionCode(unit),
    nominalTons: unit.tonnage,
    unitType: match.unit_type || unit.family,
    supplyCfm: asBlank(match.cooling_cfm) || '—',
    coolingTotal: asBlank(match.cooling_total_capacity_mbh) || '—',
    heatingTotalCapacity: asBlank(match.heating_capacity_mbtu) || '—',
    heatingInput: unit.heatType === 'None' ? 'No heat' : `${unit.heatType} ${unit.heatCapacity}`,
    voltPh: match.voltage || unit.voltage,
    mca: asBlank(match.mca) || '—',
    mocp: asBlank(match.mocp) || '—',
    weight: asBlank(match.operating_weight_lbs) || '—',
    remarks: optionSummary(unit, match) || '—',
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

async function createWorkbook(env, units) {
  const rows = [];
  for (const unit of units) {
    const match = await findMatchingImportedRow(env, unit);
    const row = buildPreviewRow(unit, match);
    rows.push({
      Tag: row.tag,
      'Area Served': row.areaServed,
      Family: unit.family,
      'Model Number': row.modelNumber,
      Tons: row.nominalTons,
      'Unit Type': row.unitType,
      'Volt/Ph': row.voltPh,
      'Supply Fan CFM': row.supplyCfm,
      'Cooling Total': row.coolingTotal,
      'Heating Total Capacity': row.heatingTotalCapacity,
      'Heating Input': row.heatingInput,
      MCA: row.mca,
      MOCP: row.mocp,
      Weight: row.weight,
      Remarks: row.remarks
    });
  }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'Schedule');
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
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