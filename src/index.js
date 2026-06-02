import ExcelJS from 'exceljs';

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json' }
  });
}

function normalizeText(value) {
  return String(value ?? '').trim();
}

function asBlank(value) {
  return value === null || value === undefined || value === '' ? '' : value;
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  const n = Number(String(value).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : null;
}

function normalizeFamily(value) {
  const v = normalizeText(value).toLowerCase();
  if (['ac', 'gas pack', 'gaspack'].includes(v)) return 'ac';
  if (['heat pump', 'heatpump', 'hp'].includes(v)) return 'hp';
  return v.replace(/\s+/g, '-');
}

function normalizeEfficiency(value) {
  const v = normalizeText(value).toLowerCase();
  if (['standard', 'std'].includes(v)) return 'std';
  if (['high', 'high efficiency', 'high-efficiency'].includes(v)) return 'high';
  return v.replace(/\s+/g, '-');
}

function normalizeVoltage(value) {
  const compact = normalizeText(value).toLowerCase().replace(/\s+/g, '').replace(/v/g, '');
  if (['208/3', '208-3', '2083', '208/230/3', '2082303'].includes(compact.replace('/230', ''))) return '208-3';
  if (['208/230/3', '2082303'].includes(compact)) return '208-3';
  if (['460/3', '460-3', '4603'].includes(compact)) return '460-3';
  return compact.replace(/\//g, '-');
}

function normalizeHeatType(value) {
  const v = normalizeText(value).toLowerCase();
  if (['aluminum gas heat', 'gas heat', 'gas'].includes(v)) return 'gas';
  if (['electric heat', 'electric'].includes(v)) return 'electric';
  if (['none', 'no heat', ''].includes(v)) return 'none';
  return v.replace(/\s+/g, '-');
}

function normalizeHeatCapacity(value) {
  const v = normalizeText(value).toLowerCase();
  if (!v) return '0';
  return v.replace(/\s+/g, '').replace('mbh', '').replace('kw', '').replace('.0', '');
}

function normalizeTonnage(value) {
  return String(value ?? '').trim().replace('.0', '');
}

function parseDescriptor(text) {
  const raw = normalizeText(text);
  const lower = raw.toLowerCase();

  const tonnageMatch = raw.match(/(\d+(?:\.\d+)?)\s*[- ]?ton/i);
  const tonnageValue = tonnageMatch ? Number(tonnageMatch[1]) : null;
  const tonnageKey = tonnageValue !== null ? normalizeTonnage(tonnageValue) : '';

  const voltageMatch = raw.match(/(208(?:\/230)?\/3|460\/3|208-3-60|460-3-60)/i);
  const voltageLabel = voltageMatch ? voltageMatch[1].replace('-60', '') : '';
  const voltageKey = normalizeVoltage(voltageLabel);

  const isHeatPump = lower.includes('heat pump') || lower.includes(' hp ');
  const familyKey = isHeatPump ? 'hp' : 'ac';
  const familyLabel = isHeatPump ? 'Heat Pump' : 'AC';

  let auxHeatTypeKey = 'none';
  let auxHeatTypeLabel = 'None';
  let auxHeatCapacityLabel = '';
  let auxHeatCapacityKey = '0';

  if (lower.includes('gas')) {
    auxHeatTypeKey = 'gas';
    auxHeatTypeLabel = 'Aluminum Gas Heat';
    const gasMatch = raw.match(/(\d+(?:\.\d+)?)\s*mbh/i);
    if (gasMatch) {
      auxHeatCapacityLabel = `${gasMatch[1]} MBH`;
      auxHeatCapacityKey = normalizeHeatCapacity(auxHeatCapacityLabel);
    }
  } else if (lower.includes('kw') || lower.includes('electric')) {
    auxHeatTypeKey = 'electric';
    auxHeatTypeLabel = 'Electric Heat';
    const kwMatch = raw.match(/(\d+(?:\.\d+)?)\s*kw/i);
    if (kwMatch) {
      auxHeatCapacityLabel = `${kwMatch[1]} kW`;
      auxHeatCapacityKey = normalizeHeatCapacity(auxHeatCapacityLabel);
    }
  }

  return {
    familyKey,
    familyLabel,
    efficiencyKey: 'std',
    efficiencyLabel: 'Standard',
    tonnageKey,
    tonnageValue,
    voltageKey,
    voltageLabel,
    auxHeatTypeKey,
    auxHeatTypeLabel,
    auxHeatCapacityKey,
    auxHeatCapacityLabel
  };
}

async function listCatalog(env, filters) {
  const familyKey = normalizeFamily(filters.family);
  const efficiencyKey = normalizeEfficiency(filters.efficiency);
  const tonnageKey = normalizeTonnage(filters.tonnage);
  const voltageKey = normalizeVoltage(filters.voltage);
  const auxHeatTypeKey = normalizeHeatType(filters.heatType);
  const auxHeatCapacityKey = normalizeHeatCapacity(filters.heatCapacity);

  let sql = `
SELECT m.*, d.cutsheet_url, d.accessories_url
FROM unit_models_v2 m
LEFT JOIN unit_documents d ON d.model_id = m.id
WHERE 1 = 1
`;
  const binds = [];
  if (filters.family) { sql += ' AND m.family_key = ?'; binds.push(familyKey); }
  if (filters.efficiency) { sql += ' AND m.efficiency_key = ?'; binds.push(efficiencyKey); }
  if (filters.tonnage) { sql += ' AND m.tonnage_key = ?'; binds.push(tonnageKey); }
  if (filters.voltage) { sql += ' AND m.voltage_key = ?'; binds.push(voltageKey); }
  if (filters.heatType) { sql += ' AND m.aux_heat_type_key = ?'; binds.push(auxHeatTypeKey); }
  if (filters.heatCapacity || auxHeatTypeKey !== 'none') { sql += ' AND m.aux_heat_capacity_key = ?'; binds.push(auxHeatCapacityKey); }
  sql += ' ORDER BY m.family_label, m.tonnage_value, m.model_number';
  const stmt = env.DB.prepare(sql).bind(...binds);
  const result = await stmt.all();
  return result.results || [];
}

async function findMatchingModel(env, unit) {
  const matches = await listCatalog(env, {
    family: unit.family,
    efficiency: unit.efficiency,
    tonnage: unit.tonnage,
    voltage: unit.voltage,
    heatType: unit.heatType,
    heatCapacity: unit.heatCapacity || ''
  });
  return matches[0] || null;
}

function optionSummary(unit, match) {
  return normalizeText(unit.remarks) || normalizeText(match?.remarks_default) || [
    unit.hotGasReheat ? 'Hot Gas Reheat' : null,
    unit.economizer === 'barometric' ? 'Economizer w/ Barometric Relief' : null,
    unit.economizer === 'powered' ? 'Economizer w/ Powered Exhaust' : null,
    unit.curb ? 'Curb' : null
  ].filter(Boolean).join(', ');
}

function joinSlash(a, b) {
  const parts = [asBlank(a), asBlank(b)].filter(v => v !== '');
  return parts.length ? parts.join(' / ') : '';
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

  for (let index = 0; index < units.length; index += 1) {
    const unit = units[index];
    const rowNumber = templateRowNumber + index;
    if (rowNumber > templateRowNumber) worksheet.duplicateRow(templateRowNumber, 1, true);
    const row = worksheet.getRow(rowNumber);
    const match = await findMatchingModel(env, unit);

    row.height = baseRow.height;
    row.getCell(COL.tag).value = normalizeText(unit.tag) || `RTU-${index + 1}`;
    row.getCell(COL.areaServed).value = asBlank(unit.areaServed);
    row.getCell(COL.manufacturer).value = asBlank(match?.brand || 'H&H Trecho');
    row.getCell(COL.modelNumber).value = asBlank(match?.model_number || 'NO MATCH');
    row.getCell(COL.nominalTons).value = asBlank(match?.tonnage_value || unit.tonnage);
    row.getCell(COL.unitType).value = asBlank(match?.unit_type || unit.family);
    row.getCell(COL.unitEer).value = asBlank(match?.unit_eer);
    row.getCell(COL.seerIeerr).value = asBlank(match?.seer_ieer);
    row.getCell(COL.supplyCfm).value = asBlank(match?.supply_airflow_cfm);
    row.getCell(COL.supplyEsp).value = asBlank(match?.supply_fan_esp_in_wg);
    row.getCell(COL.supplyQty).value = asBlank(match?.supply_fan_qty || 1);
    row.getCell(COL.supplyBhp).value = asBlank(match?.supply_fan_bhp);
    row.getCell(COL.supplyHp).value = asBlank(match?.supply_fan_hp);
    row.getCell(COL.supplyRpm).value = asBlank(match?.supply_fan_rpm);
    row.getCell(COL.coolingEat).value = joinSlash(match?.cooling_eat_db, match?.cooling_eat_wb);
    row.getCell(COL.coolingLat).value = joinSlash(match?.cooling_lat_db, match?.cooling_lat_wb);
    row.getCell(COL.coolingSensible).value = asBlank(match?.cooling_sensible_capacity_mbh);
    row.getCell(COL.coolingTotal).value = asBlank(match?.cooling_total_capacity_mbh);
    row.getCell(COL.heatingCfm).value = asBlank(match?.supply_airflow_cfm);
    row.getCell(COL.heatingEat).value = asBlank(match?.heating_eat_f);
    row.getCell(COL.heatingLat).value = asBlank(match?.heating_lat_f);
    row.getCell(COL.heatingInput).value = asBlank(match?.heating_input_capacity_mbh || unit.heatCapacity);
    row.getCell(COL.heatingOutput).value = asBlank(match?.heating_output_capacity_mbh);
    row.getCell(COL.voltPh).value = asBlank(match?.voltage_label || unit.voltage);
    row.getCell(COL.mca).value = asBlank(match?.mca);
    row.getCell(COL.mocp).value = asBlank(match?.mocp);
    row.getCell(COL.weight).value = asBlank(match?.operating_weight_lbs);
    row.getCell(COL.remarks).value = optionSummary(unit, match);
    row.commit();
  }

  return workbook.xlsx.writeBuffer();
}

async function parseDsSchedule(arrayBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  const worksheet = workbook.getWorksheet('Schedule') || workbook.worksheets[0];
  const rows = [];

  for (let rowNumber = 8; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const descriptor = normalizeText(row.getCell(1).value);
    const modelNumber = normalizeText(row.getCell(2).value);
    if (!descriptor && !modelNumber) continue;

    const parsed = parseDescriptor(descriptor);
    rows.push({
      sourceRowNumber: rowNumber,
      descriptor,
      modelNumber,
      brand: normalizeText(row.getCell(3).value),
      qty: normalizeText(row.getCell(4).value),
      airflowCfm: normalizeText(row.getCell(5).value),
      supplyFanHp: normalizeText(row.getCell(7).value),
      supplyFanEsp: normalizeText(row.getCell(8).value),
      supplyFanRpm: normalizeText(row.getCell(10).value),
      coolingTotalMbh: normalizeText(row.getCell(17).value),
      coolingSensibleMbh: normalizeText(row.getCell(18).value),
      unitEer: normalizeText(row.getCell(19).value),
      seerIeerr: normalizeText(row.getCell(20).value),
      refrigerant: normalizeText(row.getCell(23).value),
      heatingInputMbh: normalizeText(row.getCell(28).value),
      heatingOutputMbh: normalizeText(row.getCell(27).value),
      voltage: normalizeText(row.getCell(30).value),
      mca: normalizeText(row.getCell(31).value),
      mocp: normalizeText(row.getCell(33).value),
      weightLbs: normalizeText(row.getCell(35).value),
      remarks: normalizeText(row.getCell(36).value),
      ...parsed
    });
  }

  return { worksheetName: worksheet.name, rows };
}

async function createImportBatch(env, meta) {
  const stmt = env.DB.prepare(`
    INSERT INTO import_batches (source_filename, source_sheet, vendor, product_line, notes)
    VALUES (?, ?, ?, ?, ?)
    RETURNING id
  `).bind(meta.sourceFilename, meta.sourceSheet, meta.vendor, meta.productLine, meta.notes || '');
  const result = await stmt.first();
  return result?.id;
}

async function insertStagingRow(env, batchId, row) {
  const stmt = env.DB.prepare(`
    INSERT INTO staging_schedule_rows (
      batch_id, source_row_number, source_descriptor,
      raw_model_number, raw_brand, raw_qty, raw_airflow_cfm,
      raw_supply_fan_hp, raw_supply_fan_esp_in_wg, raw_supply_fan_rpm,
      raw_cooling_total_mbh, raw_cooling_sensible_mbh, raw_unit_eer, raw_seer_ieer,
      raw_refrigerant, raw_heating_input_mbh, raw_heating_output_mbh,
      raw_voltage, raw_mca, raw_mocp, raw_weight_lbs, raw_remarks,
      family_key, family_label, efficiency_key, efficiency_label,
      tonnage_key, tonnage_value, voltage_key, voltage_label,
      aux_heat_type_key, aux_heat_type_label, aux_heat_capacity_key, aux_heat_capacity_label,
      parse_status, parse_notes
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).bind(
    batchId,
    row.sourceRowNumber,
    row.descriptor,
    row.modelNumber,
    row.brand,
    row.qty,
    row.airflowCfm,
    row.supplyFanHp,
    row.supplyFanEsp,
    row.supplyFanRpm,
    row.coolingTotalMbh,
    row.coolingSensibleMbh,
    row.unitEer,
    row.seerIeerr,
    row.refrigerant,
    row.heatingInputMbh,
    row.heatingOutputMbh,
    row.voltage,
    row.mca,
    row.mocp,
    row.weightLbs,
    row.remarks,
    row.familyKey,
    row.familyLabel,
    row.efficiencyKey,
    row.efficiencyLabel,
    row.tonnageKey,
    row.tonnageValue,
    row.voltageKey,
    row.voltageLabel,
    row.auxHeatTypeKey,
    row.auxHeatTypeLabel,
    row.auxHeatCapacityKey,
    row.auxHeatCapacityLabel,
    row.modelNumber ? 'parsed' : 'review',
    row.modelNumber ? '' : 'Missing model number'
  );
  await stmt.run();
}

async function upsertUnitModelV2(env, batchId, sheetName, row) {
  const stmt = env.DB.prepare(`
    INSERT INTO unit_models_v2 (
      family_key, family_label, efficiency_key, efficiency_label,
      tonnage_key, tonnage_value, voltage_key, voltage_label,
      aux_heat_type_key, aux_heat_type_label, aux_heat_capacity_key, aux_heat_capacity_label,
      brand, model_number, unit_type, unit_eer, seer_ieer,
      supply_airflow_cfm, supply_fan_qty, supply_fan_hp, supply_fan_rpm, supply_fan_esp_in_wg,
      cooling_sensible_capacity_mbh, cooling_total_capacity_mbh,
      heating_input_capacity_mbh, heating_output_capacity_mbh,
      refrigerant_type, mca, mocp, operating_weight_lbs, remarks_default,
      source_batch_id, source_sheet, source_row_number, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
    ON CONFLICT(family_key, efficiency_key, tonnage_key, voltage_key, aux_heat_type_key, aux_heat_capacity_key, model_number)
    DO UPDATE SET
      family_label = excluded.family_label,
      efficiency_label = excluded.efficiency_label,
      tonnage_value = excluded.tonnage_value,
      voltage_label = excluded.voltage_label,
      aux_heat_type_label = excluded.aux_heat_type_label,
      aux_heat_capacity_label = excluded.aux_heat_capacity_label,
      brand = excluded.brand,
      unit_type = excluded.unit_type,
      unit_eer = excluded.unit_eer,
      seer_ieer = excluded.seer_ieer,
      supply_airflow_cfm = excluded.supply_airflow_cfm,
      supply_fan_qty = excluded.supply_fan_qty,
      supply_fan_hp = excluded.supply_fan_hp,
      supply_fan_rpm = excluded.supply_fan_rpm,
      supply_fan_esp_in_wg = excluded.supply_fan_esp_in_wg,
      cooling_sensible_capacity_mbh = excluded.cooling_sensible_capacity_mbh,
      cooling_total_capacity_mbh = excluded.cooling_total_capacity_mbh,
      heating_input_capacity_mbh = excluded.heating_input_capacity_mbh,
      heating_output_capacity_mbh = excluded.heating_output_capacity_mbh,
      refrigerant_type = excluded.refrigerant_type,
      mca = excluded.mca,
      mocp = excluded.mocp,
      operating_weight_lbs = excluded.operating_weight_lbs,
      remarks_default = excluded.remarks_default,
      source_batch_id = excluded.source_batch_id,
      source_sheet = excluded.source_sheet,
      source_row_number = excluded.source_row_number,
      updated_at = CURRENT_TIMESTAMP
  `).bind(
    row.familyKey,
    row.familyLabel,
    row.efficiencyKey,
    row.efficiencyLabel,
    row.tonnageKey,
    row.tonnageValue,
    row.voltageKey,
    row.voltageLabel,
    row.auxHeatTypeKey,
    row.auxHeatTypeLabel,
    row.auxHeatCapacityKey,
    row.auxHeatCapacityLabel,
    row.brand,
    row.modelNumber,
    row.familyLabel,
    toNumber(row.unitEer),
    row.seerIeerr,
    toNumber(row.airflowCfm),
    toNumber(row.qty) || 1,
    toNumber(row.supplyFanHp),
    toNumber(row.supplyFanRpm),
    toNumber(row.supplyFanEsp),
    toNumber(row.coolingSensibleMbh),
    toNumber(row.coolingTotalMbh),
    toNumber(row.heatingInputMbh),
    toNumber(row.heatingOutputMbh),
    row.refrigerant,
    row.mca,
    row.mocp,
    toNumber(row.weightLbs),
    row.remarks,
    batchId,
    sheetName,
    row.sourceRowNumber
  );
  await stmt.run();
}

async function importDsSchedule(env, arrayBuffer, sourceFilename) {
  const parsed = await parseDsSchedule(arrayBuffer);
  const batchId = await createImportBatch(env, {
    sourceFilename,
    sourceSheet: parsed.worksheetName,
    vendor: 'Tempmaster',
    productLine: 'DS Commercial',
    notes: 'Imported from DS Commercial workbook'
  });

  let staged = 0;
  let upserted = 0;

  for (const row of parsed.rows) {
    await insertStagingRow(env, batchId, row);
    staged += 1;
    if (row.modelNumber && row.tonnageKey && row.voltageKey) {
      await upsertUnitModelV2(env, batchId, parsed.worksheetName, row);
      upserted += 1;
    }
  }

  return {
    batchId,
    worksheet: parsed.worksheetName,
    staged,
    upserted
  };
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === 'GET' && url.pathname === '/api/catalog') {
      const filters = {
        family: url.searchParams.get('family') || '',
        efficiency: url.searchParams.get('efficiency') || '',
        tonnage: url.searchParams.get('tonnage') || '',
        voltage: url.searchParams.get('voltage') || '',
        heatType: url.searchParams.get('heatType') || '',
        heatCapacity: url.searchParams.get('heatCapacity') || ''
      };
      const results = await listCatalog(env, filters);
      return json({ items: results });
    }

    if (request.method === 'POST' && url.pathname === '/api/import-ds-schedule') {
      try {
        const formData = await request.formData();
        const file = formData.get('file');
        if (!file || typeof file.arrayBuffer !== 'function') {
          return json({ error: 'Upload a file in form-data under key "file".' }, 400);
        }
        const result = await importDsSchedule(env, await file.arrayBuffer(), file.name || 'upload.xlsx');
        return json(result, 200);
      } catch (error) {
        return json({ error: error.message || 'Import failed' }, 500);
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
