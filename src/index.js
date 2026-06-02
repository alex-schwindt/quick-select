import ExcelJS from 'exceljs';

function toInt(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? Math.trunc(n) : fallback;
}

function cleanBlank(value) {
  if (value === null || value === undefined) return null;
  const text = String(value).trim();
  return text === '' ? null : text;
}

function coerceArray(value) {
  return Array.isArray(value) ? value : [];
}

function summarizeIssueRows(rows) {
  return rows.map((row) => ({
    source_row_number: row.source_row_number,
    model_number: row.raw_model_number,
    parse_status: row.parse_status,
    parse_notes: cleanBlank(row.parse_notes),
  }));
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
      SUM(CASE WHEN action = 'inserted' THEN 1 ELSE 0 END) AS catalog_inserts,
      SUM(CASE WHEN action = 'updated' THEN 1 ELSE 0 END) AS catalog_updates,
      SUM(CASE WHEN action = 'unchanged' THEN 1 ELSE 0 END) AS catalog_unchanged,
      SUM(CASE WHEN action = 'duplicate_in_batch' THEN 1 ELSE 0 END) AS duplicates_in_batch
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
    issues.push({
      type: 'parse_warnings',
      rows: warningRows,
    });
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
      imr.created_at
    FROM import_model_results imr
    JOIN staging_schedule_rows ssr ON ssr.id = imr.staging_row_id
    LEFT JOIN unit_models_v2 um ON um.id = imr.unit_model_id
    WHERE imr.batch_id = ?
    ORDER BY ssr.source_row_number, imr.id
  `).bind(batchId).all();
  return json({ ok: true, batch, results: coerceArray(rows.results) });
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), { status, headers: { 'Content-Type': 'application/json' } });
}

function normalizeText(value) {
  return String(value ?? '').trim();
}

function asBlank(value) {
  return value === null || value === undefined || value === '' ? '' : value;
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
  return v.replace(/\s+/g, '').replace('mbh', '').replace('.0', '');
}

function normalizeTonnage(value) {
  return String(value ?? '').trim().replace('.0', '');
}

function buildSelectionCode(unit) {
  const familyCode = normalizeFamily(unit.family) === 'hp' ? 'HP' : 'AC';
  const efficiencyCode = normalizeEfficiency(unit.efficiency) === 'high' ? 'HI' : 'STD';
  const tonnageCode = String(unit.tonnage).replace('.', 'P');
  const voltageCode = normalizeVoltage(unit.voltage) === '460-3' ? '460' : '208';
  const heatTypeKey = normalizeHeatType(unit.heatType);
  const heatCode = heatTypeKey === 'none'
    ? 'NOHEAT'
    : heatTypeKey === 'electric'
      ? `ELEC-${normalizeHeatCapacity(unit.heatCapacity)}`
      : `GAS-${normalizeHeatCapacity(unit.heatCapacity)}MBH`;
  const reheatCode = unit.hotGasReheat ? 'HGRH' : 'NOHGRH';
  const econCode = unit.economizer === 'barometric' ? 'ECO-BARO' : unit.economizer === 'powered' ? 'ECO-PE' : 'NOECO';
  return `${familyCode}-${efficiencyCode}-${tonnageCode}-${voltageCode}-${heatCode}-${reheatCode}-${econCode}`;
}

function optionSummary(unit, match) {
  return normalizeText(unit.remarks) || normalizeText(match?.remarks_default) || [
    unit.hotGasReheat ? 'Hot Gas Reheat' : null,
    unit.economizer === 'barometric' ? 'Economizer w/ Barometric Relief' : null,
    unit.economizer === 'powered' ? 'Economizer w/ Powered Exhaust' : null,
    unit.curb ? 'Curb' : null
  ].filter(Boolean).join(', ');
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

async function listCatalog(env, filters = {}) {
  let sql = `
    SELECT m.*, d.cut_sheet_url, d.accessories_url, d.wiring_url, d.iom_url
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
  if (filters.heatCapacity && normalizeHeatType(filters.heatType) !== 'none') {
    sql += ' AND m.heat_capacity = ?';
    binds.push(filters.heatCapacity);
  }

  sql += ' ORDER BY m.model_number';
  const stmt = env.DB.prepare(sql).bind(...binds);
  const result = await stmt.all();
  return result.results || [];
}

async function findMatchingModel(env, unit) {
  const matches = await listCatalog(env, {
    family: normalizeFamily(unit.family),
    efficiency: normalizeEfficiency(unit.efficiency),
    tonnage: unit.tonnage,
    voltage: normalizeVoltage(unit.voltage),
    heatType: normalizeHeatType(unit.heatType),
    heatCapacity: normalizeHeatCapacity(unit.heatCapacity),
  });
  return matches[0] || null;
}

function joinSlash(a, b) {
  const parts = [asBlank(a), asBlank(b)].filter((v) => v !== '');
  return parts.length ? parts.join('/') : '';
}

function buildResolvedScheduleRow(unit, match, index = 0) {
  return {
    tag: normalizeText(unit.tag) || `RTU-${index + 1}`,
    areaServed: asBlank(unit.areaServed),
    manufacturer: asBlank(match?.brand || 'Tempmaster'),
    modelNumber: asBlank(match?.model_number || buildSelectionCode(unit)),
    nominalTons: asBlank(match?.tonnage || unit.tonnage),
    unitType: asBlank(match?.unit_type || unit.family),
    unitEer: asBlank(match?.unit_eer),
    seerIeerr: asBlank(match?.seer_ieer),
    supplyCfm: asBlank(match?.cooling_cfm),
    supplyEsp: '',
    supplyQty: 1,
    supplyBhp: '',
    supplyHp: '',
    supplyRpm: '',
    coolingEat: joinSlash(match?.cooling_eat_db, match?.cooling_eat_wb),
    coolingLat: joinSlash(match?.cooling_lat_db, match?.cooling_lat_wb),
    coolingSensible: asBlank(match?.cooling_sensible_capacity_mbh),
    coolingTotal: asBlank(match?.cooling_total_capacity_mbh),
    heatingCfm: asBlank(match?.heating_cfm || match?.cooling_cfm),
    heatingEat: asBlank(match?.heating_eat_f),
    heatingLat: asBlank(match?.heating_lat_f),
    heatingInput: asBlank(match?.heating_capacity_mbtu || unit.heatCapacity),
    heatingOutput: asBlank(match?.heating_output_capacity || ''),
    voltPh: asBlank(match?.voltage || unit.voltage),
    mca: asBlank(match?.mca),
    mocp: asBlank(match?.mocp),
    weight: asBlank(match?.operating_weight_lbs),
    remarks: optionSummary(unit, match),
    selectionCode: buildSelectionCode(unit),
    matchFound: Boolean(match),
    cutSheetUrl: asBlank(match?.cut_sheet_url),
    accessoriesUrl: asBlank(match?.accessories_url),
    wiringUrl: asBlank(match?.wiring_url),
    iomUrl: asBlank(match?.iom_url),
  };
}

async function resolveScheduleRows(env, units) {
  const rows = [];
  for (let index = 0; index < units.length; index += 1) {
    const unit = units[index];
    const match = await findMatchingModel(env, unit);
    rows.push(buildResolvedScheduleRow(unit, match, index));
  }
  return rows;
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
    row.getCell(COL.heatingInput).value = scheduleRow.heatingInput;
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

    if (request.method === 'POST' && url.pathname === '/api/preview-schedule') {
      try {
        const payload = await request.json();
        const units = Array.isArray(payload?.units) ? payload.units : [];
        const rows = await resolveScheduleRows(env, units);
        return json({ rows });
      } catch (error) {
        return new Response(error.message || 'Preview failed', { status: 500 });
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