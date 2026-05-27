import ExcelJS from 'exceljs';

function json(data, status = 200) {
  return new Response(JSON.stringify(data), { status, headers: { 'Content-Type': 'application/json' } });
}

function normalizeText(value) {
  return String(value ?? '').trim();
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

function optionSummary(unit) {
  return [
    unit.hotGasReheat ? 'Hot Gas Reheat' : null,
    unit.economizer === 'barometric' ? 'Economizer w/ Barometric Relief' : null,
    unit.economizer === 'powered' ? 'Economizer w/ Powered Exhaust' : null
  ].filter(Boolean).join(', ');
}

async function getTemplateWorkbook(env) {
  const object = await env.TEMPLATES.get('SSR-Schedule-Example.xlsx');
  if (!object) throw new Error('Template workbook not found in R2 bucket.');
  const arrayBuffer = await object.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  return workbook;
}

async function listCatalog(env, filters) {
  const familyKey = normalizeFamily(filters.family);
  const efficiencyKey = normalizeEfficiency(filters.efficiency);
  const tonnageKey = normalizeTonnage(filters.tonnage);
  const voltageKey = normalizeVoltage(filters.voltage);
  const heatTypeKey = normalizeHeatType(filters.heatType);
  const heatCapacityKey = normalizeHeatCapacity(filters.heatCapacity);

  let sql = `
    SELECT m.id, m.family, m.efficiency, m.tonnage, m.voltage, m.heat_type, m.heat_capacity,
           m.model_code, m.model_number, m.unit_type, m.unit_eer, m.seer_ieer,
           m.cooling_cfm, m.cooling_total_capacity_mbh, m.cooling_sensible_capacity_mbh,
           m.heating_capacity_mbtu, m.refrigerant_type, m.refrigerant_charge,
           m.mca, m.mocp, m.filter_type, m.operating_weight_lbs,
           m.family_key, m.efficiency_key, m.tonnage_key, m.voltage_key, m.heat_type_key, m.heat_capacity_key,
           d.cutsheet_url, d.accessories_url
    FROM unit_models m
    LEFT JOIN unit_documents d ON d.model_id = m.id
    WHERE 1 = 1
  `;
  const binds = [];
  if (filters.family) { sql += ' AND m.family_key = ?'; binds.push(familyKey); }
  if (filters.efficiency) { sql += ' AND m.efficiency_key = ?'; binds.push(efficiencyKey); }
  if (filters.tonnage) { sql += ' AND m.tonnage_key = ?'; binds.push(tonnageKey); }
  if (filters.voltage) { sql += ' AND m.voltage_key = ?'; binds.push(voltageKey); }
  if (filters.heatType) { sql += ' AND m.heat_type_key = ?'; binds.push(heatTypeKey); }
  if (filters.heatCapacity || heatTypeKey !== 'none') { sql += ' AND m.heat_capacity_key = ?'; binds.push(heatCapacityKey); }
  sql += ' ORDER BY m.family, m.tonnage, m.model_number';
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
    row.getCell(2).value = normalizeText(unit.tag);
    row.getCell(3).value = normalizeText(unit.areaServed);
    row.getCell(4).value = 'H&H Trecho';
    row.getCell(5).value = match?.model_number || buildSelectionCode(unit);
    row.getCell(6).value = Number(unit.tonnage) || '';
    row.getCell(7).value = match?.unit_type || unit.family;
    row.getCell(8).value = match?.unit_eer || '';
    row.getCell(9).value = match?.seer_ieer || unit.efficiency;
    row.getCell(10).value = match?.cooling_cfm || '';
    row.getCell(21).value = match?.cooling_sensible_capacity_mbh || '';
    row.getCell(22).value = match?.cooling_total_capacity_mbh || '';
    row.getCell(26).value = normalizeHeatType(unit.heatType) === 'gas' ? (match?.heating_capacity_mbtu || unit.heatCapacity || '--') : '--';
    row.getCell(37).value = match?.refrigerant_type || '';
    row.getCell(38).value = match?.refrigerant_charge || '';
    row.getCell(39).value = unit.voltage || '';
    row.getCell(40).value = match?.mca || '';
    row.getCell(41).value = match?.mocp || '';
    row.getCell(42).value = match?.filter_type || '';
    row.getCell(43).value = match?.operating_weight_lbs || '';
    row.getCell(44).value = normalizeText(unit.remarks) || optionSummary(unit) || '';
    row.commit();
  }

  return workbook.xlsx.writeBuffer();
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
