import ExcelJS from 'exceljs';

function json(data, status = 200) {
  return new Response(JSON.stringify(data), { status, headers: { 'Content-Type': 'application/json' } });
}

function normalizeText(value) {
  return String(value ?? '').trim();
}

function buildSelectionCode(unit) {
  const familyCode = unit.family === 'Heat Pump' ? 'HP' : 'AC';
  const efficiencyCode = unit.efficiency === 'High' ? 'HI' : 'STD';
  const tonnageCode = String(unit.tonnage).replace('.', 'P');
  const voltageCode = unit.voltage === '460/3' ? '460' : '208';
  const heatCode = unit.heatType === 'None'
    ? 'NOHEAT'
    : unit.heatType === 'Electric Heat'
      ? `ELEC-${String(unit.heatCapacity || '').replace(/\s+/g, '').replace(/[^0-9A-Za-z-]/g, '')}`
      : `GAS-${String(unit.heatCapacity || '').replace(/\s+/g, '').replace(/[^0-9A-Za-z-]/g, '')}`;
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
  let sql = `
    SELECT m.id, m.family, m.efficiency, m.tonnage, m.voltage, m.heat_type, m.heat_capacity,
           m.model_code, m.model_number, m.unit_type, m.unit_eer, m.seer_ieer,
           m.cooling_cfm, m.cooling_total_capacity_mbh, m.cooling_sensible_capacity_mbh,
           m.heating_capacity_mbtu, m.refrigerant_type, m.refrigerant_charge,
           m.mca, m.mocp, m.filter_type, m.operating_weight_lbs,
           d.cutsheet_url, d.accessories_url
    FROM unit_models m
    LEFT JOIN unit_documents d ON d.model_id = m.id
    WHERE 1 = 1
  `;
  const binds = [];
  if (filters.family) { sql += ' AND m.family = ?'; binds.push(filters.family); }
  if (filters.efficiency) { sql += ' AND m.efficiency = ?'; binds.push(filters.efficiency); }
  if (filters.tonnage) { sql += ' AND m.tonnage = ?'; binds.push(Number(filters.tonnage)); }
  if (filters.voltage) { sql += ' AND m.voltage = ?'; binds.push(filters.voltage); }
  if (filters.heatType) { sql += ' AND m.heat_type = ?'; binds.push(filters.heatType); }
  if (filters.heatCapacity) { sql += ' AND m.heat_capacity = ?'; binds.push(filters.heatCapacity); }
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
    row.getCell(23).value = unit.heatType === 'Electric Heat' ? (match?.cooling_cfm || '--') : '--';
    row.getCell(26).value = unit.heatType === 'Aluminum Gas Heat' ? (match?.heating_capacity_mbtu || unit.heatCapacity || '--') : '--';
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
