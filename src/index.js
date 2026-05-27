import ExcelJS from 'exceljs';

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
    SELECT m.*, d.cutsheet_url, d.accessories_url
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
    row.getCell(3).value = asBlank(unit.areaServed);
    row.getCell(4).value = asBlank(match?.brand || 'H&H Trecho');
    row.getCell(5).value = asBlank(match?.model_number || buildSelectionCode(unit));
    row.getCell(6).value = asBlank(match?.tonnage || unit.tonnage);
    row.getCell(7).value = asBlank(match?.unit_type || unit.family);
    row.getCell(8).value = asBlank(match?.unit_eer);
    row.getCell(9).value = asBlank(match?.seer_ieer);
    row.getCell(10).value = asBlank(match?.supply_airflow_cfm || match?.cooling_cfm);
    row.getCell(11).value = asBlank(match?.outside_air_min_cfm);
    row.getCell(12).value = asBlank(match?.outside_air_max_cfm);
    row.getCell(13).value = asBlank(match?.supply_fan_esp_in_wg);
    row.getCell(14).value = asBlank(match?.supply_fan_tsp_in_wg);
    row.getCell(15).value = asBlank(match?.supply_fan_qty);
    row.getCell(16).value = asBlank(match?.supply_fan_bhp);
    row.getCell(17).value = asBlank(match?.supply_fan_hp);
    row.getCell(18).value = asBlank(match?.supply_fan_rpm);
    row.getCell(19).value = [asBlank(match?.cooling_eat_db), asBlank(match?.cooling_eat_wb)].filter(v => v !== '').join(' / ');
    row.getCell(20).value = [asBlank(match?.cooling_lat_db), asBlank(match?.cooling_lat_wb)].filter(v => v !== '').join(' / ');
    row.getCell(21).value = asBlank(match?.cooling_sensible_capacity_mbh);
    row.getCell(22).value = asBlank(match?.cooling_total_capacity_mbh);
    row.getCell(23).value = asBlank(match?.heating_cfm);
    row.getCell(24).value = asBlank(match?.heating_eat_f);
    row.getCell(25).value = asBlank(match?.heating_lat_f);
    row.getCell(26).value = asBlank(match?.heating_capacity_mbtu || unit.heatCapacity);
    row.getCell(27).value = asBlank(match?.heating_gas_cfh);
    row.getCell(28).value = asBlank(match?.heating_afue);
    row.getCell(29).value = asBlank(match?.condenser_qty_fans);
    row.getCell(30).value = asBlank(match?.condenser_hp_each);
    row.getCell(31).value = asBlank(match?.condenser_type);
    row.getCell(32).value = asBlank(match?.refrigerant_charge);
    row.getCell(33).value = asBlank(match?.voltage_display || unit.voltage);
    row.getCell(34).value = asBlank(match?.mca);
    row.getCell(35).value = asBlank(match?.mocp);
    row.getCell(36).value = asBlank(match?.filter_type);
    row.getCell(37).value = asBlank(match?.operating_weight_lbs);
    row.getCell(38).value = optionSummary(unit, match);
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

