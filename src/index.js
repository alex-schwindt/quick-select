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
  if (['208/3', '208-3', '2083', '208/230/3', '2082303'].includes(compact.replace('/230', ''))) return '208/230/3';
  if (['208/230/3', '2082303'].includes(compact)) return '208/230/3';
  if (['460/3', '460-3', '4603'].includes(compact)) return '460/3';
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

function normalizeEconomizer(value) {
  const v = normalizeText(value).toLowerCase();
  if (['powered', 'power exhaust', 'powered exhaust'].includes(v)) return 'powered';
  if (['barometric', 'barometric relief'].includes(v)) return 'barometric';
  return 'none';
}

function economizerLookupValue(unit) {
  const econ = normalizeEconomizer(unit.economizer);
  return econ === 'barometric' ? 'none' : econ;
}

function buildSelectionCode(unit) {
  const familyCode = normalizeFamily(unit.family) === 'Heat Pump' ? 'HP' : 'AC';
  const efficiencyCode = normalizeEfficiency(unit.efficiency) === 'High' ? 'HI' : 'STD';
  const tonnageCode = String(unit.tonnage).replace('.', 'P');
  const voltageCode = normalizeVoltage(unit.voltage) === '460/3' ? '460' : '208';
  const heatTypeKey = normalizeHeatType(unit.heatType);
  const heatCode = heatTypeKey === 'None'
    ? 'NOHEAT'
    : heatTypeKey === 'Electric Heat'
      ? `ELEC-${normalizeHeatCapacity(unit.heatCapacity).replace(/\s+/g, '').replace(/[^0-9A-Za-z-]/g, '')}`
      : `GAS-${normalizeHeatCapacity(unit.heatCapacity).replace(/\s+/g, '').replace(/[^0-9A-Za-z-]/g, '')}`;
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

async function listCatalog(env, filters) {
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
  if (filters.heatCapacity || normalizeHeatType(filters.heatType) !== 'None') {
    sql += ' AND m.heat_capacity = ?';
    binds.push(filters.heatCapacity || '');
  }
  if (filters.hotGasReheat !== undefined) {
    sql += ' AND m.hot_gas_reheat = ?';
    binds.push(filters.hotGasReheat ? 1 : 0);
  }
  if (filters.economizer) {
    sql += ' AND m.economizer = ?';
    binds.push(filters.economizer);
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
    heatCapacity: normalizeHeatCapacity(unit.heatCapacity || ''),
    hotGasReheat: Boolean(unit.hotGasReheat),
    economizer: economizerLookupValue(unit)
  });

  return matches[0] || null;
}

function joinSlash(a, b) {
  const parts = [asBlank(a), asBlank(b)].filter((v) => v !== '');
  return parts.length ? parts.join(' / ') : '';
}

function buildResolvedScheduleRow(unit, match, index = 0) {
  return {
    tag: normalizeText(unit.tag) || `RTU-${index + 1}`,
    areaServed: asBlank(unit.areaServed),
    manufacturer: asBlank(match?.brand || 'H&H Trecho'),
    modelNumber: asBlank(match?.model_number || buildSelectionCode(unit)),
    nominalTons: asBlank(match?.tonnage || unit.tonnage),
    unitType: asBlank(match?.unit_type || unit.family),
    unitEer: asBlank(match?.unit_eer),
    seerIeerr: asBlank(match?.seer_ieer),
    supplyCfm: asBlank(match?.supply_airflow_cfm || match?.cooling_cfm),
    supplyEsp: asBlank(match?.supply_fan_esp_in_wg),
    supplyQty: asBlank(match?.supply_fan_qty || match?.quantity || 1),
    supplyBhp: asBlank(match?.supply_fan_bhp),
    supplyHp: asBlank(match?.supply_fan_hp),
    supplyRpm: asBlank(match?.supply_fan_rpm),
    coolingEat: joinSlash(match?.cooling_eat_db, match?.cooling_eat_wb),
    coolingLat: joinSlash(match?.cooling_lat_db, match?.cooling_lat_wb),
    coolingSensible: asBlank(match?.cooling_sensible_capacity_mbh),
    coolingTotal: asBlank(match?.cooling_total_capacity_mbh),
    heatingCfm: asBlank(match?.heating_cfm || match?.supply_airflow_cfm || match?.cooling_cfm),
    heatingEat: asBlank(match?.heating_eat_f),
    heatingLat: asBlank(match?.heating_lat_f),
    heatingInput: asBlank(match?.heating_capacity_mbtu || unit.heatCapacity),
    heatingOutput: asBlank(match?.heating_output_capacity || match?.heating_afue),
    voltPh: asBlank(match?.voltage || unit.voltage),
    mca: asBlank(match?.mca),
    mocp: asBlank(match?.mocp),
    weight: asBlank(match?.operating_weight_lbs),
    remarks: optionSummary(unit, match),
    selectionCode: buildSelectionCode(unit),
    matchFound: Boolean(match),
    cutsheetUrl: asBlank(match?.cutsheet_url),
    accessoriesUrl: asBlank(match?.accessories_url),
    wiringUrl: asBlank(match?.wiring_url),
    iomUrl: asBlank(match?.iom_url)
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

    if (rowNumber > templateRowNumber) {
      worksheet.duplicateRow(templateRowNumber, 1, true);
    }

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
        family: normalizeFamily(url.searchParams.get('family') || ''),
        efficiency: normalizeEfficiency(url.searchParams.get('efficiency') || ''),
        tonnage: url.searchParams.get('tonnage') || '',
        voltage: normalizeVoltage(url.searchParams.get('voltage') || ''),
        heatType: normalizeHeatType(url.searchParams.get('heatType') || ''),
        heatCapacity: normalizeHeatCapacity(url.searchParams.get('heatCapacity') || ''),
        hotGasReheat: (url.searchParams.get('hotGasReheat') || '') === 'true',
        economizer: economizerLookupValue({ economizer: url.searchParams.get('economizer') || 'none' })
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
