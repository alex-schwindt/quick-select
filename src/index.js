import * as XLSX from 'xlsx';

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json' },
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
  const s = normalizeText(value).toLowerCase().replace(/\s+/g, '').replace(/-/g, '/');
  if (/^208(\/230)?(\/3|\/1)?/.test(s) || s === '208' || s.startsWith('208/')) return '208/230/3';
  if (s === '2083' || s === '20830' || s === '208/3') return '208/230/3';
  if (s === '460' || s === '4603' || s === '460/3' || s.startsWith('460/')) return '460/3';
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
  const n = Number(value);
  return Number.isFinite(n) ? n : value;
}

function buildSelectionCode(unit) {
  const familyCode = normalizeFamily(unit.family) === 'Heat Pump' ? 'HP' : 'AC';
  const efficiencyCode = normalizeEfficiency(unit.efficiency) === 'High' ? 'HI' : 'STD';
  const tonnageCode = String(unit.tonnage).replace('.', 'P');
  const voltageCode = normalizeVoltage(unit.voltage) === '460/3' ? '460' : '208';
  const heatTypeKey = normalizeHeatType(unit.heatType);
  const normalizedHeatCapacity = normalizeHeatCapacity(unit.heatCapacity).replace(/\s+/g, '').replace(/[^0-9A-Za-z-]/g, '');
  const heatCode =
    heatTypeKey === 'None'
      ? 'NOHEAT'
      : heatTypeKey === 'Electric Heat'
      ? `ELEC-${normalizedHeatCapacity}`
      : `GAS-${normalizedHeatCapacity}`;
  const reheatCode = unit.hotGasReheat ? 'HGRH' : 'NOHGRH';
  const econCode =
    unit.economizer === 'barometric'
      ? 'ECO-BARO'
      : unit.economizer === 'powered'
      ? 'ECO-PE'
      : 'NOECO';
  return `${familyCode}-${efficiencyCode}-${tonnageCode}-${voltageCode}-${heatCode}-${reheatCode}-${econCode}`;
}

function optionSummary(unit, match) {
  return (
    normalizeText(unit.remarks) ||
    normalizeText(match?.remarks) ||
    [
      unit.hotGasReheat ? 'Hot Gas Reheat' : null,
      unit.economizer === 'barometric' ? 'Economizer w/ Barometric Relief' : null,
      unit.economizer === 'powered' ? 'Economizer w/ Powered Exhaust' : null,
      unit.curb ? 'Curb' : null,
    ]
      .filter(Boolean)
      .join(', ')
  );
}

async function listCatalog(env, filters = {}) {
  let sql = `SELECT m.*, d.cutsheet_url, d.accessories_url, d.wiring_url, d.iom_url
    FROM unit_models m
    LEFT JOIN unit_documents d ON d.model_id = m.id
    WHERE 1=1`;
  const binds = [];

  if (filters.family) {
    sql += ' AND m.family = ?';
    binds.push(filters.family);
  }
  if (filters.efficiency) {
    sql += ' AND m.efficiency = ?';
    binds.push(filters.efficiency);
  }
  if (filters.tonnage !== null && filters.tonnage !== undefined && filters.tonnage !== '') {
    sql += ' AND m.tonnage = ?';
    binds.push(Number(filters.tonnage));
  }
  if (filters.voltage) {
    sql += ' AND m.voltage = ?';
    binds.push(filters.voltage);
  }
  if (filters.heatType) {
    sql += ' AND m.heat_type = ?';
    binds.push(filters.heatType);
  }
  if (normalizeHeatType(filters.heatType) === 'None') {
    sql += " AND COALESCE(m.heat_capacity, '') = ''";
  } else if (filters.heatCapacity) {
    sql += ' AND m.heat_capacity = ?';
    binds.push(filters.heatCapacity);
  }

  sql += ' ORDER BY m.family, m.tonnage, m.voltage, m.heat_type, m.heat_capacity, m.model_number';
  const result = await env.DB.prepare(sql).bind(...binds).all();
  return result.results || [];
}

async function findCatalogMatch(env, unit) {
  const family = normalizeFamily(unit.family);
  const efficiency = normalizeEfficiency(unit.efficiency);
  const tonnage = normalizeTonnage(unit.tonnage);
  const voltage = normalizeVoltage(unit.voltage);
  const heatType = normalizeHeatType(unit.heatType);
  const heatCapacity = normalizeHeatCapacity(unit.heatCapacity);

  let sql = `SELECT m.*, d.cutsheet_url, d.accessories_url, d.wiring_url, d.iom_url
    FROM unit_models m
    LEFT JOIN unit_documents d ON d.model_id = m.id
    WHERE m.family = ?
      AND m.efficiency = ?
      AND m.tonnage = ?
      AND m.voltage = ?
      AND m.heat_type = ?`;
  const binds = [family, efficiency, Number(tonnage), voltage, heatType];

  if (heatType === 'None') {
    sql += " AND COALESCE(m.heat_capacity, '') = ''";
  } else {
    sql += ' AND m.heat_capacity = ?';
    binds.push(heatCapacity);
  }

  sql += ' LIMIT 1';
  return (await env.DB.prepare(sql).bind(...binds).first()) || null;
}

function buildResolvedScheduleRow(unit, match, index = 0) {
  const modelNumber = asBlank(match?.model_number) || buildSelectionCode(unit);
  const isMatched = Boolean(match);

  return {
    tag: normalizeText(unit.tag) || `RTU-${index + 1}`,
    areaServed: asBlank(unit.areaServed),
    manufacturer: isMatched ? 'Tempmaster' : 'H&H Trecho',
    modelNumber,
    nominalTons: asBlank(match?.tonnage) || unit.tonnage,
    unitType: asBlank(match?.unit_type) || normalizeFamily(unit.family),
    unitEer: asBlank(match?.unit_eer),
    seerIeerr: asBlank(match?.seer_ieer),
    supplyCfm: asBlank(match?.cooling_cfm),
    supplyEsp: '',
    supplyQty: 1,
    supplyBhp: '',
    supplyHp: '',
    supplyRpm: '',
    coolingEat: '',
    coolingLat: '',
    coolingSensible: asBlank(match?.cooling_sensible_capacity_mbh),
    coolingTotal: asBlank(match?.cooling_total_capacity_mbh),
    heatingCfm: asBlank(match?.cooling_cfm),
    heatingEat: '',
    heatingLat: '',
    heatingInput: asBlank(match?.heating_capacity_mbtu) || (normalizeHeatType(unit.heatType) === 'None' ? '' : asBlank(unit.heatCapacity)),
    heatingOutput: '',
    voltPh: asBlank(match?.voltage) || normalizeVoltage(unit.voltage),
    mca: asBlank(match?.mca),
    mocp: asBlank(match?.mocp),
    weight: asBlank(match?.operating_weight_lbs),
    remarks: optionSummary(unit, match),
    selectionCode: buildSelectionCode(unit),
    matchFound: isMatched,
    cutsheetUrl: asBlank(match?.cutsheet_url),
    accessoriesUrl: asBlank(match?.accessories_url),
    wiringUrl: asBlank(match?.wiring_url),
    iomUrl: asBlank(match?.iom_url),
  };
}

async function resolveScheduleRows(env, units) {
  const rows = [];
  for (let i = 0; i < units.length; i++) {
    const unit = units[i];
    const match = await findCatalogMatch(env, unit);
    rows.push(buildResolvedScheduleRow(unit, match, i));
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

function setCellValue(sheet, col, row, value) {
  const addr = XLSX.utils.encode_cell({ c: col - 1, r: row - 1 });
  if (!sheet[addr]) sheet[addr] = {};
  const cell = sheet[addr];
  if (value === '' || value == null) {
    cell.t = 's';
    cell.v = '';
    cell.w = '';
  } else if (typeof value === 'number') {
    cell.t = 'n';
    cell.v = value;
    cell.w = String(value);
  } else {
    cell.t = 's';
    cell.v = String(value);
    cell.w = String(value);
  }
}

function expandSheetRange(sheet, maxRow) {
  if (!sheet['!ref']) return;
  const range = XLSX.utils.decode_range(sheet['!ref']);
  if (maxRow - 1 > range.e.r) {
    range.e.r = maxRow - 1;
    sheet['!ref'] = XLSX.utils.encode_range(range);
  }
}

async function createWorkbook(env, units) {
  const obj = await env.TEMPLATES.get('SSR-Schedule-Example.xlsx');
  if (!obj) throw new Error('Template workbook not found in R2 bucket.');
  const buf = await obj.arrayBuffer();
  const workbook = XLSX.read(buf, { type: 'array', cellStyles: true, cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const templateRow = 4;
  const rows = await resolveScheduleRows(env, units);

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const rowNum = templateRow + i;
    setCellValue(sheet, COL.tag, rowNum, r.tag);
    setCellValue(sheet, COL.areaServed, rowNum, r.areaServed);
    setCellValue(sheet, COL.manufacturer, rowNum, r.manufacturer);
    setCellValue(sheet, COL.modelNumber, rowNum, r.modelNumber);
    setCellValue(sheet, COL.nominalTons, rowNum, r.nominalTons);
    setCellValue(sheet, COL.unitType, rowNum, r.unitType);
    setCellValue(sheet, COL.unitEer, rowNum, r.unitEer);
    setCellValue(sheet, COL.seerIeerr, rowNum, r.seerIeerr);
    setCellValue(sheet, COL.supplyCfm, rowNum, r.supplyCfm);
    setCellValue(sheet, COL.supplyEsp, rowNum, r.supplyEsp);
    setCellValue(sheet, COL.supplyQty, rowNum, r.supplyQty);
    setCellValue(sheet, COL.supplyBhp, rowNum, r.supplyBhp);
    setCellValue(sheet, COL.supplyHp, rowNum, r.supplyHp);
    setCellValue(sheet, COL.supplyRpm, rowNum, r.supplyRpm);
    setCellValue(sheet, COL.coolingEat, rowNum, r.coolingEat);
    setCellValue(sheet, COL.coolingLat, rowNum, r.coolingLat);
    setCellValue(sheet, COL.coolingSensible, rowNum, r.coolingSensible);
    setCellValue(sheet, COL.coolingTotal, rowNum, r.coolingTotal);
    setCellValue(sheet, COL.heatingCfm, rowNum, r.heatingCfm);
    setCellValue(sheet, COL.heatingEat, rowNum, r.heatingEat);
    setCellValue(sheet, COL.heatingLat, rowNum, r.heatingLat);
    setCellValue(sheet, COL.heatingInput, rowNum, r.heatingInput);
    setCellValue(sheet, COL.heatingOutput, rowNum, r.heatingOutput);
    setCellValue(sheet, COL.voltPh, rowNum, r.voltPh);
    setCellValue(sheet, COL.mca, rowNum, r.mca);
    setCellValue(sheet, COL.mocp, rowNum, r.mocp);
    setCellValue(sheet, COL.weight, rowNum, r.weight);
    setCellValue(sheet, COL.remarks, rowNum, r.remarks);
  }

  expandSheetRange(sheet, templateRow + rows.length - 1);
  return XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
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
      return json({ items: await listCatalog(env, filters) });
    }

    if (request.method === 'POST' && url.pathname === '/api/preview-schedule') {
      try {
        const payload = await request.json();
        const units = Array.isArray(payload?.units) ? payload.units : [];
        const rows = await resolveScheduleRows(env, units);
        return json({ rows });
      } catch (e) {
        return json({ error: e.message || 'Unable to resolve preview schedule.' }, 500);
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
            'Content-Disposition': 'attachment; filename="SSR-Schedule-Export.xlsx"',
          },
        });
      } catch (e) {
        return new Response(e.message || 'Export failed', { status: 500 });
      }
    }

    return env.ASSETS.fetch(request);
  },
};
