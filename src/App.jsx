import { useEffect, useMemo, useState } from 'react';

const PRODUCT_FAMILIES = ['Heat Pump', 'AC'];
const TONNAGES = [3, 4, 5, 6.5, 7.5, 8.5, 10, 12.5, 15, 17.5, 20, 25, 27.5];
const VOLTAGES = ['208/230/3', '460/3'];
const FAMILY_RULES = {
  'Heat Pump': { efficiencies: ['Standard'], heatTypes: ['None', 'Electric Heat'] },
  AC: { efficiencies: ['Standard', 'High'], heatTypes: ['None', 'Electric Heat', 'Aluminum Gas Heat'] }
};
const HEAT_BY_TONNAGE = {
  3: { gas: ['60 MBH', '80 MBH'], electric: ['3 kW', '6 kW', '9 kW', '15 kW'] },
  4: { gas: ['60 MBH', '80 MBH', '120 MBH'], electric: ['6 kW', '9 kW', '15 kW', '20 kW'] },
  5: { gas: ['80 MBH', '120 MBH', '160 MBH'], electric: ['6 kW', '9 kW', '15 kW', '20 kW', '24 kW'] },
  6.5: { gas: ['120 MBH', '180 MBH'], electric: ['9 kW', '18 kW', '24 kW', '36 kW'] },
  7.5: { gas: ['120 MBH', '180 MBH'], electric: ['9 kW', '18 kW', '24 kW', '36 kW'] },
  8.5: { gas: ['120 MBH', '180 MBH'], electric: ['9 kW', '18 kW', '24 kW', '36 kW'] },
  10: { gas: ['180 MBH', '240 MBH'], electric: ['18 kW', '24 kW', '36 kW', '54 kW'] },
  12.5: { gas: ['180 MBH', '240 MBH'], electric: ['18 kW', '24 kW', '36 kW', '54 kW'] },
  15: { gas: ['220 MBH', '400 MBH'], electric: ['25 kW', '50 kW', '75 kW'] },
  17.5: { gas: ['220 MBH', '400 MBH'], electric: ['25 kW', '50 kW', '75 kW'] },
  20: { gas: ['220 MBH', '400 MBH'], electric: ['25 kW', '50 kW', '75 kW'] },
  25: { gas: ['220 MBH', '400 MBH'], electric: ['25 kW', '50 kW', '75 kW'] },
  27.5: { gas: ['220 MBH', '400 MBH'], electric: ['25 kW', '50 kW', '75 kW'] }
};

function nextDefaultTag(units) {
  let i = 1;
  while (units.some((unit) => unit.tag === `RTU-${i}`)) i += 1;
  return `RTU-${i}`;
}

function slugify(value) {
  return String(value).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
}

function buildSelectionCode(state) {
  const familyCode = state.family === 'Heat Pump' ? 'HP' : 'AC';
  const efficiencyCode = state.efficiency === 'High' ? 'HI' : 'STD';
  const tonnageCode = String(state.tonnage).replace('.', 'P');
  const voltageCode = state.voltage === '460/3' ? '460' : '208';
  const heatCode = state.heatType === 'None'
    ? 'NOHEAT'
    : state.heatType === 'Electric Heat'
      ? `ELEC-${String(state.heatCapacity || '').replace(/\s+/g, '').replace(/[^0-9A-Za-z-]/g, '')}`
      : `GAS-${String(state.heatCapacity || '').replace(/\s+/g, '').replace(/[^0-9A-Za-z-]/g, '')}`;
  const reheatCode = state.hotGasReheat ? 'HGRH' : 'NOHGRH';
  const econCode = state.economizer === 'barometric' ? 'ECO-BARO' : state.economizer === 'powered' ? 'ECO-PE' : 'NOECO';
  return `${familyCode}-${efficiencyCode}-${tonnageCode}-${voltageCode}-${heatCode}-${reheatCode}-${econCode}`;
}

function optionSummary(unit) {
  return [
    unit.hotGasReheat ? 'Hot Gas Reheat' : null,
    unit.economizer === 'barometric' ? 'Economizer w/ Barometric Relief' : null,
    unit.economizer === 'powered' ? 'Economizer w/ Powered Exhaust' : null
  ].filter(Boolean);
}

function makeInitialState(tag = 'RTU-1') {
  return {
    tag,
    family: 'Heat Pump',
    efficiency: 'Standard',
    tonnage: 7.5,
    voltage: '460/3',
    heatType: 'None',
    heatCapacity: '',
    hotGasReheat: false,
    economizer: 'none',
    areaServed: '',
    remarks: ''
  };
}

export default function App() {
  const [scheduledUnits, setScheduledUnits] = useState([]);
  const [editingIndex, setEditingIndex] = useState(null);
  const [formError, setFormError] = useState('');
  const [exportError, setExportError] = useState('');
  const [isExporting, setIsExporting] = useState(false);
  const [selection, setSelection] = useState(makeInitialState());
  const familyRule = FAMILY_RULES[selection.family];

  const heatOptions = useMemo(() => {
    if (selection.heatType === 'Electric Heat') return HEAT_BY_TONNAGE[selection.tonnage]?.electric ?? [];
    if (selection.heatType === 'Aluminum Gas Heat') return HEAT_BY_TONNAGE[selection.tonnage]?.gas ?? [];
    return [];
  }, [selection.heatType, selection.tonnage]);

  useEffect(() => {
    if (!familyRule.efficiencies.includes(selection.efficiency)) {
      setSelection((current) => ({ ...current, efficiency: familyRule.efficiencies[0] }));
    }
  }, [familyRule.efficiencies, selection.efficiency]);

  useEffect(() => {
    if (!familyRule.heatTypes.includes(selection.heatType)) {
      setSelection((current) => ({ ...current, heatType: familyRule.heatTypes[0], heatCapacity: '' }));
      return;
    }
    if (selection.heatType === 'None' && selection.heatCapacity !== '') {
      setSelection((current) => ({ ...current, heatCapacity: '' }));
      return;
    }
    if (selection.heatType !== 'None' && heatOptions.length > 0 && !heatOptions.includes(selection.heatCapacity)) {
      setSelection((current) => ({ ...current, heatCapacity: heatOptions[0] }));
    }
  }, [familyRule.heatTypes, heatOptions, selection.heatCapacity, selection.heatType]);

  const selectedOptionLabels = optionSummary(selection);
  const selectionCode = useMemo(() => buildSelectionCode(selection), [selection]);
  const cutsheetBase = `https://selections.hhtrecho.com/cutsheets/${slugify(selectionCode)}`;

  function updateField(field, value) {
    setSelection((current) => ({ ...current, [field]: value }));
  }

  function resetForm(nextTag = nextDefaultTag(scheduledUnits)) {
    setSelection(makeInitialState(nextTag));
    setEditingIndex(null);
    setFormError('');
  }

  function saveUnit() {
    const tag = selection.tag.trim();
    if (!tag) {
      setFormError('Tag is required.');
      return;
    }
    const duplicate = scheduledUnits.some((unit, index) => unit.tag.trim().toLowerCase() === tag.toLowerCase() && index !== editingIndex);
    if (duplicate) {
      setFormError('Tag must be unique within the schedule.');
      return;
    }
    setFormError('');
    if (editingIndex === null) {
      const nextUnits = [...scheduledUnits, { ...selection, tag }];
      setScheduledUnits(nextUnits);
      resetForm(nextDefaultTag(nextUnits));
    } else {
      setScheduledUnits((current) => current.map((unit, index) => index === editingIndex ? { ...selection, tag } : unit));
      resetForm();
    }
  }

  function editUnit(index) {
    setSelection({ ...scheduledUnits[index] });
    setEditingIndex(index);
    setFormError('');
  }

  function duplicateUnit(index) {
    const source = scheduledUnits[index];
    setSelection({ ...source, tag: nextDefaultTag(scheduledUnits) });
    setEditingIndex(null);
    setFormError('');
  }

  function removeUnit(index) {
    const nextUnits = scheduledUnits.filter((_, i) => i !== index);
    setScheduledUnits(nextUnits);
    if (editingIndex === index) resetForm(nextDefaultTag(nextUnits));
  }

  async function exportWorkbook() {
    if (!scheduledUnits.length) return;
    setIsExporting(true);
    setExportError('');
    try {
      const response = await fetch('/api/export-schedule', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ units: scheduledUnits })
      });
      if (!response.ok) {
        const text = await response.text();
        throw new Error(text || 'Export failed');
      }
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'SSR-Schedule-Export.xlsx';
      a.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      setExportError(error.message || 'Unable to export workbook.');
    } finally {
      setIsExporting(false);
    }
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">H&H Trecho</p>
          <h1>Quick Select</h1>
          <p className="intro-copy">Build multi-unit schedules, assign custom equipment tags, and export the real Excel template through a Cloudflare Worker.</p>
        </div>
        <div className="topbar-actions">
          <span className="domain-pill">selections.hhtrecho.com</span>
          <button className="secondary-btn" type="button" onClick={() => resetForm()}>Reset form</button>
        </div>
      </header>

      <main className="workspace multi-layout">
        <section className="panel selector-panel">
          <div className="panel-header">
            <div>
              <h2>{editingIndex === null ? 'Unit builder' : `Editing ${selection.tag}`}</h2>
              <p>Create one unit at a time, then add it to the project schedule.</p>
            </div>
          </div>

          <div className="form-grid">
            <label className="form-field"><span className="meta-label">Tag</span><input value={selection.tag} onChange={(event) => updateField('tag', event.target.value)} placeholder="RTU-1" /></label>
            <label className="form-field"><span className="meta-label">Area served</span><input value={selection.areaServed} onChange={(event) => updateField('areaServed', event.target.value)} placeholder="Level 2 Offices" /></label>
            <label className="form-field"><span className="meta-label">Product family</span><select value={selection.family} onChange={(event) => updateField('family', event.target.value)}>{PRODUCT_FAMILIES.map((option) => <option key={option} value={option}>{option}</option>)}</select></label>
            <label className="form-field"><span className="meta-label">Efficiency</span><select value={selection.efficiency} onChange={(event) => updateField('efficiency', event.target.value)}>{familyRule.efficiencies.map((option) => <option key={option} value={option}>{option}</option>)}</select></label>
            <label className="form-field"><span className="meta-label">Tonnage</span><select value={selection.tonnage} onChange={(event) => updateField('tonnage', Number(event.target.value))}>{TONNAGES.map((option) => <option key={option} value={option}>{option}</option>)}</select></label>
            <label className="form-field"><span className="meta-label">Voltage</span><select value={selection.voltage} onChange={(event) => updateField('voltage', event.target.value)}>{VOLTAGES.map((option) => <option key={option} value={option}>{option}</option>)}</select></label>
            <label className="form-field"><span className="meta-label">Heat type</span><select value={selection.heatType} onChange={(event) => updateField('heatType', event.target.value)}>{familyRule.heatTypes.map((option) => <option key={option} value={option}>{option}</option>)}</select></label>
            <label className="form-field"><span className="meta-label">Heat capacity</span><select value={selection.heatCapacity} onChange={(event) => updateField('heatCapacity', event.target.value)} disabled={selection.heatType === 'None'}>{selection.heatType === 'None' ? <option value="">No heat selected</option> : heatOptions.map((option) => <option key={option} value={option}>{option}</option>)}</select></label>
          </div>

          <div className="option-section">
            <div><span className="meta-label">Options</span><p className="helper-text">Powered exhaust replaces barometric relief.</p></div>
            <div className="chip-grid">
              <button type="button" className={selection.hotGasReheat ? 'chip chip-active' : 'chip'} onClick={() => updateField('hotGasReheat', !selection.hotGasReheat)}>Hot Gas Reheat</button>
              <button type="button" className={selection.economizer === 'barometric' ? 'chip chip-active' : 'chip'} onClick={() => updateField('economizer', selection.economizer === 'barometric' ? 'none' : 'barometric')}>Economizer w/ Barometric Relief</button>
              <button type="button" className={selection.economizer === 'powered' ? 'chip chip-active' : 'chip'} onClick={() => updateField('economizer', selection.economizer === 'powered' ? 'none' : 'powered')}>Economizer w/ Powered Exhaust</button>
            </div>
          </div>

          <label className="form-field notes-field"><span className="meta-label">Remarks</span><textarea value={selection.remarks} onChange={(event) => updateField('remarks', event.target.value)} rows={4} placeholder="Optional notes for the exported schedule" /></label>
          {formError ? <div className="error-banner">{formError}</div> : null}
          {exportError ? <div className="error-banner">{exportError}</div> : null}

          <div className="builder-actions">
            <button className="primary-btn" type="button" onClick={saveUnit}>{editingIndex === null ? 'Add unit to schedule' : 'Update unit'}</button>
            <button className="secondary-btn" type="button" onClick={exportWorkbook} disabled={!scheduledUnits.length || isExporting}>{isExporting ? 'Exporting...' : 'Download Excel schedule'}</button>
            <a className="secondary-btn link-btn" href={`${cutsheetBase}.pdf`} target="_blank" rel="noopener noreferrer">Open base cut sheet</a>
          </div>

          <div className="logic-card">
            <span className="meta-label">Current unit summary</span>
            <ul>
              <li>{selection.tag || 'Untitled tag'} · {selection.family} · {selection.efficiency} · {selection.tonnage} Tons</li>
              <li>{selection.voltage} · {selection.heatType === 'None' ? 'No heat' : `${selection.heatType} ${selection.heatCapacity}`}</li>
              <li>{selectedOptionLabels.length ? selectedOptionLabels.join(', ') : 'No factory options selected'}</li>
              <li>Selection code: {selectionCode}</li>
            </ul>
          </div>
        </section>

        <section className="panel output-panel schedule-panel">
          <div className="panel-header output-header">
            <div>
              <h2>Project schedule</h2>
              <p>{scheduledUnits.length} unit{scheduledUnits.length === 1 ? '' : 's'} in schedule.</p>
            </div>
            <button className="primary-btn" type="button" onClick={exportWorkbook} disabled={!scheduledUnits.length || isExporting}>{isExporting ? 'Exporting...' : 'Export Excel'}</button>
          </div>

          {!scheduledUnits.length ? (
            <div className="empty-state">No units added yet. Build a unit on the left and add it to the schedule.</div>
          ) : (
            <div className="unit-list">
              {scheduledUnits.map((unit, index) => {
                const options = optionSummary(unit).join(', ') || 'No options';
                return (
                  <article className="unit-card" key={`${unit.tag}-${index}`}>
                    <div className="unit-card-top">
                      <div><strong>{unit.tag}</strong><p>{unit.areaServed || 'Area served not entered'}</p></div>
                      <span className="tag-pill">{unit.family}</span>
                    </div>
                    <div className="unit-meta-grid">
                      <span>{unit.tonnage} Tons</span><span>{unit.efficiency}</span><span>{unit.voltage}</span><span>{unit.heatType === 'None' ? 'No heat' : `${unit.heatType} ${unit.heatCapacity}`}</span>
                    </div>
                    <p className="unit-options">{options}</p>
                    <div className="unit-actions">
                      <button className="secondary-btn" type="button" onClick={() => editUnit(index)}>Edit</button>
                      <button className="secondary-btn" type="button" onClick={() => duplicateUnit(index)}>Duplicate</button>
                      <button className="secondary-btn danger-btn" type="button" onClick={() => removeUnit(index)}>Remove</button>
                    </div>
                  </article>
                );
              })}
            </div>
          )}
        </section>
      </main>
    </div>
  );
}

