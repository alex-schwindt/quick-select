import { useEffect, useMemo, useState } from 'react';

const PRODUCT_FAMILIES = ['Heat Pump', 'AC'];
const TONNAGES = [3, 4, 5, 6.5, 7.5, 8.5, 10, 12.5, 15, 17.5, 20, 25, 27.5];
const VOLTAGES = ['208/230/3', '460/3'];

const FAMILY_RULES = {
  'Heat Pump': {
    efficiencies: ['Standard'],
    heatTypes: ['None', 'Electric Heat']
  },
  AC: {
    efficiencies: ['Standard', 'High'],
    heatTypes: ['None', 'Electric Heat', 'Aluminum Gas Heat']
  }
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

const INITIAL_STATE = {
  family: 'Heat Pump',
  efficiency: 'Standard',
  tonnage: 7.5,
  voltage: '460/3',
  heatType: 'None',
  heatCapacity: '',
  hotGasReheat: false,
  economizer: 'none'
};

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
      ? `ELEC-${(state.heatCapacity || '').replace(/\s+/g, '').replace(/[^0-9A-Za-z-]/g, '')}`
      : `GAS-${(state.heatCapacity || '').replace(/\s+/g, '').replace(/[^0-9A-Za-z-]/g, '')}`;
  const reheatCode = state.hotGasReheat ? 'HGRH' : 'NOHGRH';
  const econCode = state.economizer === 'barometric' ? 'ECO-BARO' : state.economizer === 'powered' ? 'ECO-PE' : 'NOECO';
  return `${familyCode}-${efficiencyCode}-${tonnageCode}-${voltageCode}-${heatCode}-${reheatCode}-${econCode}`;
}

export default function App() {
  const [selection, setSelection] = useState(INITIAL_STATE);
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

  const selectionCode = useMemo(() => buildSelectionCode(selection), [selection]);
  const cutsheetBase = `https://selections.hhtrecho.com/cutsheets/${slugify(selectionCode)}`;

  const selectedOptionLabels = [
    selection.hotGasReheat ? 'Hot Gas Reheat' : null,
    selection.economizer === 'barometric' ? 'Economizer w/ Barometric Relief' : null,
    selection.economizer === 'powered' ? 'Economizer w/ Powered Exhaust' : null
  ].filter(Boolean);

  const scheduleRows = [
    { label: 'Product Family', value: selection.family },
    { label: 'Efficiency', value: selection.efficiency },
    { label: 'Nominal Tonnage', value: `${selection.tonnage} Tons` },
    { label: 'Voltage', value: selection.voltage },
    { label: 'Heat', value: selection.heatType === 'None' ? 'None' : `${selection.heatType} / ${selection.heatCapacity}` },
    { label: 'Factory Options', value: selectedOptionLabels.length ? selectedOptionLabels.join(', ') : 'None' },
    { label: 'Selection Code', value: selectionCode }
  ];

  function updateField(field, value) {
    setSelection((current) => ({ ...current, [field]: value }));
  }

  function resetSelection() {
    setSelection(INITIAL_STATE);
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">H&H Trecho</p>
          <h1>Quick Select</h1>
          <p className="intro-copy">A clean engineering selector for rooftop equipment schedules and Cloudflare-hosted cut sheets.</p>
        </div>
        <div className="topbar-actions">
          <span className="domain-pill">selections.hhtrecho.com</span>
          <button className="secondary-btn" type="button" onClick={resetSelection}>Reset</button>
        </div>
      </header>

      <main className="workspace">
        <section className="panel selector-panel">
          <div className="panel-header">
            <div>
              <h2>Selection inputs</h2>
              <p>Choose equipment basics, then add heat and accessories.</p>
            </div>
          </div>

          <div className="form-grid">
            <label className="form-field">
              <span className="meta-label">Product family</span>
              <select value={selection.family} onChange={(event) => updateField('family', event.target.value)}>
                {PRODUCT_FAMILIES.map((option) => <option key={option} value={option}>{option}</option>)}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Efficiency</span>
              <select value={selection.efficiency} onChange={(event) => updateField('efficiency', event.target.value)}>
                {familyRule.efficiencies.map((option) => <option key={option} value={option}>{option}</option>)}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Tonnage</span>
              <select value={selection.tonnage} onChange={(event) => updateField('tonnage', Number(event.target.value))}>
                {TONNAGES.map((option) => <option key={option} value={option}>{option}</option>)}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Voltage</span>
              <select value={selection.voltage} onChange={(event) => updateField('voltage', event.target.value)}>
                {VOLTAGES.map((option) => <option key={option} value={option}>{option}</option>)}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Heat type</span>
              <select value={selection.heatType} onChange={(event) => updateField('heatType', event.target.value)}>
                {familyRule.heatTypes.map((option) => <option key={option} value={option}>{option}</option>)}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Heat capacity</span>
              <select value={selection.heatCapacity} onChange={(event) => updateField('heatCapacity', event.target.value)} disabled={selection.heatType === 'None'}>
                {selection.heatType === 'None'
                  ? <option value="">No heat selected</option>
                  : heatOptions.map((option) => <option key={option} value={option}>{option}</option>)}
              </select>
            </label>
          </div>

          <div className="option-section">
            <div>
              <span className="meta-label">Options</span>
              <p className="helper-text">Powered exhaust replaces barometric relief.</p>
            </div>
            <div className="chip-grid">
              <button
                type="button"
                className={selection.hotGasReheat ? 'chip chip-active' : 'chip'}
                onClick={() => updateField('hotGasReheat', !selection.hotGasReheat)}
              >
                Hot Gas Reheat
              </button>
              <button
                type="button"
                className={selection.economizer === 'barometric' ? 'chip chip-active' : 'chip'}
                onClick={() => updateField('economizer', selection.economizer === 'barometric' ? 'none' : 'barometric')}
              >
                Economizer w/ Barometric Relief
              </button>
              <button
                type="button"
                className={selection.economizer === 'powered' ? 'chip chip-active' : 'chip'}
                onClick={() => updateField('economizer', selection.economizer === 'powered' ? 'none' : 'powered')}
              >
                Economizer w/ Powered Exhaust
              </button>
            </div>
          </div>

          <div className="logic-card">
            <span className="meta-label">Current rules</span>
            <ul>
              <li>Heat Pump supports Standard efficiency only.</li>
              <li>AC supports Standard and High efficiency.</li>
              <li>Heat capacities are filtered by tonnage.</li>
              <li>Economizer selections are mutually exclusive.</li>
            </ul>
          </div>
        </section>

        <section className="panel output-panel">
          <div className="panel-header output-header">
            <div>
              <h2>Schedule output</h2>
              <p>Review the selection summary and document links.</p>
            </div>
            <button className="primary-btn" type="button" onClick={() => navigator.clipboard?.writeText(selectionCode)}>
              Copy selection code
            </button>
          </div>

          <div className="schedule-grid">
            {scheduleRows.map((row) => (
              <article className="schedule-card" key={row.label}>
                <span className="meta-label">{row.label}</span>
                <strong>{row.value}</strong>
              </article>
            ))}
          </div>

          <div className="document-panel">
            <span className="meta-label">Document package</span>
            <div className="document-grid">
              <a className="document-link" href={`${cutsheetBase}.pdf`} target="_blank" rel="noopener noreferrer">
                <strong>Unit cut sheet</strong>
                <span>{`${cutsheetBase}.pdf`}</span>
              </a>
              <a className="document-link" href={`${cutsheetBase}-schedule.pdf`} target="_blank" rel="noopener noreferrer">
                <strong>Selection schedule</strong>
                <span>{`${cutsheetBase}-schedule.pdf`}</span>
              </a>
              {selection.hotGasReheat ? (
                <a className="document-link" href="https://selections.hhtrecho.com/cutsheets/accessories/hot-gas-reheat.pdf" target="_blank" rel="noopener noreferrer">
                  <strong>Hot gas reheat accessory</strong>
                  <span>Accessory cut sheet</span>
                </a>
              ) : null}
              {selection.economizer === 'barometric' ? (
                <a className="document-link" href="https://selections.hhtrecho.com/cutsheets/accessories/economizer-barometric-relief.pdf" target="_blank" rel="noopener noreferrer">
                  <strong>Economizer / barometric relief</strong>
                  <span>Accessory cut sheet</span>
                </a>
              ) : null}
              {selection.economizer === 'powered' ? (
                <a className="document-link" href="https://selections.hhtrecho.com/cutsheets/accessories/economizer-powered-exhaust.pdf" target="_blank" rel="noopener noreferrer">
                  <strong>Economizer / powered exhaust</strong>
                  <span>Accessory cut sheet</span>
                </a>
              ) : null}
            </div>
          </div>

          <div className="summary-panel">
            <span className="meta-label">Selection summary</span>
            <p>
              {selection.family} · {selection.efficiency} · {selection.tonnage} Tons · {selection.voltage} · {selection.heatType === 'None' ? 'No heat' : `${selection.heatType} ${selection.heatCapacity}`}
              {selectedOptionLabels.length ? ` · ${selectedOptionLabels.join(' · ')}` : ' · No factory options'}
            </p>
          </div>
        </section>
      </main>
    </div>
  );
}
