import { useMemo, useState } from 'react';
import './styles.css';

const PRODUCT_FAMILIES = ['Heat Pump', 'AC'];
const TONNAGES = [3, 4, 5, 6.5, 7.5, 8.5, 10, 12.5, 15, 17.5, 20, 25, 27.5];
const VOLTAGES = ['208/230/3', '460/3'];
const EFFICIENCY_BY_FAMILY = {
  'Heat Pump': ['Standard'],
  AC: ['Standard', 'High'],
};
const HEAT_TYPES_BY_FAMILY = {
  'Heat Pump': ['None', 'Electric Heat'],
  AC: ['None', 'Electric Heat', 'Aluminum Gas Heat'],
};
const HEAT_MAP = {
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
  27.5: { gas: ['220 MBH', '400 MBH'], electric: ['25 kW', '50 kW', '75 kW'] },
};
const OPTION_GROUPS = {
  reheat: 'Hot Gas Reheat',
  economizerBarometric: 'Economizer w/ Barometric Relief',
  economizerPowered: 'Economizer w/ Powered Exhaust',
};

function slugify(value) {
  return String(value).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
}

function buildModelNumber({ family, efficiency, tonnage, voltage, heatType, heatCapacity, options }) {
  const familyCode = family === 'Heat Pump' ? 'HP' : 'AC';
  const effCode = efficiency === 'High' ? 'HI' : 'STD';
  const tonCode = String(tonnage).replace('.', 'P');
  const voltageCode = voltage === '460/3' ? '460' : '208';
  const heatCode = heatType === 'Electric Heat'
    ? `E${(heatCapacity || '0').replace(/[^0-9]/g, '')}`
    : heatType === 'Aluminum Gas Heat'
      ? `G${(heatCapacity || '0').replace(/[^0-9]/g, '')}`
      : 'NOHT';
  const optionCode = [
    options.hotGasReheat ? 'HGRH' : null,
    options.economizerPowered ? 'ECPX' : options.economizerBarometric ? 'ECBR' : null,
  ].filter(Boolean).join('-') || 'BASE';
  return `${familyCode}-${effCode}-${tonCode}-${voltageCode}-${heatCode}-${optionCode}`;
}

export default function App() {
  const [family, setFamily] = useState('Heat Pump');
  const [efficiency, setEfficiency] = useState('Standard');
  const [tonnage, setTonnage] = useState(7.5);
  const [voltage, setVoltage] = useState('460/3');
  const [heatType, setHeatType] = useState('None');
  const [heatCapacity, setHeatCapacity] = useState('');
  const [options, setOptions] = useState({
    hotGasReheat: false,
    economizerBarometric: false,
    economizerPowered: false,
  });

  const efficiencyOptions = EFFICIENCY_BY_FAMILY[family];
  const heatTypeOptions = HEAT_TYPES_BY_FAMILY[family];
  const heatChoices = useMemo(() => {
    if (heatType === 'Electric Heat') return HEAT_MAP[tonnage]?.electric ?? [];
    if (heatType === 'Aluminum Gas Heat') return HEAT_MAP[tonnage]?.gas ?? [];
    return [];
  }, [tonnage, heatType]);

  useMemo(() => {
    if (!efficiencyOptions.includes(efficiency)) setEfficiency(efficiencyOptions[0]);
    if (!heatTypeOptions.includes(heatType)) {
      setHeatType(heatTypeOptions[0]);
      setHeatCapacity('');
    }
    if (heatType !== 'None' && heatChoices.length && !heatChoices.includes(heatCapacity)) {
      setHeatCapacity(heatChoices[0]);
    }
    if (heatType === 'None' && heatCapacity) setHeatCapacity('');
  }, [family, efficiency, efficiencyOptions, heatChoices, heatCapacity, heatType, heatTypeOptions]);

  const selectedOptions = [
    options.hotGasReheat ? OPTION_GROUPS.reheat : null,
    options.economizerBarometric ? OPTION_GROUPS.economizerBarometric : null,
    options.economizerPowered ? OPTION_GROUPS.economizerPowered : null,
  ].filter(Boolean);

  const modelNumber = buildModelNumber({ family, efficiency, tonnage, voltage, heatType, heatCapacity, options });
  const cutsheetBase = `https://selections.hhtrecho.com/cutsheets/${slugify(modelNumber)}`;
  const scheduleRows = [
    ['Product Family', family],
    ['Efficiency', efficiency],
    ['Tonnage', `${tonnage} Tons`],
    ['Voltage', voltage],
    ['Heat', heatType === 'None' ? 'None' : `${heatType} / ${heatCapacity}`],
    ['Options', selectedOptions.length ? selectedOptions.join(', ') : 'None'],
    ['Model Number', modelNumber],
  ];

  function setEconomizer(type) {
    setOptions((current) => ({
      ...current,
      economizerBarometric: type === 'barometric',
      economizerPowered: type === 'powered',
    }));
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">Hoffman & Hoffman</p>
          <h1>Quick Select</h1>
          <p className="intro-copy">Rapid rooftop unit selection for engineers using the same deployment stack as your existing internal tools.</p>
        </div>
        <div className="topbar-meta">
          <span className="sync-time">Domain target: selections.hhtrecho.com</span>
          <button
            className="secondary-btn"
            type="button"
            onClick={() => {
              navigator.clipboard?.writeText(modelNumber);
            }}
          >
            Copy model
          </button>
        </div>
      </header>

      <main className="workspace">
        <section className="list-panel selector-panel">
          <div className="panel-header">
            <h2>Selection inputs</h2>
            <span>{family} / {tonnage} Tons</span>
          </div>

          <div className="selector-grid">
            <label className="form-field">
              <span className="meta-label">Product family</span>
              <select value={family} onChange={(event) => setFamily(event.target.value)}>
                {PRODUCT_FAMILIES.map((option) => <option key={option} value={option}>{option}</option>)}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Efficiency</span>
              <select value={efficiency} onChange={(event) => setEfficiency(event.target.value)}>
                {efficiencyOptions.map((option) => <option key={option} value={option}>{option}</option>)}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Tonnage</span>
              <select value={tonnage} onChange={(event) => setTonnage(Number(event.target.value))}>
                {TONNAGES.map((option) => <option key={option} value={option}>{option}</option>)}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Voltage</span>
              <select value={voltage} onChange={(event) => setVoltage(event.target.value)}>
                {VOLTAGES.map((option) => <option key={option} value={option}>{option}</option>)}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Heat type</span>
              <select value={heatType} onChange={(event) => setHeatType(event.target.value)}>
                {heatTypeOptions.map((option) => <option key={option} value={option}>{option}</option>)}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Heat capacity</span>
              <select value={heatCapacity} onChange={(event) => setHeatCapacity(event.target.value)} disabled={heatType === 'None'}>
                {heatType === 'None' ? <option value="">No heat selected</option> : heatChoices.map((option) => <option key={option} value={option}>{option}</option>)}
              </select>
            </label>
          </div>

          <div className="stack-section">
            <span className="meta-label">Options</span>
            <div className="option-grid">
              <button
                type="button"
                className={options.hotGasReheat ? 'filter-chip active' : 'filter-chip'}
                onClick={() => setOptions((current) => ({ ...current, hotGasReheat: !current.hotGasReheat }))}
              >
                {OPTION_GROUPS.reheat}
              </button>
              <button
                type="button"
                className={options.economizerBarometric ? 'filter-chip active' : 'filter-chip'}
                onClick={() => setEconomizer(options.economizerBarometric ? null : 'barometric')}
              >
                {OPTION_GROUPS.economizerBarometric}
              </button>
              <button
                type="button"
                className={options.economizerPowered ? 'filter-chip active' : 'filter-chip'}
                onClick={() => setEconomizer(options.economizerPowered ? null : 'powered')}
              >
                {OPTION_GROUPS.economizerPowered}
              </button>
            </div>
          </div>

          <div className="detail-card note-card">
            <span className="meta-label">Current logic</span>
            <ul className="note-list">
              <li>High efficiency is available only on AC.</li>
              <li>Heat Pump supports none or electric heat.</li>
              <li>Powered exhaust replaces barometric relief.</li>
              <li>Heat capacities are filtered by tonnage.</li>
            </ul>
          </div>
        </section>

        <section className="detail-panel">
          <div className="panel-header detail-header">
            <div>
              <h2>Schedule output</h2>
              <p className="job-subtitle">Live engineering schedule and hosted document package.</p>
            </div>
            <a className="primary-btn enabled inline-btn" href={`${cutsheetBase}.pdf`} target="_blank" rel="noopener noreferrer">Open base cut sheet</a>
          </div>

          <div className="detail-grid">
            {scheduleRows.map(([label, value]) => (
              <div className="detail-card" key={label}>
                <span className="meta-label">{label}</span>
                <strong>{value}</strong>
              </div>
            ))}
          </div>

          <div className="stack-section">
            <span className="meta-label">Document package</span>
            <div className="doc-grid">
              <a className="doc-link" href={`${cutsheetBase}.pdf`} target="_blank" rel="noopener noreferrer">
                <strong>Unit cut sheet</strong>
                <span>{`${cutsheetBase}.pdf`}</span>
              </a>
              <a className="doc-link" href={`${cutsheetBase}-schedule.pdf`} target="_blank" rel="noopener noreferrer">
                <strong>Selection schedule</strong>
                <span>{`${cutsheetBase}-schedule.pdf`}</span>
              </a>
              {options.hotGasReheat ? (
                <a className="doc-link" href={`https://selections.hhtrecho.com/cutsheets/accessories/hot-gas-reheat.pdf`} target="_blank" rel="noopener noreferrer">
                  <strong>Hot gas reheat</strong>
                  <span>Accessory cut sheet</span>
                </a>
              ) : null}
              {options.economizerBarometric ? (
                <a className="doc-link" href={`https://selections.hhtrecho.com/cutsheets/accessories/economizer-barometric.pdf`} target="_blank" rel="noopener noreferrer">
                  <strong>Economizer / barometric relief</strong>
                  <span>Accessory cut sheet</span>
                </a>
              ) : null}
              {options.economizerPowered ? (
                <a className="doc-link" href={`https://selections.hhtrecho.com/cutsheets/accessories/economizer-powered-exhaust.pdf`} target="_blank" rel="noopener noreferrer">
                  <strong>Economizer / powered exhaust</strong>
                  <span>Accessory cut sheet</span>
                </a>
              ) : null}
            </div>
          </div>

          <div className="feedback-panel">
            <span className="meta-label">Selection summary</span>
            <div className="feedback-scroll">
              <p>
                {family} / {efficiency} / {tonnage} Tons / {voltage}
                {heatType !== 'None' ? ` / ${heatType} / ${heatCapacity}` : ' / No heat'}
                {selectedOptions.length ? ` / ${selectedOptions.join(' / ')}` : ' / No factory options selected'}
              </p>
            </div>
          </div>
        </section>
      </main>
    </div>
  );
}
