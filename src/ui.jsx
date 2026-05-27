import { useMemo, useState } from 'react';
import './styles.css';

const TONNAGE_OPTIONS = ['3', '4', '5', '6', '7.5', '8.5', '10'];
const ELECTRICAL_OPTIONS = ['208/230/3', '460/3'];
const HEAT_CAPACITY_OPTIONS = ['0', '30 MBH', '60 MBH', '90 MBH', '120 MBH'];

function buildSelectionCode(unit) {
  const familyCode = unit.family === 'Heat Pump' ? 'HP' : 'AC';
  const efficiencyCode = unit.efficiency === 'High Efficiency' ? 'HI' : 'STD';
  const tonnageCode = String(unit.tonnage).replace('.', 'P');
  const voltageCode = unit.electrical === '460/3' ? '460' : '208';
  const heatCode = unit.heatType === 'Electric Heat'
    ? `ELEC-${unit.heatCapacity.replace(' MBH', '')}`
    : unit.heatType === 'Aluminum Gas Heat'
      ? `GAS-${unit.heatCapacity.replace(' MBH', '')}MBH`
      : 'NOHEAT';
  return `${familyCode}-${efficiencyCode}-${tonnageCode}-${voltageCode}-${heatCode}`;
}

function buildRemarks(unit) {
  return [
    unit.economizer ? 'Economizer' : null,
    unit.curb ? 'Curb' : null,
    unit.hotGasReheat ? 'Hot Gas Reheat' : null,
    unit.poweredExhaust ? 'Powered Exhaust' : null,
  ].filter(Boolean).join(', ');
}

function createUnit(tagNumber) {
  return {
    id: crypto.randomUUID(),
    tag: `RTU-${tagNumber}`,
    areaServed: '',
    family: 'AC',
    efficiency: 'Standard',
    tonnage: '3',
    electrical: '208/230/3',
    heatType: 'Aluminum Gas Heat',
    heatCapacity: '60 MBH',
    economizer: false,
    curb: false,
    hotGasReheat: false,
    poweredExhaust: false,
  };
}

function formatCooling(unit) {
  return `${unit.tonnage} Tons`;
}

function formatHeating(unit) {
  if (unit.heatType === 'No Heat') return 'None';
  return unit.heatCapacity;
}

export default function App() {
  const [units, setUnits] = useState([createUnit(1)]);
  const [activeUnitId, setActiveUnitId] = useState(units[0].id);

  const activeUnit = useMemo(
    () => units.find((unit) => unit.id === activeUnitId) || units[0],
    [units, activeUnitId]
  );

  function updateUnit(id, patch) {
    setUnits((current) => current.map((unit) => (unit.id === id ? { ...unit, ...patch } : unit)));
  }

  function addUnit() {
    const next = createUnit(units.length + 1);
    setUnits((current) => [...current, next]);
    setActiveUnitId(next.id);
  }

  return (
    <div className="app-shell selector-shell">
      <header className="topbar selector-topbar">
        <div>
          <p className="eyebrow">H&H Trecho</p>
          <h1>Quick Select Tool</h1>
        </div>
        <div className="topbar-meta selector-meta">
          <span>{units.length} unit{units.length === 1 ? '' : 's'} in schedule</span>
          <button className="primary-btn enabled compact-btn" onClick={addUnit}>Add unit</button>
        </div>
      </header>

      <main className="selector-workspace">
        <section className="builder-panel">
          <div className="panel-header builder-header">
            <div>
              <h2>Selection Builder</h2>
              <span>Choose the active unit and adjust selection inputs.</span>
            </div>
          </div>

          <div className="unit-chip-row">
            {units.map((unit) => (
              <button
                key={unit.id}
                className={unit.id === activeUnitId ? 'unit-chip active' : 'unit-chip'}
                onClick={() => setActiveUnitId(unit.id)}
              >
                {unit.tag}
              </button>
            ))}
          </div>

          {activeUnit && (
            <div className="builder-form-grid">
              <label className="form-field">
                <span className="meta-label">Tag</span>
                <input value={activeUnit.tag} onChange={(e) => updateUnit(activeUnit.id, { tag: e.target.value })} />
              </label>

              <label className="form-field">
                <span className="meta-label">Area served</span>
                <input value={activeUnit.areaServed} onChange={(e) => updateUnit(activeUnit.id, { areaServed: e.target.value })} />
              </label>

              <label className="form-field">
                <span className="meta-label">Family</span>
                <select value={activeUnit.family} onChange={(e) => updateUnit(activeUnit.id, { family: e.target.value })}>
                  <option>AC</option>
                  <option>Heat Pump</option>
                </select>
              </label>

              <label className="form-field">
                <span className="meta-label">Efficiency</span>
                <select value={activeUnit.efficiency} onChange={(e) => updateUnit(activeUnit.id, { efficiency: e.target.value })}>
                  <option>Standard</option>
                  <option>High Efficiency</option>
                </select>
              </label>

              <label className="form-field">
                <span className="meta-label">Tonnage</span>
                <select value={activeUnit.tonnage} onChange={(e) => updateUnit(activeUnit.id, { tonnage: e.target.value })}>
                  {TONNAGE_OPTIONS.map((option) => <option key={option}>{option}</option>)}
                </select>
              </label>

              <label className="form-field">
                <span className="meta-label">Electrical</span>
                <select value={activeUnit.electrical} onChange={(e) => updateUnit(activeUnit.id, { electrical: e.target.value })}>
                  {ELECTRICAL_OPTIONS.map((option) => <option key={option}>{option}</option>)}
                </select>
              </label>

              <label className="form-field">
                <span className="meta-label">Heat type</span>
                <select value={activeUnit.heatType} onChange={(e) => updateUnit(activeUnit.id, { heatType: e.target.value })}>
                  <option>Aluminum Gas Heat</option>
                  <option>Electric Heat</option>
                  <option>No Heat</option>
                </select>
              </label>

              <label className="form-field">
                <span className="meta-label">Heat capacity</span>
                <select value={activeUnit.heatCapacity} onChange={(e) => updateUnit(activeUnit.id, { heatCapacity: e.target.value })}>
                  {HEAT_CAPACITY_OPTIONS.map((option) => <option key={option}>{option}</option>)}
                </select>
              </label>

              <div className="option-grid">
                {[
                  ['economizer', 'Economizer'],
                  ['curb', 'Curb'],
                  ['hotGasReheat', 'Hot Gas Reheat'],
                  ['poweredExhaust', 'Powered Exhaust'],
                ].map(([key, label]) => (
                  <label key={key} className="toggle-card">
                    <input
                      type="checkbox"
                      checked={activeUnit[key]}
                      onChange={(e) => updateUnit(activeUnit.id, { [key]: e.target.checked })}
                    />
                    <span>{label}</span>
                  </label>
                ))}
              </div>
            </div>
          )}
        </section>

        <section className="preview-panel">
          <div className="panel-header preview-header">
            <div>
              <h2>Schedule Preview</h2>
              <span>High-level live view of the schedule being built.</span>
            </div>
          </div>

          <div className="schedule-preview-card">
            <div className="schedule-preview-table-wrap">
              <table className="schedule-preview-table">
                <thead>
                  <tr className="schedule-group-row">
                    <th colSpan="4">Identification</th>
                    <th colSpan="2">Selection</th>
                    <th colSpan="2">Performance</th>
                    <th colSpan="2">Electrical</th>
                    <th colSpan="1">Remarks</th>
                  </tr>
                  <tr className="schedule-column-row">
                    <th>Tag</th>
                    <th>Area Served</th>
                    <th>Manufacturer</th>
                    <th>Model Number</th>
                    <th>Tons</th>
                    <th>Unit Type</th>
                    <th>Cooling</th>
                    <th>Heating</th>
                    <th>Volt/Ph</th>
                    <th>Selection Code</th>
                    <th>Options</th>
                  </tr>
                </thead>
                <tbody>
                  {units.map((unit) => (
                    <tr key={unit.id} className={unit.id === activeUnitId ? 'active-schedule-row' : ''} onClick={() => setActiveUnitId(unit.id)}>
                      <td>{unit.tag}</td>
                      <td>{unit.areaServed || '—'}</td>
                      <td>H&amp;H Trecho</td>
                      <td className="mono-cell">{buildSelectionCode(unit)}</td>
                      <td>{unit.tonnage}</td>
                      <td>{unit.family}</td>
                      <td>{formatCooling(unit)}</td>
                      <td>{formatHeating(unit)}</td>
                      <td>{unit.electrical}</td>
                      <td className="mono-cell compact-code">{buildSelectionCode(unit)}</td>
                      <td>{buildRemarks(unit) || '—'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {activeUnit && (
            <div className="preview-detail-strip">
              <div className="preview-detail-card">
                <span className="meta-label">Active unit</span>
                <strong>{activeUnit.tag}</strong>
              </div>
              <div className="preview-detail-card">
                <span className="meta-label">Selection</span>
                <strong>{activeUnit.family} · {activeUnit.efficiency} · {activeUnit.tonnage} Tons</strong>
              </div>
              <div className="preview-detail-card">
                <span className="meta-label">Heat / Power</span>
                <strong>{activeUnit.electrical} · {formatHeating(activeUnit)}</strong>
              </div>
            </div>
          )}
        </section>
      </main>
    </div>
  );
}
