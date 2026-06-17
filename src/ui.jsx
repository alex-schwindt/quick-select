import { useEffect, useMemo, useState } from 'react';

const PRODUCT_FAMILIES = ['Heat Pump', 'AC'];
const TONNAGES = [3, 4, 5, 6.5, 7.5, 8.5, 10, 12.5, 15, 17.5, 20, 25, 27.5];
const VOLTAGES = ['208/230/3', '460/3'];

const FAMILY_RULES = {
  'Heat Pump': { efficiencies: ['Standard'], heatTypes: ['None', 'Electric Heat'] },
  AC: { efficiencies: ['Standard', 'High'], heatTypes: ['None', 'Electric Heat', 'Aluminum Gas Heat'] },
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
  27.5: { gas: ['220 MBH', '400 MBH'], electric: ['25 kW', '50 kW', '75 kW'] },
};

function nextDefaultTag(units) {
  let i = 1;
  while (units.some((unit) => unit.tag === `RTU-${i}`)) i += 1;
  return `RTU-${i}`;
}

function slugify(value) {
  return String(value).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
}

function formatVoltageDisplay(value) {
  return value === '208/230/3' ? '208-230/3' : value;
}

function buildSelectionCode(state) {
  const familyCode = state.family === 'Heat Pump' ? 'HP' : 'AC';
  const efficiencyCode = state.efficiency === 'High' ? 'HI' : 'STD';
  const tonnageCode = String(state.tonnage).replace('.', 'P');
  const voltageCode = state.voltage === '460/3' ? '460' : '208';
  const heatCode =
    state.heatType === 'None'
      ? 'NOHEAT'
      : state.heatType === 'Electric Heat'
      ? `ELEC-${String(state.heatCapacity || '').replace(/\s+/g, '').replace(/[^0-9A-Za-z-]/g, '')}`
      : `GAS-${String(state.heatCapacity || '').replace(/\s+/g, '').replace(/[^0-9A-Za-z-]/g, '')}`;
  const reheatCode = state.hotGasReheat ? 'HGRH' : 'NOHGRH';
  const econCode =
    state.economizer === 'barometric' ? 'ECO-BARO' : state.economizer === 'powered' ? 'ECO-PE' : 'NOECO';
  return `${familyCode}-${efficiencyCode}-${tonnageCode}-${voltageCode}-${heatCode}-${reheatCode}-${econCode}`;
}

function optionSummary(unit) {
  return [
    unit.hotGasReheat ? 'Hot Gas Reheat' : null,
    unit.economizer === 'barometric' ? 'Economizer w/ Barometric Relief' : null,
    unit.economizer === 'powered' ? 'Economizer w/ Powered Exhaust' : null,
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
    remarks: '',
  };
}

function previewFallback(unit) {
  const options = optionSummary(unit).join(', ');
  return {
    tag: unit.tag,
    areaServed: unit.areaServed || '—',
    manufacturer: 'H&H Trecho',
    modelNumber: buildSelectionCode(unit),
    nominalTons: unit.tonnage,
    unitType: unit.family,
    supplyCfm: '—',
    coolingTotal: '—',
    heatingInput: unit.heatType === 'None' ? 'No heat' : `${unit.heatType} ${unit.heatCapacity}`,
    voltPh: unit.voltage,
    mca: '—',
    mocp: '—',
    weight: '—',
    remarks: unit.remarks?.trim() || options || '—',
    matchFound: false,
    cutsheetUrl: '',
    accessoriesUrl: '',
    wiringUrl: '',
    iomUrl: '',
  };
}

export default function App() {
  const [scheduledUnits, setScheduledUnits] = useState([]);
  const [previewRows, setPreviewRows] = useState([]);
  const [editingIndex, setEditingIndex] = useState(null);
  const [formError, setFormError] = useState('');
  const [exportError, setExportError] = useState('');
  const [previewError, setPreviewError] = useState('');
  const [isExporting, setIsExporting] = useState(false);
  const [isResolvingPreview, setIsResolvingPreview] = useState(false);
  const [selection, setSelection] = useState(makeInitialState());

  const [catalogFile, setCatalogFile] = useState(null);
  const [isImportingCatalog, setIsImportingCatalog] = useState(false);
  const [catalogImportError, setCatalogImportError] = useState('');
  const [catalogImportResult, setCatalogImportResult] = useState(null);

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

  useEffect(() => {
    let cancelled = false;
    async function resolvePreview() {
      if (!scheduledUnits.length) {
        setPreviewRows([]);
        setPreviewError('');
        return;
      }
      setIsResolvingPreview(true);
      setPreviewError('');
      try {
        const response = await fetch('/api/preview-schedule', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ units: scheduledUnits }),
        });
        if (!response.ok) {
          const text = await response.text();
          throw new Error(text || 'Unable to resolve preview schedule.');
        }
        const data = await response.json();
        if (!cancelled) setPreviewRows(Array.isArray(data?.rows) ? data.rows : []);
      } catch (error) {
        if (!cancelled) {
          setPreviewError(error.message || 'Unable to resolve preview schedule.');
          setPreviewRows(scheduledUnits.map(previewFallback));
        }
      } finally {
        if (!cancelled) setIsResolvingPreview(false);
      }
    }
    resolvePreview();
    return () => {
      cancelled = true;
    };
  }, [scheduledUnits]);

  const selectedOptionLabels = optionSummary(selection);
  const selectionCode = useMemo(() => buildSelectionCode(selection), [selection]);
  const currentResolvedRow = editingIndex !== null ? previewRows[editingIndex] : null;
  const currentCutsheetHref =
    currentResolvedRow?.cutsheetUrl || `https://selections.hhtrecho.com/cutsheets/${slugify(selectionCode)}.pdf`;

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
    const duplicate = scheduledUnits.some(
      (unit, index) => unit.tag.trim().toLowerCase() === tag.toLowerCase() && index !== editingIndex,
    );
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
      setScheduledUnits((current) =>
        current.map((unit, index) => (index === editingIndex ? { ...selection, tag } : unit)),
      );
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

  async function importCatalogFile() {
    if (!catalogFile) {
      setCatalogImportError('Please choose a catalog workbook or CSV first.');
      return;
    }
    setIsImportingCatalog(true);
    setCatalogImportError('');
    setCatalogImportResult(null);
    try {
      const formData = new FormData();
      formData.append('file', catalogFile);
      const response = await fetch('/api/admin/import-catalog', {
        method: 'POST',
        body: formData,
      });
      const data = await response.json().catch(() => null);
      if (!response.ok) {
        const issues = Array.isArray(data?.issues) ? ` ${data.issues.join(' | ')}` : '';
        throw new Error((data?.error || 'Catalog import failed.') + issues);
      }
      setCatalogImportResult(data);
    } catch (error) {
      setCatalogImportError(error.message || 'Unable to import catalog file.');
    } finally {
      setIsImportingCatalog(false);
    }
  }

  async function exportWorkbook() {
    if (!scheduledUnits.length) return;
    setIsExporting(true);
    setExportError('');
    try {
      const response = await fetch('/api/export-schedule', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ units: scheduledUnits }),
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
      <header className="hero">
        <div>
          <p className="eyebrow">H&amp;H Trecho</p>
          <h1>Quick Select</h1>
          <p className="hero-copy">
            Build a schedule from standard rooftop selections, preview matched catalog data, and export the SSR workbook.
          </p>
        </div>
      </header>

      <section className="panel">
        <div className="panel-heading">
          <div>
            <h2>Catalog admin</h2>
            <p>Upload a normalized catalog file to refresh the product database used by preview and export.</p>
          </div>
        </div>

        <div className="upload-row">
          <input
            type="file"
            accept=".csv,.xlsx,.xls"
            onChange={(event) => setCatalogFile(event.target.files?.[0] || null)}
          />
          <button type="button" onClick={importCatalogFile} disabled={isImportingCatalog}>
            {isImportingCatalog ? 'Importing…' : 'Import catalog'}
          </button>
        </div>

        {catalogFile ? <p className="status-text">Selected file: {catalogFile.name}</p> : null}
        {catalogImportError ? <p className="error-text">{catalogImportError}</p> : null}
        {catalogImportResult ? (
          <div className="import-summary">
            <p>
              Catalog import complete: {catalogImportResult.rows_read ?? 0} rows read, {catalogImportResult.inserted ?? 0} inserted,
              {` ${catalogImportResult.updated ?? 0} updated.`}
            </p>
          </div>
        ) : null}
      </section>

      <div className="content-grid">
        <section className="panel form-panel">
          <div className="panel-heading">
            <div>
              <h2>{editingIndex === null ? 'Add unit' : `Edit ${selection.tag}`}</h2>
              <p>Create one unit at a time, then add it to the project schedule.</p>
            </div>
          </div>

          <div className="form-grid">
            <label>
              <span>Tag</span>
              <input value={selection.tag} onChange={(e) => updateField('tag', e.target.value)} />
            </label>
            <label>
              <span>Area served</span>
              <input value={selection.areaServed} onChange={(e) => updateField('areaServed', e.target.value)} />
            </label>
            <label>
              <span>Family</span>
              <select value={selection.family} onChange={(e) => updateField('family', e.target.value)}>
                {PRODUCT_FAMILIES.map((value) => (
                  <option key={value} value={value}>
                    {value}
                  </option>
                ))}
              </select>
            </label>
            <label>
              <span>Efficiency</span>
              <select value={selection.efficiency} onChange={(e) => updateField('efficiency', e.target.value)}>
                {familyRule.efficiencies.map((value) => (
                  <option key={value} value={value}>
                    {value}
                  </option>
                ))}
              </select>
            </label>
            <label>
              <span>Tonnage</span>
              <select value={selection.tonnage} onChange={(e) => updateField('tonnage', Number(e.target.value))}>
                {TONNAGES.map((value) => (
                  <option key={value} value={value}>
                    {value}
                  </option>
                ))}
              </select>
            </label>
            <label>
              <span>Voltage</span>
              <select value={selection.voltage} onChange={(e) => updateField('voltage', e.target.value)}>
                {VOLTAGES.map((value) => (
                  <option key={value} value={value}>
                    {formatVoltageDisplay(value)}
                  </option>
                ))}
              </select>
            </label>
            <label>
              <span>Heat type</span>
              <select value={selection.heatType} onChange={(e) => updateField('heatType', e.target.value)}>
                {familyRule.heatTypes.map((value) => (
                  <option key={value} value={value}>
                    {value}
                  </option>
                ))}
              </select>
            </label>
            <label>
              <span>Heat capacity</span>
              <select
                value={selection.heatCapacity}
                onChange={(e) => updateField('heatCapacity', e.target.value)}
                disabled={selection.heatType === 'None'}
              >
                {selection.heatType === 'None' ? <option value="">No heat</option> : null}
                {heatOptions.map((value) => (
                  <option key={value} value={value}>
                    {value}
                  </option>
                ))}
              </select>
            </label>
            <label className="checkbox-field">
              <input
                type="checkbox"
                checked={selection.hotGasReheat}
                onChange={(e) => updateField('hotGasReheat', e.target.checked)}
              />
              <span>Hot gas reheat</span>
            </label>
            <label>
              <span>Economizer</span>
              <select value={selection.economizer} onChange={(e) => updateField('economizer', e.target.value)}>
                <option value="none">None</option>
                <option value="barometric">Barometric relief</option>
                <option value="powered">Powered exhaust</option>
              </select>
            </label>
            <label className="full-span">
              <span>Remarks</span>
              <textarea rows="3" value={selection.remarks} onChange={(e) => updateField('remarks', e.target.value)} />
            </label>
          </div>

          <div className="selection-code-box">
            <span>Selection code</span>
            <strong>{selectionCode}</strong>
            {selectedOptionLabels.length ? <small>{selectedOptionLabels.join(', ')}</small> : null}
          </div>

          <div className="actions-row">
            <button type="button" onClick={saveUnit}>
              {editingIndex === null ? 'Add to schedule' : 'Save changes'}
            </button>
            <button type="button" className="ghost-button" onClick={() => resetForm()}>
              Reset
            </button>
            <a href={currentCutsheetHref} target="_blank" rel="noopener noreferrer" className="ghost-link">
              Open cutsheet
            </a>
          </div>

          {formError ? <p className="error-text">{formError}</p> : null}
        </section>

        <section className="panel schedule-panel">
          <div className="panel-heading">
            <div>
              <h2>Schedule</h2>
              <p>
                {scheduledUnits.length} unit{scheduledUnits.length === 1 ? '' : 's'} in schedule.
              </p>
            </div>
            <button type="button" onClick={exportWorkbook} disabled={!scheduledUnits.length || isExporting}>
              {isExporting ? 'Exporting…' : 'Export workbook'}
            </button>
          </div>

          {exportError ? <p className="error-text">{exportError}</p> : null}
          {previewError ? <p className="error-text">{previewError}</p> : null}
          {isResolvingPreview ? <p className="status-text">Refreshing preview…</p> : null}

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Tag</th>
                  <th>Area Served</th>
                  <th>Family</th>
                  <th>Model Number</th>
                  <th>Tons</th>
                  <th>CFM</th>
                  <th>Cooling</th>
                  <th>Heat</th>
                  <th>Voltage</th>
                  <th>MCA</th>
                  <th>MOCP</th>
                  <th>Weight</th>
                  <th>Remarks</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {!scheduledUnits.length ? (
                  <tr>
                    <td colSpan="14" className="empty-state-cell">
                      Add a unit to begin building the schedule.
                    </td>
                  </tr>
                ) : (
                  scheduledUnits.map((unit, index) => {
                    const row = previewRows[index] || previewFallback(unit);
                    return (
                      <tr key={unit.tag}>
                        <td>{row.tag}</td>
                        <td>{row.areaServed || '—'}</td>
                        <td>{unit.family || '—'}</td>
                        <td>
                          {row.modelNumber}
                          {!row.matchFound ? ' *' : ''}
                        </td>
                        <td>{row.nominalTons || '—'}</td>
                        <td>{row.supplyCfm || '—'}</td>
                        <td>{row.coolingTotal || '—'}</td>
                        <td>{row.heatingInput || '—'}</td>
                        <td>{row.voltPh || '—'}</td>
                        <td>{row.mca || '—'}</td>
                        <td>{row.mocp || '—'}</td>
                        <td>{row.weight || '—'}</td>
                        <td>{row.remarks || '—'}</td>
                        <td>
                          <div className="row-actions">
                            <button type="button" className="link-button" onClick={() => editUnit(index)}>
                              Edit
                            </button>
                            <button type="button" className="link-button" onClick={() => duplicateUnit(index)}>
                              Copy
                            </button>
                            <button type="button" className="link-button danger" onClick={() => removeUnit(index)}>
                              Delete
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {scheduledUnits.length ? (
            <p className="status-text">Rows marked with * are showing fallback selection codes because no catalog match was returned.</p>
          ) : null}
        </section>
      </div>
    </div>
  );
}
