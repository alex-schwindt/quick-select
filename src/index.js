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
  return String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/(^-|-$)/g, '');
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
    state.economizer === 'barometric'
      ? 'ECO-BARO'
      : state.economizer === 'powered'
        ? 'ECO-PE'
        : 'NOECO';

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
    iomUrl: ''
  };
}

function SchedulePreviewTable({ rows }) {
  if (!rows.length) {
    return <div className="empty-state">No schedule rows yet. Add a unit to see the live preview.</div>;
  }

  return (
    <div className="schedule-preview-shell">
      <div className="schedule-preview-scroll">
        <table className="schedule-preview-table">
          <thead>
            <tr className="schedule-preview-column-row">
              <th>Tag</th>
              <th>Area Served</th>
              <th>Manufacturer</th>
              <th>Model Number</th>
              <th>Tons</th>
              <th>Unit Type</th>
              <th>Supply CFM</th>
              <th>Cooling Total</th>
              <th>Heating Input</th>
              <th>Voltage</th>
              <th>MCA</th>
              <th>MOCP</th>
              <th>Weight</th>
              <th>Remarks</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={`${row.tag}-${row.modelNumber}`}>
                <td className="schedule-cell-tag">{row.tag}</td>
                <td>{row.areaServed || '—'}</td>
                <td>{row.manufacturer || '—'}</td>
                <td className="schedule-cell-code">
                  {row.modelNumber}
                  {!row.matchFound ? ' *' : ''}
                </td>
                <td>{row.nominalTons || '—'}</td>
                <td>{row.unitType || '—'}</td>
                <td>{row.supplyCfm || '—'}</td>
                <td>{row.coolingTotal || '—'}</td>
                <td>{row.heatingInput || '—'}</td>
                <td>{row.voltPh || '—'}</td>
                <td>{row.mca || '—'}</td>
                <td>{row.mocp || '—'}</td>
                <td>{row.weight || '—'}</td>
                <td>{row.remarks || '—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="helper-text schedule-actions-list">
        Rows marked with * are showing fallback selection codes because no database match was returned.
      </p>
    </div>
  );
}

function ImportSummary({ result }) {
  if (!result?.ok) return null;

  const summary = result.summary || {};
  const issues = Array.isArray(result.issues) ? result.issues : [];

  return (
    <div className="import-summary">
      <div className="import-summary-grid">
        <div className="import-stat">
          <span className="meta-label">Batch</span>
          <strong>#{result.batch?.id}</strong>
        </div>
        <div className="import-stat">
          <span className="meta-label">Rows staged</span>
          <strong>{summary.rows_staged ?? 0}</strong>
        </div>
        <div className="import-stat">
          <span className="meta-label">Unique models</span>
          <strong>{summary.unique_models_in_batch ?? 0}</strong>
        </div>
        <div className="import-stat">
          <span className="meta-label">Inserted</span>
          <strong>{summary.catalog_inserts ?? 0}</strong>
        </div>
        <div className="import-stat">
          <span className="meta-label">Updated</span>
          <strong>{summary.catalog_updates ?? 0}</strong>
        </div>
        <div className="import-stat">
          <span className="meta-label">Unchanged</span>
          <strong>{summary.catalog_unchanged ?? 0}</strong>
        </div>
      </div>

      {!!issues.length && (
        <div className="import-issues">
          <div className="meta-label">Issues</div>
          <ul>
            {issues.map((issue, index) => {
              if (issue.type === 'duplicate_model_in_batch') {
                return (
                  <li key={index}>
                    Duplicate model <code>{issue.model_number}</code> at rows {issue.source_row_numbers?.join(', ')}.
                  </li>
                );
              }

              return <li key={index}>{issue.type}</li>;
            })}
          </ul>
        </div>
      )}
    </div>
  );
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

  const [importFile, setImportFile] = useState(null);
  const [importVendor, setImportVendor] = useState('Tempmaster');
  const [importProductLine, setImportProductLine] = useState('DS Commercial');
  const [importNotes, setImportNotes] = useState('Imported from DS Commercial workbook');
  const [importLoading, setImportLoading] = useState(false);
  const [importResult, setImportResult] = useState(null);
  const [importError, setImportError] = useState('');

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
      setSelection((current) => ({
        ...current,
        heatType: familyRule.heatTypes[0],
        heatCapacity: ''
      }));
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
          body: JSON.stringify({ units: scheduledUnits })
        });

        if (!response.ok) {
          const text = await response.text();
          throw new Error(text || 'Unable to resolve preview schedule.');
        }

        const data = await response.json();
        if (!cancelled) {
          setPreviewRows(Array.isArray(data?.rows) ? data.rows : []);
        }
      } catch (error) {
        if (!cancelled) {
          setPreviewError(error.message || 'Unable to resolve preview schedule.');
          setPreviewRows(scheduledUnits.map(previewFallback));
        }
      } finally {
        if (!cancelled) {
          setIsResolvingPreview(false);
        }
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
      (unit, index) => unit.tag.trim().toLowerCase() === tag.toLowerCase() && index !== editingIndex
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
        current.map((unit, index) => (index === editingIndex ? { ...selection, tag } : unit))
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

  async function handleImportSchedule() {
    try {
      setImportError('');
      setImportResult(null);

      if (!importFile) {
        setImportError('Choose a workbook file first.');
        return;
      }

      setImportLoading(true);

      const uploadForm = new FormData();
      uploadForm.append('file', importFile);

      const uploadRes = await fetch('/api/upload-template', {
        method: 'POST',
        body: uploadForm
      });

      const uploadJson = await uploadRes.json().catch(() => ({}));
      if (!uploadRes.ok || !uploadJson?.ok) {
        throw new Error(uploadJson.error || 'Upload failed');
      }

      const importRes = await fetch('/api/import-schedule', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          source_filename: uploadJson.key,
          source_sheet: 'Schedule',
          vendor: importVendor || undefined,
          product_line: importProductLine || undefined,
          notes: importNotes || undefined
        })
      });

      const importJson = await importRes.json().catch(() => ({}));
      if (!importRes.ok || !importJson?.ok) {
        throw new Error(importJson.error || 'Import failed');
      }

      setImportResult(importJson);
    } catch (error) {
      setImportError(error.message || 'Import failed');
    } finally {
      setImportLoading(false);
    }
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">H&amp;H Trecho</p>
          <h1>Quick Select Schedule Builder</h1>
          <p className="intro-copy">
            Create one unit at a time, build the project schedule, preview matched catalog rows, export the workbook,
            and now import source schedule workbooks directly into the catalog flow.
          </p>
        </div>
        <div className="topbar-actions">
          <div className="domain-pill">{scheduledUnits.length} unit{scheduledUnits.length === 1 ? '' : 's'} in schedule.</div>
        </div>
      </header>

      <main className="workspace multi-layout">
        <section className="panel selector-panel">
          <div className="panel-header">
            <div>
              <h2>{editingIndex === null ? 'Add unit' : 'Edit unit'}</h2>
              <p>Configure the selection and add it to the schedule.</p>
            </div>
          </div>

          <div className="form-grid">
            <label className="form-field">
              <span className="meta-label">Tag</span>
              <input value={selection.tag} onChange={(e) => updateField('tag', e.target.value)} />
            </label>

            <label className="form-field">
              <span className="meta-label">Area served</span>
              <input value={selection.areaServed} onChange={(e) => updateField('areaServed', e.target.value)} />
            </label>

            <label className="form-field">
              <span className="meta-label">Family</span>
              <select value={selection.family} onChange={(e) => updateField('family', e.target.value)}>
                {PRODUCT_FAMILIES.map((value) => (
                  <option key={value} value={value}>{value}</option>
                ))}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Efficiency</span>
              <select value={selection.efficiency} onChange={(e) => updateField('efficiency', e.target.value)}>
                {familyRule.efficiencies.map((value) => (
                  <option key={value} value={value}>{value}</option>
                ))}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Tonnage</span>
              <select value={selection.tonnage} onChange={(e) => updateField('tonnage', Number(e.target.value))}>
                {TONNAGES.map((value) => (
                  <option key={value} value={value}>{value}</option>
                ))}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Voltage</span>
              <select value={selection.voltage} onChange={(e) => updateField('voltage', e.target.value)}>
                {VOLTAGES.map((value) => (
                  <option key={value} value={value}>{formatVoltageDisplay(value)}</option>
                ))}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Heat type</span>
              <select value={selection.heatType} onChange={(e) => updateField('heatType', e.target.value)}>
                {familyRule.heatTypes.map((value) => (
                  <option key={value} value={value}>{value}</option>
                ))}
              </select>
            </label>

            <label className="form-field">
              <span className="meta-label">Heat capacity</span>
              <select
                value={selection.heatCapacity}
                onChange={(e) => updateField('heatCapacity', e.target.value)}
                disabled={selection.heatType === 'None'}
              >
                <option value="">{selection.heatType === 'None' ? 'No heat selected' : 'Select heat size'}</option>
                {heatOptions.map((value) => (
                  <option key={value} value={value}>{value}</option>
                ))}
              </select>
            </label>
          </div>

          <div className="option-section">
            <div className="meta-label">Options</div>
            <div className="chip-grid">
              <button
                type="button"
                className={`chip ${selection.hotGasReheat ? 'chip-active' : ''}`}
                onClick={() => updateField('hotGasReheat', !selection.hotGasReheat)}
              >
                Hot Gas Reheat
              </button>
              <button
                type="button"
                className={`chip ${selection.economizer === 'barometric' ? 'chip-active' : ''}`}
                onClick={() => updateField('economizer', selection.economizer === 'barometric' ? 'none' : 'barometric')}
              >
                Economizer / Barometric
              </button>
              <button
                type="button"
                className={`chip ${selection.economizer === 'powered' ? 'chip-active' : ''}`}
                onClick={() => updateField('economizer', selection.economizer === 'powered' ? 'none' : 'powered')}
              >
                Economizer / Powered Exhaust
              </button>
            </div>
          </div>

          <label className="form-field notes-field">
            <span className="meta-label">Remarks</span>
            <textarea rows="3" value={selection.remarks} onChange={(e) => updateField('remarks', e.target.value)} />
          </label>

          <div className="logic-card">
            <div className="meta-label">Selection code</div>
            <strong>{selectionCode}</strong>
            {!!selectedOptionLabels.length && (
              <ul>
                {selectedOptionLabels.map((item) => <li key={item}>{item}</li>)}
              </ul>
            )}
          </div>

          {formError && <div className="error-banner">{formError}</div>}

          <div className="builder-actions">
            <button type="button" className="primary-btn" onClick={saveUnit}>
              {editingIndex === null ? 'Add to schedule' : 'Save unit'}
            </button>
            <button type="button" className="secondary-btn" onClick={() => resetForm()}>
              Reset form
            </button>
            <a className="secondary-btn link-btn" href={currentCutsheetHref} target="_blank" rel="noopener noreferrer">
              Cutsheet
            </a>
          </div>

          <div className="import-panel">
            <div className="panel-header">
              <div>
                <h2>Schedule import</h2>
                <p>Upload a workbook to R2 and run the DS Commercial importer from the app.</p>
              </div>
            </div>

            <div className="form-grid">
              <label className="form-field">
                <span className="meta-label">Workbook (.xlsx)</span>
                <input
                  type="file"
                  accept=".xlsx"
                  onChange={(e) => setImportFile(e.target.files?.[0] || null)}
                />
              </label>

              <label className="form-field">
                <span className="meta-label">Vendor</span>
                <input value={importVendor} onChange={(e) => setImportVendor(e.target.value)} />
              </label>

              <label className="form-field">
                <span className="meta-label">Product line</span>
                <input value={importProductLine} onChange={(e) => setImportProductLine(e.target.value)} />
              </label>

              <label className="form-field">
                <span className="meta-label">Notes</span>
                <input value={importNotes} onChange={(e) => setImportNotes(e.target.value)} />
              </label>
            </div>

            {importError && <div className="error-banner">{importError}</div>}

            <div className="builder-actions">
              <button
                type="button"
                className="primary-btn"
                disabled={importLoading || !importFile}
                onClick={handleImportSchedule}
              >
                {importLoading ? 'Uploading + importing…' : 'Import workbook'}
              </button>
            </div>

            <ImportSummary result={importResult} />
          </div>
        </section>

        <section className="panel output-panel">
          <div className="panel-header">
            <div>
              <h2>Schedule preview</h2>
              <p>Preview resolved equipment rows before exporting the SSR workbook.</p>
            </div>
            <div className="topbar-actions">
              <button
                type="button"
                className="primary-btn"
                onClick={exportWorkbook}
                disabled={!scheduledUnits.length || isExporting}
              >
                {isExporting ? 'Exporting…' : 'Export workbook'}
              </button>
            </div>
          </div>

          {previewError && <div className="error-banner">{previewError}</div>}
          {exportError && <div className="error-banner">{exportError}</div>}
          {isResolvingPreview && <p className="helper-text">Resolving live preview…</p>}

          <SchedulePreviewTable rows={previewRows} />

          <div className="unit-list">
            {scheduledUnits.length === 0 ? (
              <div className="empty-state">No units added yet. Start by creating your first RTU selection.</div>
            ) : (
              scheduledUnits.map((unit, index) => {
                const row = previewRows[index] || previewFallback(unit);
                const statusText = row.matchFound
                  ? `Matched model ${row.modelNumber}.`
                  : 'Using generated selection code because no model match was returned.';

                return (
                  <article key={`${unit.tag}-${index}`} className="unit-card unit-card-compact">
                    <div className="unit-card-top">
                      <div>
                        <strong>{unit.tag}</strong>
                        <p>{unit.areaServed || 'Area not specified'}</p>
                      </div>
                      <span className="tag-pill">{unit.family}</span>
                    </div>

                    <div className="unit-meta-grid">
                      <div><span className="meta-label">Model</span><div>{row.modelNumber}</div></div>
                      <div><span className="meta-label">Tons</span><div>{row.nominalTons || unit.tonnage}</div></div>
                      <div><span className="meta-label">Voltage</span><div>{row.voltPh || unit.voltage}</div></div>
                      <div><span className="meta-label">Heat</span><div>{unit.heatType === 'None' ? 'No heat' : `${unit.heatType} ${unit.heatCapacity}`}</div></div>
                    </div>

                    <p className="unit-status">{statusText}</p>

                    <div className="unit-actions">
                      <button type="button" className="secondary-btn" onClick={() => editUnit(index)}>Edit</button>
                      <button type="button" className="secondary-btn" onClick={() => duplicateUnit(index)}>Duplicate</button>
                      <button type="button" className="secondary-btn danger-btn" onClick={() => removeUnit(index)}>Remove</button>
                    </div>
                  </article>
                );
              })
            )}
          </div>
        </section>
      </main>
    </div>
  );
}