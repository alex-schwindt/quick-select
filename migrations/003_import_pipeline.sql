CREATE TABLE IF NOT EXISTS import_batches (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  source_filename TEXT NOT NULL,
  source_sheet TEXT NOT NULL DEFAULT 'Schedule',
  vendor TEXT,
  product_line TEXT,
  notes TEXT,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS unit_models_v2 (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  model_number TEXT NOT NULL UNIQUE,
  family_key TEXT NOT NULL,
  family_label TEXT NOT NULL,
  tonnage_key TEXT,
  tonnage_value REAL,
  voltage_key TEXT,
  voltage_label TEXT,
  aux_heat_type_key TEXT,
  aux_heat_type_label TEXT,
  aux_heat_capacity_key TEXT NOT NULL DEFAULT '',
  aux_heat_capacity_label TEXT NOT NULL DEFAULT '',
  efficiency_key TEXT NOT NULL,
  efficiency_label TEXT NOT NULL,
  source_batch_id INTEGER,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY (source_batch_id) REFERENCES import_batches(id)
);

CREATE TABLE IF NOT EXISTS staging_schedule_rows (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  batch_id INTEGER NOT NULL,
  source_row_number INTEGER NOT NULL,
  source_descriptor TEXT,
  raw_model_number TEXT,
  raw_brand TEXT,
  raw_qty TEXT,
  raw_airflow_cfm TEXT,
  raw_supply_fan_hp TEXT,
  raw_supply_fan_esp_in_wg TEXT,
  raw_supply_fan_rpm TEXT,
  raw_cooling_total_mbh TEXT,
  raw_cooling_sensible_mbh TEXT,
  raw_unit_eer TEXT,
  raw_seer_ieer TEXT,
  raw_refrigerant TEXT,
  raw_heating_input_mbh TEXT,
  raw_heating_output_mbh TEXT,
  raw_voltage TEXT,
  raw_mca TEXT,
  raw_mocp TEXT,
  raw_weight_lbs TEXT,
  raw_remarks TEXT,
  family_key TEXT,
  family_label TEXT,
  efficiency_key TEXT,
  efficiency_label TEXT,
  tonnage_key TEXT,
  tonnage_value REAL,
  voltage_key TEXT,
  voltage_label TEXT,
  aux_heat_type_key TEXT,
  aux_heat_type_label TEXT,
  aux_heat_capacity_key TEXT,
  aux_heat_capacity_label TEXT,
  parse_status TEXT NOT NULL DEFAULT 'parsed',
  parse_notes TEXT,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY (batch_id) REFERENCES import_batches(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS import_model_results (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  batch_id INTEGER NOT NULL,
  staging_row_id INTEGER NOT NULL,
  model_number TEXT,
  unit_model_id INTEGER,
  action TEXT NOT NULL,
  reason TEXT,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY (batch_id) REFERENCES import_batches(id) ON DELETE CASCADE,
  FOREIGN KEY (staging_row_id) REFERENCES staging_schedule_rows(id) ON DELETE CASCADE,
  FOREIGN KEY (unit_model_id) REFERENCES unit_models_v2(id)
);

CREATE INDEX IF NOT EXISTS idx_staging_batch_parse
ON staging_schedule_rows (batch_id, parse_status);

CREATE INDEX IF NOT EXISTS idx_staging_batch_model
ON staging_schedule_rows (batch_id, raw_model_number);

CREATE INDEX IF NOT EXISTS idx_staging_batch_lookup
ON staging_schedule_rows (batch_id, family_label, tonnage_value, voltage_label);

CREATE INDEX IF NOT EXISTS idx_unit_models_v2_model_number
ON unit_models_v2 (model_number);
