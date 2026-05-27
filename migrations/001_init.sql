CREATE TABLE IF NOT EXISTS unit_models (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  family TEXT NOT NULL,
  efficiency TEXT NOT NULL,
  tonnage REAL NOT NULL,
  voltage TEXT NOT NULL,
  heat_type TEXT NOT NULL,
  heat_capacity TEXT NOT NULL DEFAULT '',
  model_code TEXT,
  model_number TEXT NOT NULL,
  unit_type TEXT,
  unit_eer REAL,
  seer_ieer TEXT,
  cooling_cfm REAL,
  cooling_sensible_capacity_mbh REAL,
  cooling_total_capacity_mbh REAL,
  heating_capacity_mbtu TEXT,
  refrigerant_type TEXT,
  refrigerant_charge TEXT,
  mca TEXT,
  mocp TEXT,
  filter_type TEXT,
  operating_weight_lbs REAL,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_unit_models_lookup
ON unit_models (family, efficiency, tonnage, voltage, heat_type, heat_capacity);

CREATE TABLE IF NOT EXISTS unit_documents (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  model_id INTEGER NOT NULL,
  cutsheet_url TEXT,
  accessories_url TEXT,
  wiring_url TEXT,
  iom_url TEXT,
  FOREIGN KEY (model_id) REFERENCES unit_models(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS selection_events (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  event_type TEXT NOT NULL,
  tag TEXT,
  payload_json TEXT
);
