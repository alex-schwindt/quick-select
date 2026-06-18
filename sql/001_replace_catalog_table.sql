DROP TABLE IF EXISTS unit_documents;
DROP TABLE IF EXISTS unit_models;

CREATE TABLE unit_models (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  model_number TEXT NOT NULL UNIQUE,
  manufacturer TEXT,
  unit_type TEXT NOT NULL,
  nominal_tonnage REAL NOT NULL,
  cfm REAL,
  hp REAL,
  esp REAL,
  rpm REAL,
  cooling_eat_db TEXT,
  cooling_eat_wb TEXT,
  cooling_lat_db TEXT,
  cooling_lat_wb TEXT,
  cooling_total_capacity REAL,
  cooling_sensible_capacity REAL,
  eer TEXT,
  seer_ieer TEXT,
  heating_eat TEXT,
  heating_lat TEXT,
  heating_capacity TEXT,
  heating_gas_input TEXT,
  heatpump_total_capacity TEXT,
  heat_pump_hspf TEXT,
  electric_heat_capacity TEXT,
  voltage TEXT,
  mca TEXT,
  mocp TEXT,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE unit_documents (
  model_id INTEGER PRIMARY KEY,
  cutsheet_url TEXT,
  accessories_url TEXT,
  wiring_url TEXT,
  iom_url TEXT,
  FOREIGN KEY (model_id) REFERENCES unit_models(id) ON DELETE CASCADE
);

CREATE INDEX idx_unit_models_lookup
  ON unit_models (unit_type, nominal_tonnage, voltage);
