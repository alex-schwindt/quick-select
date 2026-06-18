-- Generated from catalog CSV
-- Review before running in production
DELETE FROM unit_documents;
DELETE FROM unit_models;

INSERT INTO unit_models (
  family,
  efficiency,
  tonnage,
  voltage,
  heat_type,
  heat_capacity,
  model_code,
  model_number,
  unit_type,
  unit_eer,
  seer_ieer,
  cooling_cfm,
  cooling_total_capacity_mbh,
  heating_capacity_mbtu,
  refrigerant_type,
  mca,
  mocp,
  operating_weight_lbs
) VALUES
('AC', 'Standard', '15', '208/230/3', 'Aluminum Gas Heat', '220 MBH', 'LV15N1DP2E1AAN82A1', 'LV15N1DP2E1AAN82A1', 'Packaged AC', '10.9', '14.0', '6000', '182.0', '178.2', 'R454B', '80.9', '100', '2115'),
('AC', 'Standard', '15', '208/230/3', 'Aluminum Gas Heat', '400 MBH', 'LV15N3DP2E1AAN82A1', 'LV15N3DP2E1AAN82A1', 'Packaged AC', '10.9', '14.0', '6000', '182.0', '324.0', 'R454B', '80.9', '100', '2115'),
('AC', 'Standard', '30', '208/230/3', 'Aluminum Gas Heat', '400 MBH', 'RV30N1DP2C1AAN12A1', 'RV30N1DP2C1AAN12A1', 'Packaged AC', '10.6', '14.0', '12000', '351.3', '324.0', 'R454B', '168.9', '200', '4942'),
('AC', 'Standard', '30', '208/230/3', 'Aluminum Gas Heat', '620 MBH', 'RV30N3DP2C1AAN12A1', 'RV30N3DP2C1AAN12A1', 'Packaged AC', '10.6', '14.0', '12000', '351.3', '502.2', 'R454B', '168.9', '200', '4942');

INSERT INTO unit_documents (
  model_id,
  cutsheet_url,
  accessories_url,
  wiring_url,
  iom_url
)
SELECT
  id,
  'https://selections.hhtrecho.com/cutsheets/' || lower(model_number) || '.pdf',
  'https://selections.hhtrecho.com/accessories/' || lower(model_number) || '.pdf',
  'https://selections.hhtrecho.com/wiring/' || lower(model_number) || '.pdf',
  'https://selections.hhtrecho.com/iom/' || lower(model_number) || '.pdf'
FROM unit_models;
