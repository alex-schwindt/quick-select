INSERT INTO unit_models (
  family, efficiency, tonnage, voltage, heat_type, heat_capacity, model_code, model_number, unit_type,
  unit_eer, seer_ieer, cooling_cfm, cooling_sensible_capacity_mbh, cooling_total_capacity_mbh,
  heating_capacity_mbtu, refrigerant_type, refrigerant_charge, mca, mocp, filter_type, operating_weight_lbs
) VALUES
('Heat Pump', 'Standard', 7.5, '460/3', 'None', '', 'HP-STD-7P5-460-NOHEAT', 'SAHP075460A', 'Packaged Heat Pump', 11.0, '15.2 IEER', 2500, 210, 265, '', 'R-454B', '18 / 18', '72', '90', '2 in. pleated', 1680),
('Heat Pump', 'Standard', 7.5, '460/3', 'Electric Heat', '24 kW', 'HP-STD-7P5-460-ELEC24', 'SAHP075460E24', 'Packaged Heat Pump', 11.0, '15.0 IEER', 2500, 210, 265, '24 kW', 'R-454B', '18 / 18', '92', '110', '2 in. pleated', 1810),
('AC', 'High', 10, '460/3', 'Aluminum Gas Heat', '180 MBH', 'AC-HI-10-460-GAS180', 'SAC100460G180', 'Packaged AC', 11.2, '16.8 IEER', 3200, 275, 340, '180 MBH', 'R-454B', '22 / 22', '88', '110', '2 in. pleated', 2140);

INSERT INTO unit_documents (model_id, cutsheet_url, accessories_url, wiring_url, iom_url)
SELECT id,
  'https://selections.hhtrecho.com/cutsheets/' || lower(replace(model_number, ' ', '-')) || '.pdf',
  'https://selections.hhtrecho.com/accessories/' || lower(replace(model_number, ' ', '-')) || '.pdf',
  'https://selections.hhtrecho.com/wiring/' || lower(replace(model_number, ' ', '-')) || '.pdf',
  'https://selections.hhtrecho.com/iom/' || lower(replace(model_number, ' ', '-')) || '.pdf'
FROM unit_models;
