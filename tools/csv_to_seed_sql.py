import csv
import sys
from pathlib import Path

REQUIRED = [
    'family','efficiency','tonnage','voltage','heat_type','heat_capacity','model_code','model_number',
    'unit_type','unit_eer','seer_ieer','cooling_cfm','cooling_total_capacity_mbh','heating_capacity_mbtu',
    'refrigerant_type','mca','mocp','operating_weight_lbs'
]

TABLE_COLS = REQUIRED[:]


def q(value):
    if value is None:
        return 'NULL'
    s = str(value).strip()
    if s == '':
        return 'NULL'
    return "'" + s.replace("'", "''") + "'"


def main():
    if len(sys.argv) != 3:
        print('Usage: python csv_to_seed_sql.py input.csv output.sql')
        sys.exit(1)

    input_csv = Path(sys.argv[1])
    output_sql = Path(sys.argv[2])

    with input_csv.open(newline='', encoding='utf-8-sig') as f:
        rows = list(csv.DictReader(f))

    if not rows:
        raise SystemExit('CSV has no data rows.')

    missing = [c for c in REQUIRED if c not in rows[0].keys()]
    if missing:
        raise SystemExit('Missing required columns: ' + ', '.join(missing))

    lines = []
    lines.append('-- Generated from catalog CSV')
    lines.append('-- Review before running in production')
    lines.append('DELETE FROM unit_documents;')
    lines.append('DELETE FROM unit_models;')
    lines.append('')
    lines.append('INSERT INTO unit_models (')
    lines.append('  ' + ',\n  '.join(TABLE_COLS))
    lines.append(') VALUES')

    values_sql = []
    for row in rows:
        vals = [q(row.get(col, '')) for col in TABLE_COLS]
        values_sql.append('(' + ', '.join(vals) + ')')
    lines.append(',\n'.join(values_sql) + ';')
    lines.append('')
    lines.append('INSERT INTO unit_documents (')
    lines.append('  model_id,')
    lines.append('  cutsheet_url,')
    lines.append('  accessories_url,')
    lines.append('  wiring_url,')
    lines.append('  iom_url')
    lines.append(')')
    lines.append('SELECT')
    lines.append('  id,')
    lines.append("  'https://selections.hhtrecho.com/cutsheets/' || lower(model_number) || '.pdf',")
    lines.append("  'https://selections.hhtrecho.com/accessories/' || lower(model_number) || '.pdf',")
    lines.append("  'https://selections.hhtrecho.com/wiring/' || lower(model_number) || '.pdf',")
    lines.append("  'https://selections.hhtrecho.com/iom/' || lower(model_number) || '.pdf'")
    lines.append('FROM unit_models;')
    lines.append('')

    output_sql.write_text('\n'.join(lines), encoding='utf-8')
    print(f'Wrote {output_sql}')
    print(f'Rows: {len(rows)}')


if __name__ == '__main__':
    main()
