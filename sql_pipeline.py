"""Rebuild an idempotent SQLite analytics dataset from the bundled workbooks."""
import sqlite3
import pandas as pd
from analysis import ROOT, MONTHS, load_measurements

DATES = ['2024-10-01', '2024-11-01', '2024-12-01', '2025-01-01', '2025-02-01', '2025-03-01']


def populate(connection):
    connection.executescript((ROOT / 'sql/schema.sql').read_text())
    with connection:
        # Rebuild the owned snapshot atomically so deleted source rows do not linger.
        connection.execute('DELETE FROM measurements')
        connection.execute('DELETE FROM parameters')
        connection.execute('DELETE FROM equipment')
        for equipment, filename in [('Turbine', 'project.xlsx'), ('Boiler', 'Boiler.xlsx')]:
            cursor = connection.execute('INSERT INTO equipment(name) VALUES (?)', (equipment,))
            equipment_id = cursor.lastrowid
            for _, row in load_measurements(ROOT / filename, equipment).iterrows():
                design = None if pd.isna(row['Design']) else float(row['Design'])
                cursor = connection.execute('INSERT INTO parameters(equipment_id,name,design_value) VALUES (?,?,?)', (equipment_id, str(row['Parameter']).strip(), design))
                parameter_id = cursor.lastrowid
                connection.executemany('INSERT INTO measurements VALUES (?,?,?)', [(parameter_id, date, None if pd.isna(row[month]) else float(row[month])) for month, date in zip(MONTHS, DATES)])


def main():
    output = ROOT / 'outputs'
    output.mkdir(exist_ok=True)
    with sqlite3.connect(output / 'energy.sqlite') as connection:
        populate(connection)
        for name in ['monthly_deviations', 'best_month', 'month_over_month', 'data_quality']:
            frame = pd.read_sql_query((ROOT / f'sql/{name}.sql').read_text(), connection)
            frame.to_csv(output / f'{name}.csv', index=False)
            print(f'{name}: {len(frame)} rows exported')


if __name__ == '__main__':
    main()
