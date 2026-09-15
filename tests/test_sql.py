import sqlite3
import unittest
from analysis import ROOT
from sql_pipeline import populate

class SQLTests(unittest.TestCase):
    def test_rebuild_is_idempotent_and_queries_run(self):
        with sqlite3.connect(':memory:') as conn:
            populate(conn)
            first = conn.execute('SELECT * FROM measurements ORDER BY parameter_id, month').fetchall()
            populate(conn)
            self.assertEqual(first, conn.execute('SELECT * FROM measurements ORDER BY parameter_id, month').fetchall())
            self.assertTrue(first)
            for name in ['monthly_deviations', 'best_month', 'month_over_month', 'data_quality']:
                self.assertTrue(conn.execute((ROOT / f'sql/{name}.sql').read_text()).fetchall())

    def test_best_month_tie_and_null_baseline(self):
        with sqlite3.connect(':memory:') as conn:
            conn.executescript((ROOT / 'sql/schema.sql').read_text())
            conn.execute("INSERT INTO equipment VALUES (1,'Example')")
            conn.executemany('INSERT INTO parameters VALUES (?,?,?,?)', [(1,1,'Pressure',100),(2,1,'Zero baseline',0)])
            conn.executemany('INSERT INTO measurements VALUES (?,?,?)', [(1,'2025-01-01',99),(1,'2025-02-01',101),(2,'2025-01-01',10)])
            rows=conn.execute((ROOT / 'sql/best_month.sql').read_text()).fetchall()
            self.assertEqual(rows, [('Example','Pressure','2025-01-01',-1.0)])
            rows=conn.execute((ROOT / 'sql/monthly_deviations.sql').read_text()).fetchall()
            self.assertIsNone(rows[-1][-1])

if __name__ == '__main__':
    unittest.main()
