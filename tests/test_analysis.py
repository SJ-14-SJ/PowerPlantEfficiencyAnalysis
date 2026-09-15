import unittest
import numpy as np
import pandas as pd
from analysis import MONTHS, ROOT, deviation_percent, forecast_values, load_measurements, summarize

class AnalysisTests(unittest.TestCase):
    def test_missing_months_keep_calendar_positions(self):
        np.testing.assert_allclose(forecast_values([10, np.nan, 14, 16, 18, np.nan], 2), [22, 24])

    def test_insufficient_observations_rejected(self):
        with self.assertRaises(ValueError):
            forecast_values([1, np.nan, 3])

    def test_zero_and_missing_design(self):
        self.assertTrue(np.isnan(deviation_percent(100, 0)))
        self.assertTrue(np.isnan(deviation_percent(np.nan, 100)))
        self.assertAlmostEqual(deviation_percent(95, 100), -5)

    def test_best_observed_is_not_average(self):
        frame = pd.DataFrame([dict(Parameter='Example', Design=100, Performance=0, **dict(zip(MONTHS, [80, 99, 120, 130, 140, 150])))])
        original = frame.copy(deep=True)
        result = summarize(frame)
        self.assertEqual(result.loc[0, 'Calculated Performance'], 99)
        self.assertAlmostEqual(result.loc[0, '% Deviation'], -1)
        pd.testing.assert_frame_equal(frame, original)

    def test_all_missing_is_unknown(self):
        frame = pd.DataFrame([dict(Parameter='Example', Design=100, Performance=0, **dict.fromkeys(MONTHS, np.nan))])
        self.assertTrue(pd.isna(summarize(frame).loc[0, 'Significant Deviation']))

    def test_bundled_workbooks(self):
        for file, sheet in [('project.xlsx', 'Turbine'), ('Boiler.xlsx', 'Boiler')]:
            frame = load_measurements(ROOT / file, sheet)
            self.assertFalse(frame.empty)
            for _, row in frame.iterrows():
                self.assertEqual(len(forecast_values(row[MONTHS])), 6)

if __name__ == '__main__':
    unittest.main()
