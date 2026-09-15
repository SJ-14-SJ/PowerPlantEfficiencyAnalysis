"""Validated workbook loading and exploratory performance calculations."""
from pathlib import Path
import numpy as np
import pandas as pd
from sklearn.linear_model import LinearRegression

MONTHS = ['October', 'November', 'December', 'January', 'February', 'March']
FORECAST_MONTHS = [f'{m} 2025' for m in ['April', 'May', 'June', 'July', 'August', 'September']]
ROOT = Path(__file__).resolve().parent


def load_measurements(path, sheet_name):
    frame = pd.read_excel(path, sheet_name=sheet_name, skiprows=1)
    frame.columns = frame.columns.astype(str).str.strip()
    frame = frame.rename(columns={'Parameters': 'Parameter', 'Design Data': 'Design', 'Performance Data': 'Performance'})
    required = ['Parameter', 'Design', *MONTHS, 'Performance']
    missing = set(required) - set(frame.columns)
    if missing:
        raise ValueError(f'Missing workbook columns: {sorted(missing)}')
    frame = frame[required].dropna(subset=['Parameter']).copy()
    for col in required[1:]:
        frame[col] = pd.to_numeric(frame[col], errors='coerce')
    return frame


def deviation_percent(observed, design):
    if pd.isna(observed) or pd.isna(design) or not np.isfinite(observed) or not np.isfinite(design) or design == 0:
        return np.nan
    return 100.0 * (observed - design) / design


def summarize(frame):
    result = frame.copy(deep=True)
    def closest(row):
        values = pd.to_numeric(row[MONTHS], errors='coerce')
        values = values[np.isfinite(values)]
        if values.empty or not np.isfinite(row['Design']):
            return np.nan
        return values.loc[(values - row['Design']).abs().idxmin()]
    result['Calculated Performance'] = result.apply(closest, axis=1)
    result['% Deviation'] = result.apply(lambda r: deviation_percent(r['Calculated Performance'], r['Design']), axis=1)
    result['Significant Deviation'] = result['% Deviation'].abs().gt(5).where(result['% Deviation'].notna(), pd.NA).astype('boolean')
    return result


def forecast_values(values, horizon=6):
    """Keep real month positions, including interior and trailing missing months."""
    y = pd.to_numeric(pd.Series(values), errors='coerce').to_numpy(dtype=float)
    mask = np.isfinite(y)
    if mask.sum() < 3:
        raise ValueError('At least three finite observations are required')
    if not isinstance(horizon, int) or horizon < 1:
        raise ValueError('horizon must be a positive integer')
    x = np.arange(len(y))
    model = LinearRegression().fit(x[mask, None], y[mask])
    return model.predict(np.arange(len(y), len(y) + horizon).reshape(-1, 1))


def forecast_frame(frame):
    rows = []
    for _, row in frame.iterrows():
        try:
            prediction = forecast_values(row[MONTHS]); status = 'exploratory'
        except ValueError:
            prediction = [np.nan] * 6; status = 'insufficient observations'
        rows.append([row['Parameter'], *prediction, status])
    return pd.DataFrame(rows, columns=['Parameter', *FORECAST_MONTHS, 'Status'])
