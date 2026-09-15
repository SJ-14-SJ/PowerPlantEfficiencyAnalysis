"""Print and export exploratory forecasts; run from any working directory."""
from analysis import ROOT, forecast_frame, load_measurements


def main():
    output = ROOT / 'outputs'
    output.mkdir(exist_ok=True)
    for source, filename in [('Turbine', 'project.xlsx'), ('Boiler', 'Boiler.xlsx')]:
        predictions = forecast_frame(load_measurements(ROOT / filename, source))
        print(f'\n{source}: exploratory forecasts, not validated operating guidance')
        print(predictions.to_string(index=False))
        predictions.to_excel(output / f'{source}_Linear_Forecast.xlsx', index=False)


if __name__ == '__main__':
    main()
