PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS equipment (
    equipment_id INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE
);
CREATE TABLE IF NOT EXISTS parameters (
    parameter_id INTEGER PRIMARY KEY,
    equipment_id INTEGER NOT NULL REFERENCES equipment(equipment_id),
    name TEXT NOT NULL,
    design_value REAL,
    UNIQUE(equipment_id, name)
);
CREATE TABLE IF NOT EXISTS measurements (
    parameter_id INTEGER NOT NULL REFERENCES parameters(parameter_id),
    month TEXT NOT NULL CHECK (month GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-01'),
    observed_value REAL,
    PRIMARY KEY(parameter_id, month)
);
CREATE INDEX IF NOT EXISTS measurements_by_month ON measurements(month);
