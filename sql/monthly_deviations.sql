-- One row per equipment, parameter and month. NULL protects zero baselines.
SELECT e.name AS equipment, p.name AS parameter, m.month,
       p.design_value, m.observed_value,
       100.0 * (m.observed_value - p.design_value) / NULLIF(p.design_value, 0) AS deviation_pct
FROM measurements m
JOIN parameters p USING (parameter_id)
JOIN equipment e USING (equipment_id)
ORDER BY e.name, p.name, m.month;
