SELECT e.name AS equipment, p.name AS parameter,
       COUNT(m.month) AS month_rows,
       COUNT(m.observed_value) AS observed_months,
       SUM(CASE WHEN m.month IS NOT NULL AND m.observed_value IS NULL THEN 1 ELSE 0 END) AS missing_values,
       CASE WHEN p.design_value IS NULL OR p.design_value = 0 THEN 'invalid baseline' ELSE 'available' END AS baseline_status
FROM parameters p
JOIN equipment e USING (equipment_id)
LEFT JOIN measurements m USING (parameter_id)
GROUP BY e.name, p.parameter_id, p.name, p.design_value
ORDER BY e.name, p.name;
