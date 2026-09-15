-- Missing values stay NULL. Do not compare against a more distant observed month.
WITH changes AS (
    SELECT e.name AS equipment, p.name AS parameter, m.month, m.observed_value,
           LAG(m.observed_value) OVER (PARTITION BY m.parameter_id ORDER BY m.month) AS previous_value
    FROM measurements m
    JOIN parameters p USING (parameter_id)
    JOIN equipment e USING (equipment_id)
)
SELECT *, observed_value - previous_value AS absolute_change,
       100.0 * (observed_value - previous_value) / NULLIF(previous_value, 0) AS percent_change
FROM changes ORDER BY equipment, parameter, month;
