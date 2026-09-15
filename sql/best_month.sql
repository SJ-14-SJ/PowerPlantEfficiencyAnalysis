-- Deterministic tie-breaking: choose the earliest equally close month.
WITH deviations AS (
    SELECT e.name AS equipment, p.name AS parameter, m.month,
           100.0 * (m.observed_value - p.design_value) / NULLIF(p.design_value, 0) AS deviation_pct
    FROM measurements m
    JOIN parameters p USING (parameter_id)
    JOIN equipment e USING (equipment_id)
    WHERE m.observed_value IS NOT NULL AND p.design_value IS NOT NULL AND p.design_value != 0
), ranked AS (
    SELECT *, ROW_NUMBER() OVER (
        PARTITION BY equipment, parameter ORDER BY ABS(deviation_pct), month
    ) AS position
    FROM deviations
)
SELECT equipment, parameter, month AS closest_month, deviation_pct
FROM ranked WHERE position = 1 ORDER BY equipment, parameter;
