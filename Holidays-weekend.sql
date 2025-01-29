WITH holidays AS (
    SELECT date_column,
           -- New Year's Day (January 1st)
           date_trunc('year', date_column) + INTERVAL '0 days' AS new_year,
           
           -- Martin Luther King Jr. Day (Third Monday of January)
           date_trunc('year', date_column) + INTERVAL '14 days' +
           (8 - EXTRACT(DOW FROM date_trunc('year', date_column) + INTERVAL '14 days')) % 7 AS mlk_day,
           
           -- Presidents' Day (Third Monday of February)
           date_trunc('year', date_column) + INTERVAL '31 days' +
           (8 - EXTRACT(DOW FROM date_trunc('year', date_column) + INTERVAL '31 days')) % 7 AS presidents_day,
           
           -- Memorial Day (Last Monday of May)
           date_trunc('year', date_column) + INTERVAL '151 days' -
           (EXTRACT(DOW FROM date_trunc('year', date_column) + INTERVAL '151 days') + 6) % 7 AS memorial_day,
           
           -- Independence Day (July 4th)
           date_trunc('year', date_column) + INTERVAL '185 days' AS independence_day,
           
           -- Labor Day (First Monday of September)
           date_trunc('year', date_column) + INTERVAL '244 days' +
           (8 - EXTRACT(DOW FROM date_trunc('year', date_column) + INTERVAL '244 days')) % 7 AS labor_day,
           
           -- Columbus Day (Second Monday of October)
           date_trunc('year', date_column) + INTERVAL '275 days' +
           (8 - EXTRACT(DOW FROM date_trunc('year', date_column) + INTERVAL '275 days')) % 7 AS columbus_day,
           
           -- Veterans Day (November 11th)
           date_trunc('year', date_column) + INTERVAL '314 days' AS veterans_day,
           
           -- Thanksgiving Day (Fourth Thursday of November)
           date_trunc('year', date_column) + INTERVAL '325 days' +
           (11 - EXTRACT(DOW FROM date_trunc('year', date_column) + INTERVAL '325 days')) % 7 AS thanksgiving_day,
           
           -- Christmas Day (December 25th)
           date_trunc('year', date_column) + INTERVAL '358 days' AS christmas_day
    FROM your_table
)
SELECT t.date_column
FROM your_table t
JOIN holidays h ON t.date_column IN (
    h.new_year, h.mlk_day, h.presidents_day, h.memorial_day, 
    h.independence_day, h.labor_day, h.columbus_day, 
    h.veterans_day, h.thanksgiving_day, h.christmas_day
)
OR EXTRACT(DOW FROM t.date_column) IN (0, 6); -- Weekend (Sunday=0, Saturday=6)
