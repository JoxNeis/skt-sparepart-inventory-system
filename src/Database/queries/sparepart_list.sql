-- =====================================================================
-- Query   : sparepart_list
-- Type    : SELECT
-- Purpose : The catalogue (not archived), with a "code - name" label
--           for drop-downs (the spareparts.display calculated field).
-- Replaces: select_catalogues
-- =====================================================================
SELECT
    s.id,
    s.code,
    s.[name],
    s.alias,
    s.[display] AS display_name,
    s.category,
    s.applicable,
    s.unit_of_measurement,
    s.status
FROM
    spareparts AS s
WHERE
    s.deleted_at IS NULL
ORDER BY
    s.[name];
