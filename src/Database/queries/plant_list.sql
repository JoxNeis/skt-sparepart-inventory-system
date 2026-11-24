-- =====================================================================
-- Query   : plant_list
-- Type    : SELECT
-- Purpose : All plants that are not archived.
-- Replaces: select_manufacturing_plants
-- =====================================================================
SELECT
    p.id,
    p.code,
    p.[name],
    p.address
FROM
    manufacturing_plants AS p
WHERE
    p.deleted_at IS NULL
ORDER BY
    p.[name];
