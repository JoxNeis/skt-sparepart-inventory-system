-- =====================================================================
-- Query   : mobile_maintenance_list
-- Type    : SELECT
-- Purpose : The plant's mobile maintenances that are not archived.
-- Replaces: select_mobile_post_by_manufacturing_plant_name
-- =====================================================================
PARAMETERS
    [p_manufacturing_plant_id] Long;

SELECT
    mm.id,
    mm.code,
    mm.[name]
FROM
    mobile_maintenances AS mm
WHERE
    mm.manufacturing_plant_id = [p_manufacturing_plant_id]
    AND mm.deleted_at IS NULL
ORDER BY
    mm.[name];
