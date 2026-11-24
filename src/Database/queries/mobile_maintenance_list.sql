-- =====================================================================
-- Query   : mobile_maintenance_list
-- Type    : SELECT
-- Purpose : The plant's mobile maintenances that are not archived.
-- Replaces: select_mobile_post_by_manufacturing_plant_name
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_id] Long;

SELECT
	mm.id AS id,
	mm.code AS code,
	mm.[name] AS [name],
	mm.code & ' - ' & mm.[name] AS display
FROM
	mobile_maintenances AS mm
WHERE
	mm.manufacturing_plant_id = [p_manufacturing_plant_id]
	AND mm.deleted_at IS NULL
ORDER BY
	mm.[name];
