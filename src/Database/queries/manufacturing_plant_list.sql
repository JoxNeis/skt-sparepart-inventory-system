-- =====================================================================
-- Query   : manufacturing_plant_list
-- Type    : SELECT
-- Purpose : All plants that are not archived, with a "code - name"
--           label for drop-downs (same as sparepart_list.display_name).
-- Replaces: select_manufacturing_plants
-- =====================================================================
SELECT
	p.id AS id,
	p.code AS code,
	p.[name] AS [name],
	p.address AS address,
	p.[type] AS [type],
	p.region AS region,
	p.code & ' - ' & p.[name] AS display_name
FROM
	manufacturing_plants AS p
WHERE
	p.deleted_at IS NULL
ORDER BY
	p.[name];
