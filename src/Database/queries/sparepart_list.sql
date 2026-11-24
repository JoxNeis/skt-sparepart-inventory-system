-- =====================================================================
-- Query   : sparepart_list
-- Type    : SELECT
-- Purpose : The catalogue (not archived), with a "code - name" label
--           for drop-downs (the spareparts.display calculated field).
-- Replaces: select_catalogues
-- =====================================================================
SELECT
	s.id AS id,
	s.code AS code,
	s.[name] AS [name],
	s.alias AS alias,
	s.[display] AS display_name,
	s.category AS category,
	s.applicable AS applicable,
	s.unit_of_measurement AS unit_of_measurement,
	s.status AS [status]
FROM
	spareparts AS s
WHERE
	s.deleted_at IS NULL
ORDER BY
	s.[name];
