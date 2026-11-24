-- =====================================================================
-- Query   : sparepart_find_by_name
-- Type    : SELECT
-- Purpose : Resolve a name used by a plant to the catalogue sparepart:
--           matches the catalogue code, name or alias.
--           No row = the name is not in the catalogue.
--           matched_on: CATALOGUE (code or name) or ALIAS.
-- =====================================================================
PARAMETERS
	[p_name] Text(255);

SELECT
	s.id,
	s.code,
	s.[name],
	s.alias,
	s.unit_of_measurement,
	IIf(
		s.[name] = [p_name]
		OR s.code = [p_name],
		'CATALOGUE',
		'ALIAS'
	) AS matched_on
FROM
	spareparts AS s
WHERE
	(
		s.[name] = [p_name]
		OR s.code = [p_name]
		OR s.alias = [p_name]
	)
	AND s.deleted_at IS NULL;
