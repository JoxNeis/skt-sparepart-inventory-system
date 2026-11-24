-- =====================================================================
-- Query   : vendor_list
-- Type    : SELECT
-- Purpose : All vendors that are not archived.
-- Replaces: select_vendors
-- =====================================================================
SELECT
	v.id AS id,
	v.code AS code,
	v.[name] AS [name],
	v.company_name AS company_name,
	v.email AS email
FROM
	vendors AS v
WHERE
	v.deleted_at IS NULL
ORDER BY
	v.[name];
