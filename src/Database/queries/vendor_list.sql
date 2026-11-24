-- =====================================================================
-- Query   : vendor_list
-- Type    : SELECT
-- Purpose : All vendors that are not archived.
-- Replaces: select_vendors
-- =====================================================================
SELECT
	v.id,
	v.code,
	v.[name],
	v.company_name,
	v.email
FROM
	vendors AS v
WHERE
	v.deleted_at IS NULL
ORDER BY
	v.[name];
