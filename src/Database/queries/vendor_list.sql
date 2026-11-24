-- =====================================================================
-- Query   : vendor_list
-- Type    : SELECT
-- Purpose : All vendors that are not archived, with a "code - name"
--           label for drop-downs (same as sparepart_list.display_name).
-- Replaces: select_vendors
-- =====================================================================
SELECT
	v.id AS id,
	v.code AS code,
	v.[name] AS [name],
	v.company_name AS company_name,
	v.email AS email,
	v.code & ' - ' & v.[name] AS display_name
FROM
	vendors AS v
WHERE
	v.deleted_at IS NULL
ORDER BY
	v.[name];
