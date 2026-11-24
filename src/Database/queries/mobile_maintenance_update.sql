-- =====================================================================
-- Query   : mobile_maintenance_update
-- Type    : UPDATE
-- Purpose : Overwrite every editable column of one mobile maintenance.
--           The plant is not editable (requests already point here).
--           0 rows affected = not found or archived.
-- =====================================================================
PARAMETERS
	[p_mobile_maintenance_id] Long,
	[p_code] Text(40),
	[p_name] Text(255);

UPDATE
	mobile_maintenances
SET
	code = [p_code],
	[name] = [p_name],
	updated_at = Now()
WHERE
	id = [p_mobile_maintenance_id]
	AND deleted_at IS NULL;
