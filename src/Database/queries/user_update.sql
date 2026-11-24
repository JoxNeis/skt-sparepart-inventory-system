-- =====================================================================
-- Query   : user_update
-- Type    : UPDATE
-- Purpose : Overwrite every editable column of one user.
--           username is not editable: it is the key stored in every
--           *_by column of the history.
--           0 rows affected = not found or archived.
-- =====================================================================
PARAMETERS
	[p_user_id] Long,
	[p_manufacturing_plant_id] Long,
	[p_first_name] Text(100),
	[p_middle_name] Text(100),
	[p_last_name] Text(100),
	[p_jobdesk] Text(255),
	[p_tier] Long;

UPDATE
	users
SET
	manufacturing_plant_id = [p_manufacturing_plant_id],
	first_name = [p_first_name],
	middle_name = [p_middle_name],
	last_name = [p_last_name],
	jobdesk = [p_jobdesk],
	tier = [p_tier],
	updated_at = Now()
WHERE
	id = [p_user_id]
	AND deleted_at IS NULL;
