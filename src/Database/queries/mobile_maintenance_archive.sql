-- =====================================================================
-- Query   : mobile_maintenance_archive
-- Type    : UPDATE (soft delete)
-- Purpose : Hide a mobile maintenance from lists. Existing requests keep it.
-- =====================================================================
PARAMETERS
    [p_mobile_maintenance_id] Long;

UPDATE
    mobile_maintenances
SET
    deleted_at = Now(),
    updated_at = Now()
WHERE
    id = [p_mobile_maintenance_id]
    AND deleted_at IS NULL;
