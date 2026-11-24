-- =====================================================================
-- Query   : user_archive
-- Type    : UPDATE (soft delete)
-- Purpose : Deactivate a user. Their history stays intact.
-- =====================================================================
PARAMETERS
    [p_user_id] Long;

UPDATE
    users
SET
    deleted_at = Now(),
    updated_at = Now()
WHERE
    id = [p_user_id]
    AND deleted_at IS NULL;
