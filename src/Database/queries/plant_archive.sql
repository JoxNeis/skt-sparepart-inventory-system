-- =====================================================================
-- Query   : plant_archive
-- Type    : UPDATE (soft delete)
-- Purpose : Hide a plant from lists. Its history stays intact.
-- =====================================================================
PARAMETERS
    [p_manufacturing_plant_id] Long;

UPDATE
    manufacturing_plants
SET
    deleted_at = Now(),
    updated_at = Now()
WHERE
    id = [p_manufacturing_plant_id]
    AND deleted_at IS NULL;
