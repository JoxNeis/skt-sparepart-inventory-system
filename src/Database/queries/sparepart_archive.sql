-- =====================================================================
-- Query   : sparepart_archive
-- Type    : UPDATE (soft delete)
-- Purpose : Remove a sparepart from the catalogue lists. New orders
--           and requests can no longer use it; history stays intact.
-- =====================================================================
PARAMETERS
    [p_sparepart_id] Long;

UPDATE
    spareparts
SET
    deleted_at = Now(),
    updated_at = Now()
WHERE
    id = [p_sparepart_id]
    AND deleted_at IS NULL;
