-- =====================================================================
-- Query   : vendor_archive
-- Type    : UPDATE (soft delete)
-- Purpose : Hide a vendor from lists. Existing orders keep it.
-- =====================================================================
PARAMETERS
	[p_vendor_id] Long;

UPDATE
	vendors
SET
	deleted_at = Now(),
	updated_at = Now()
WHERE
	id = [p_vendor_id]
	AND deleted_at IS NULL;
