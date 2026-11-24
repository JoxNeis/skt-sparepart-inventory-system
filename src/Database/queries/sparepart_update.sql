-- =====================================================================
-- Query   : sparepart_update
-- Type    : UPDATE
-- Purpose : Overwrite every editable column of one sparepart.
--           0 rows affected = not found or archived.
-- =====================================================================
PARAMETERS
    [p_sparepart_id] Long,
    [p_code] Text(40),
    [p_name] Text(255),
    [p_category] Text(50),
    [p_applicable] Text(255),
    [p_unit_of_measurement] Text(15),
    [p_alias] Text(255),
    [p_status] Text(50);

UPDATE
    spareparts
SET
    code = [p_code],
    [name] = [p_name],
    category = [p_category],
    applicable = [p_applicable],
    unit_of_measurement = [p_unit_of_measurement],
    alias = [p_alias],
    status = [p_status],
    updated_at = Now()
WHERE
    id = [p_sparepart_id]
    AND deleted_at IS NULL;
