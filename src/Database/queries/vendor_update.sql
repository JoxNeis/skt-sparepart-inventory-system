-- =====================================================================
-- Query   : vendor_update
-- Type    : UPDATE
-- Purpose : Overwrite every editable column of one vendor.
--           0 rows affected = not found or archived.
-- =====================================================================
PARAMETERS
    [p_vendor_id] Long,
    [p_code] Text(40),
    [p_name] Text(255),
    [p_company_name] Text(255),
    [p_email] Text(255);

UPDATE
    vendors
SET
    code = [p_code],
    [name] = [p_name],
    company_name = [p_company_name],
    email = [p_email],
    updated_at = Now()
WHERE
    id = [p_vendor_id]
    AND deleted_at IS NULL;
