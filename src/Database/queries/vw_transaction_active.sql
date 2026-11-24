-- =====================================================================
-- Query   : vw_transaction_active
-- Type    : SELECT (view - no parameters, used by other queries)
-- Purpose : Transactions that count: activated and not reversed.
--           The one place where "active" is defined.
-- =====================================================================
SELECT
    t.id,
    t.uuid,
    t.transaction_type,
    t.manufacturing_plant_id,
    t.created_by,
    t.created_at,
    t.activated_at
FROM
    transactions AS t
WHERE
    t.activated_at IS NOT NULL
    AND t.reversed_at IS NULL;
