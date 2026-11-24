-- =====================================================================
-- Query   : transaction_reverse
-- Type    : UPDATE
-- Purpose : Reverse (cancel) a transaction and record who did it.
--           0 rows affected = not found or already reversed.
-- Replaces: update_transaction_reversal
-- =====================================================================
PARAMETERS
    [p_uuid] Text(36),
    [p_reversed_by] Text(64);

UPDATE
    transactions
SET
    reversed_at = Now(),
    reversed_by = [p_reversed_by]
WHERE
    uuid = [p_uuid]
    AND reversed_at IS NULL;
