-- =====================================================================
-- Query   : transaction_activate
-- Type    : UPDATE
-- Purpose : Activate a draft transaction.
--           0 rows affected = not found, already active or reversed.
-- Replaces: update_transaction_activation
-- =====================================================================
PARAMETERS
	[p_uuid] Text(36);

UPDATE
	transactions
SET
	activated_at = Now()
WHERE
	uuid = [p_uuid]
	AND activated_at IS NULL
	AND reversed_at IS NULL;
