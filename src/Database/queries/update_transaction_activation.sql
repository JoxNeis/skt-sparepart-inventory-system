-- Query: update_transaction_activation
-- Type: UPDATE

PARAMETERS [p_original_transaction_uuid] Text(255);
UPDATE transactions
SET actived_at = Now()
WHERE uuid = [p_original_transaction_uuid];
