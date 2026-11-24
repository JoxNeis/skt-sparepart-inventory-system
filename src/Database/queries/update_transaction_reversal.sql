-- Query: update_transaction_reversal
-- Type: UPDATE

PARAMETERS [p_transaction_uuid] Text(255);
UPDATE transactions
SET reversed_at = Now ()
WHERE uuid = [p_transaction_uuid];
