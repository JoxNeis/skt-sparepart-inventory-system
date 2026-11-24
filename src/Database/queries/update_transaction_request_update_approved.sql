-- Query: update_transaction_request_update_approved
-- Type: UPDATE

PARAMETERS [p_original_transaction_uuid] Text(255), [p_proposed_transaction_uuid] Text(255), [p_user] Text(255);
UPDATE (transaction_update_requests AS ur INNER JOIN transactions AS original_t ON ur.original_transaction_id = original_t.id) INNER JOIN transactions AS proposed_t ON ur.proposed_transaction_id = proposed_t.id
SET ur.approved_by = [p_user], ur.approved_at = Now()
WHERE original_t.[uuid] = [p_original_transaction_uuid]
    AND proposed_t.[uuid] = [p_proposed_transaction_uuid];
