-- Query: create_transaction_update_request
-- Type: APPEND

PARAMETERS [p_original_transaction_uuid] Text(255), [p_proposed_transaction_uuid] Text(255), [p_requested_by] Text(255);
INSERT INTO transaction_update_requests (original_transaction_id, proposed_transaction_id, requested_by, requested_at)
SELECT TOP 1 ot.id, pt.id, [p_requested_by], Now()
FROM transactions AS ot, transactions AS pt
WHERE ot.uuid = [p_original_transaction_uuid]
    AND pt.uuid = [p_proposed_transaction_uuid];
