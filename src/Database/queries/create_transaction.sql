-- Query: create_transaction
-- Type: APPEND

PARAMETERS [p_uuid] Text(255), [p_transaction_type] Text(255), [p_created_by] Text(255), [p_active_at] DateTime;
INSERT INTO transactions (uuid, transaction_type, created_by, actived_at, created_at, updated_at)
VALUES ([p_uuid], [p_transaction_type], [p_created_by], [p_active_at], Now(), Now());
