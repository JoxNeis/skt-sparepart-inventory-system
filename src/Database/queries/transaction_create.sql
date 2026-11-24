-- =====================================================================
-- Query   : transaction_create
-- Type    : APPEND
-- Purpose : Create a document header. Pass p_activated_at = Null for a
--           draft (e.g. the proposed side of an update request), or
--           Now() for a document that is active immediately.
-- Replaces: create_transaction
-- =====================================================================
PARAMETERS
    [p_uuid] Text(36),
    [p_transaction_type] Text(15),
    [p_manufacturing_plant_id] Long,
    [p_created_by] Text(64),
    [p_activated_at] DateTime;

INSERT INTO
    transactions
        (
            uuid,
            transaction_type,
            manufacturing_plant_id,
            created_by,
            created_at,
            activated_at
        )
    VALUES
        (
            [p_uuid],
            [p_transaction_type],
            [p_manufacturing_plant_id],
            [p_created_by],
            Now(),
            [p_activated_at]
        );
