-- =====================================================================
-- Query   : activity_log_create
-- Type    : APPEND
-- Purpose : Write one audit-trail row. p_json is the optional
--           machine-readable payload (e.g. old/new values) next to the
--           human-readable p_detail.
-- Replaces: create_activity_log
-- =====================================================================
PARAMETERS
    [p_activity_code] Text(50),
    [p_detail] LongText,
    [p_json] LongText,
    [p_username] Text(64);

INSERT INTO
    activity_logs
        (
            activity_code,
            detail,
            [json],
            username,
            created_at
        )
    VALUES
        (
            [p_activity_code],
            [p_detail],
            [p_json],
            [p_username],
            Now()
        );
