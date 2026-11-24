-- Query: create_activity_log
-- Type: APPEND

PARAMETERS [p_activity_code] Text(255), [p_detail] LongText, [p_user] Text(255);
INSERT INTO activity_logs (activity_code, detail, user, created_at)
VALUES ([p_activity_code], [p_detail], [p_user], Now ());
