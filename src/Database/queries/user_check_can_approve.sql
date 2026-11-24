-- =====================================================================
-- Query   : user_check_can_approve
-- Type    : SELECT
-- Purpose : True when the approver's tier is higher than the
--           requester's. No row when either user is unknown or the
--           approver is archived.
-- Replaces: compare_user
-- =====================================================================
PARAMETERS
    [p_requested_by] Text(64),
    [p_approved_by] Text(64);

SELECT
    IIf(approver.tier > requester.tier, True, False) AS can_approve
FROM
    users AS requester,
    users AS approver
WHERE
    requester.username = [p_requested_by]
    AND approver.username = [p_approved_by]
    AND approver.deleted_at IS NULL;
