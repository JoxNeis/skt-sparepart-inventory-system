-- =====================================================================
-- Query   : update_request_reject
-- Type    : UPDATE
-- Purpose : Mark a pending update request rejected (same tier rule as
--           approving). The caller should reverse the proposed draft
--           transaction so it can never be activated.
-- =====================================================================
PARAMETERS
    [p_update_request_id] Long,
    [p_rejected_by] Text(64),
    [p_decision_note] LongText;

UPDATE
    transaction_update_requests AS ur
SET
    ur.rejected_by = [p_rejected_by],
    ur.rejected_at = Now(),
    ur.decision_note = [p_decision_note]
WHERE
    ur.id = [p_update_request_id]
    AND ur.approved_at IS NULL
    AND ur.rejected_at IS NULL
    AND EXISTS
        (
            SELECT
                1
            FROM
                users AS approver,
                users AS requester
            WHERE
                approver.username = [p_rejected_by]
                AND requester.username = ur.requested_by
                AND approver.deleted_at IS NULL
                AND approver.tier > requester.tier
        );
