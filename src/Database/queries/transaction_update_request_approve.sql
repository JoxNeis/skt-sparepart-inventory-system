-- =====================================================================
-- Query   : transaction_update_request_approve
-- Type    : UPDATE
-- Purpose : Mark a pending update request approved.
--           0 rows affected = not pending, or the approver's tier is
--           not higher than the requester's (same rule as
--           user_check_can_approve, enforced here as well).
--           The caller then reverses the original transaction and
--           activates the proposed one, in the same ADO transaction.
-- Replaces: update_transaction_request_update_approved
-- =====================================================================
PARAMETERS
	[p_update_request_id] Long,
	[p_approved_by] Text(64),
	[p_decision_note] LongText;

UPDATE
	transaction_update_requests AS ur
SET
	ur.approved_by = [p_approved_by],
	ur.approved_at = Now(),
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
				approver.username = [p_approved_by]
				AND requester.username = ur.requested_by
				AND approver.deleted_at IS NULL
				AND approver.tier > requester.tier
		);
