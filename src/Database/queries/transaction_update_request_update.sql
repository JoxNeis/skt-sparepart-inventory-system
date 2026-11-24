-- =====================================================================
-- Query   : transaction_update_request_update
-- Type    : UPDATE
-- Purpose : Record the verdict ('APPROVED' or 'REJECTED') on a pending
--           update request.
--           0 rows affected = not pending, an unknown verdict, or the
--           reviewer's tier is not higher than the requester's (same
--           rule as user_check_can_approve, enforced here as well).
--           In the same ADO transaction the caller then:
--             APPROVED -> reverses the original transaction and
--                         activates the proposed one
--             REJECTED -> reverses the proposed draft so it can never
--                         be activated
-- Replaces: transaction_update_request_approve,
--           transaction_update_request_reject
-- =====================================================================
PARAMETERS
	[p_update_request_id] Long,
	[p_verdict] Text(10),
	[p_reviewer] Text(64),
	[p_decision_note] LongText;

UPDATE
	transaction_update_requests AS ur
SET
	ur.verdict = [p_verdict],
	ur.reviewer = [p_reviewer],
	ur.reviewed_at = Now(),
	ur.decision_note = [p_decision_note]
WHERE
	ur.id = [p_update_request_id]
	AND ur.verdict IS NULL
	AND [p_verdict] IN ('APPROVED', 'REJECTED')
	AND EXISTS
		(
			SELECT
				1
			FROM
				users AS reviewer,
				users AS requester
			WHERE
				reviewer.username = [p_reviewer]
				AND requester.username = ur.requested_by
				AND reviewer.deleted_at IS NULL
				AND reviewer.tier > requester.tier
		);
