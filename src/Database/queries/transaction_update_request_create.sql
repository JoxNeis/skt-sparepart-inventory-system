-- =====================================================================
-- Query   : transaction_update_request_create
-- Type    : APPEND
-- Purpose : Ask for an active transaction (p_original_uuid) to be
--           replaced by a draft one (p_proposed_uuid).
--           Inserts nothing (0 rows) unless: both exist, same type,
--           same plant, the original is active, the proposal is still
--           a draft, and the original has no other pending request.
-- Replaces: create_transaction_update_request
-- =====================================================================
PARAMETERS
	[p_original_uuid] Text(36),
	[p_proposed_uuid] Text(36),
	[p_reason] LongText,
	[p_requested_by] Text(64);

INSERT INTO
	transaction_update_requests
		(
			original_transaction_id,
			proposed_transaction_id,
			reason,
			requested_by,
			requested_at
		)
	SELECT
		old_t.id AS original_transaction_id,
		new_t.id AS proposed_transaction_id,
		[p_reason] AS reason,
		[p_requested_by] AS requested_by,
		Now() AS requested_at
	FROM
		transactions AS old_t,
		transactions AS new_t
	WHERE
		old_t.uuid = [p_original_uuid]
		AND new_t.uuid = [p_proposed_uuid]
		AND new_t.transaction_type = old_t.transaction_type
		AND new_t.manufacturing_plant_id = old_t.manufacturing_plant_id
		AND old_t.activated_at IS NOT NULL
		AND old_t.reversed_at IS NULL
		AND new_t.activated_at IS NULL
		AND NOT EXISTS
			(
				SELECT
					1
				FROM
					transaction_update_requests AS ur
				WHERE
					ur.original_transaction_id = old_t.id
					AND ur.verdict IS NULL
			);
