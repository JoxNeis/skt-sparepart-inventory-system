-- =====================================================================
-- Table  : transaction_update_requests
-- Purpose: Approval workflow for editing an active transaction. The
--          edit is stored as a new draft (proposed) transaction.
--          Approving reverses the original and activates the proposal.
-- Life   : pending  -> approved_at IS NULL AND rejected_at IS NULL
--          approved -> approved_at set
--          rejected -> rejected_at set
-- Notes  : original_transaction_id is no longer unique, so a
--          transaction can be requested again after a rejection.
--          The old composite FK (original, proposed) -> (id, id) is
--          replaced by two separate foreign keys.
-- =====================================================================
CREATE TABLE [transaction_update_requests] (
	[id]                       AUTOINCREMENT,
	[original_transaction_id]  LONG      NOT NULL,
	[proposed_transaction_id]  LONG      NOT NULL,
	[reason]                   LONGTEXT,
	[requested_by]             TEXT(64)  NOT NULL,
	[requested_at]             DATETIME  NOT NULL DEFAULT Now(),
	[approved_by]              TEXT(64),
	[approved_at]              DATETIME,
	[rejected_by]              TEXT(64),
	[rejected_at]              DATETIME,
	[decision_note]            LONGTEXT,

	CONSTRAINT
		[pk_transaction_update_requests]
			PRIMARY KEY ([id]),
	CONSTRAINT
		[uq_transaction_update_requests_proposed_transaction]
			UNIQUE ([proposed_transaction_id]),
	CONSTRAINT
		[fk_transaction_update_requests_original_transaction]
			FOREIGN KEY ([original_transaction_id])
				REFERENCES [transactions] ([id]),
	CONSTRAINT
		[fk_transaction_update_requests_proposed_transaction]
			FOREIGN KEY ([proposed_transaction_id])
				REFERENCES [transactions] ([id]),
	CONSTRAINT
		[fk_transaction_update_requests_requested_by]
			FOREIGN KEY ([requested_by])
				REFERENCES [users] ([username]),
	CONSTRAINT
		[fk_transaction_update_requests_approved_by]
			FOREIGN KEY ([approved_by])
				REFERENCES [users] ([username]),
	CONSTRAINT
		[fk_transaction_update_requests_rejected_by]
			FOREIGN KEY ([rejected_by])
				REFERENCES [users] ([username]),
	CONSTRAINT
		[ck_transaction_update_requests_decision]
			CHECK (
				[original_transaction_id] <> [proposed_transaction_id]
				AND (
					[approved_at] IS NULL
					OR [rejected_at] IS NULL
				)
			)
);

CREATE INDEX
	[idx_transaction_update_requests_requested_at]
		ON [transaction_update_requests] ([requested_at]);
