-- =====================================================================
-- Table  : transactions
-- Purpose: Header shared by every document (ORDER, RECEIVE, REQUEST).
--          Each header has exactly one detail row in orders, receives
--          or requests (1:1 through the detail's unique transaction_id).
-- Life   : draft    -> activated_at IS NULL
--          active   -> activated_at set, reversed_at IS NULL
--          reversed -> reversed_at set
--          A transaction is never edited or deleted. An edit is a new
--          proposed transaction (see transaction_update_requests).
-- Notes  : [uuid] is the document number, e.g. ORD-20261124-001, made
--          by transaction_get_next_uuid (it is not an RFC 4122 UUID).
--          [manufacturing_plant_id] sits on the header so that every
--          document type is filtered by plant in the same way.
-- =====================================================================
CREATE TABLE [transactions] (
	[id]                      AUTOINCREMENT,
	[uuid]                    TEXT(36)  NOT NULL,
	[transaction_type]        TEXT(15)  NOT NULL,
	[manufacturing_plant_id]  LONG      NOT NULL,
	[created_by]              TEXT(64)  NOT NULL,
	[created_at]              DATETIME  NOT NULL DEFAULT Now(),
	[activated_at]            DATETIME,
	[reversed_by]             TEXT(64),
	[reversed_at]             DATETIME,
	[signed_at]               DATETIME,
	[digital_signature]       LONGTEXT,
	[synced_by]               TEXT(64),
	[synced_at]               DATETIME,

	CONSTRAINT
		[pk_transactions]
			PRIMARY KEY ([id]),
	CONSTRAINT
		[uq_transactions_uuid]
			UNIQUE ([uuid]),
	CONSTRAINT
		[fk_transactions_manufacturing_plant]
			FOREIGN KEY ([manufacturing_plant_id])
				REFERENCES [manufacturing_plants] ([id]),
	CONSTRAINT
		[fk_transactions_created_by]
			FOREIGN KEY ([created_by])
				REFERENCES [users] ([username]),
	CONSTRAINT
		[fk_transactions_reversed_by]
			FOREIGN KEY ([reversed_by])
				REFERENCES [users] ([username]),
	CONSTRAINT
		[fk_transactions_synced_by]
			FOREIGN KEY ([synced_by])
				REFERENCES [users] ([username]),
	CONSTRAINT
		[ck_transactions_type]
			CHECK (
				[transaction_type] IN ('ORDER', 'RECEIVE', 'REQUEST')
			)
);

CREATE INDEX
	[idx_transactions_type]
		ON [transactions] ([transaction_type]);
CREATE INDEX
	[idx_transactions_activated_at]
		ON [transactions] ([activated_at]);
CREATE INDEX
	[idx_transactions_reversed_at]
		ON [transactions] ([reversed_at]);
