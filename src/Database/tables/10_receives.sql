-- =====================================================================
-- Table  : receives
-- Purpose: Detail of a RECEIVE transaction - goods arriving against an
--          order, split into OK and NOT OK after verification.
--          Only ok_quantity enters stock, on verification_date.
-- Notes  : ck_receives_quantity is the same rule serviceReceive
--          already checks (total = OK + NOT OK).
--          received_quantity 0 = the receive was deleted (through an
--          approved update), the same as quantity 0 on orders/requests.
--          [transaction_id] is one-to-one with transactions. Access
--          DDL can only make one-to-many relations, so
--          fk_receives_transaction is added through DAO after this
--          file runs (see $OneToOneRelations in
--          migration\generate_migration_vba.ps1).
-- =====================================================================
CREATE TABLE [receives] (
	[id]                 AUTOINCREMENT,
	[transaction_id]     LONG       NOT NULL,
	[order_id]           LONG       NOT NULL,
	[arrival_date]       DATETIME   NOT NULL,
	[verification_date]  DATETIME,
	[received_quantity]  LONG       NOT NULL DEFAULT 0,
	[ok_quantity]        LONG       NOT NULL DEFAULT 0,
	[not_ok_quantity]    LONG       NOT NULL DEFAULT 0,
	[remarks]            LONGTEXT,
	[system_remarks]     LONGTEXT,
	[person_in_charge]   TEXT(255),

	CONSTRAINT
		[pk_receives]
			PRIMARY KEY ([id]),
	CONSTRAINT
		[uq_receives_transaction]
			UNIQUE ([transaction_id]),
	CONSTRAINT
		[fk_receives_order]
			FOREIGN KEY ([order_id])
				REFERENCES [orders] ([id]),
	CONSTRAINT
		[ck_receives_quantity]
			CHECK (
				[ok_quantity] >= 0
				AND [not_ok_quantity] >= 0
				AND [received_quantity] = [ok_quantity] + [not_ok_quantity]
			)
);

CREATE INDEX
	[idx_receives_arrival_date]
		ON [receives] ([arrival_date]);
CREATE INDEX
	[idx_receives_verification_date]
		ON [receives] ([verification_date]);
