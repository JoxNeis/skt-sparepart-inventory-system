-- =====================================================================
-- Table  : maintenance_works
-- Purpose: Detail of a MAINTENANCE transaction - one mobile maintenance
--          job done on a date. Its tasks and the spareparts used live in
--          maintenance_work_details. The document number (MTW-...) and
--          the plant come from the transaction header.
-- Notes  : [maintenance_type] is free text for now, e.g. PREVENTIVE or
--          CORRECTIVE.
--          [uuid] was removed: the document number is transactions.uuid.
--          [transaction_id] is one-to-one with transactions. Access
--          DDL can only make one-to-many relations, so
--          fk_maintenance_works_transaction is added through DAO after this
--          file runs (see $OneToOneRelations in
--          migration\generate_migration_vba.ps1).
-- =====================================================================
CREATE TABLE [maintenance_works] (
	[id]                AUTOINCREMENT,
	[transaction_id]    LONG       NOT NULL,
	[maintenance_name]  TEXT(255)  NOT NULL,
	[maintenance_date]  DATETIME   NOT NULL,
	[maintenance_type]  TEXT(50)   NOT NULL,

	CONSTRAINT
		[pk_maintenance_works]
			PRIMARY KEY ([id]),
	CONSTRAINT
		[uq_maintenance_works_transaction]
			UNIQUE ([transaction_id])
);

CREATE INDEX
	[idx_maintenance_works_date]
		ON [maintenance_works] ([maintenance_date]);
CREATE INDEX
	[idx_maintenance_works_type]
		ON [maintenance_works] ([maintenance_type]);
