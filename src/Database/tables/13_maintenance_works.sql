-- =====================================================================
-- Table  : maintenance_works
-- Purpose: Header of a mobile maintenance work (one job done on a date).
--          Its tasks and the spareparts used live in
--          maintenance_work_details.
-- Notes  : [uuid] is the work document number (same TEXT(36) format as
--          transactions.uuid; it is not required to be an RFC 4122 UUID).
--          [maintenance_type] is free text for now, e.g. PREVENTIVE or
--          CORRECTIVE.
-- =====================================================================
CREATE TABLE [maintenance_works] (
	[id]                AUTOINCREMENT,
	[uuid]              TEXT(36)   NOT NULL,
	[maintenance_name]  TEXT(255)  NOT NULL,
	[maintenance_date]  DATETIME   NOT NULL,
	[maintenance_type]  TEXT(50)   NOT NULL,

	CONSTRAINT
		[pk_maintenance_works]
			PRIMARY KEY ([id]),
	CONSTRAINT
		[uq_maintenance_works_uuid]
			UNIQUE ([uuid])
);

CREATE INDEX
	[idx_maintenance_works_date]
		ON [maintenance_works] ([maintenance_date]);
CREATE INDEX
	[idx_maintenance_works_type]
		ON [maintenance_works] ([maintenance_type]);
