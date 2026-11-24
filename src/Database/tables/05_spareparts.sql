-- =====================================================================
-- Table  : spareparts
-- Purpose: Master data - the central sparepart catalogue.
-- Notes  : [display] ("code - name") is a calculated field. Access DDL
--          cannot create one, so build.ps1 adds it through DAO after
--          this file runs (see $CalculatedFields there). It cannot be
--          indexed and is never written by a query.
--          [alias] is the one other name a sparepart is known by.
-- =====================================================================
CREATE TABLE [spareparts] (
	[id]                   AUTOINCREMENT,
	[code]                 TEXT(40)   NOT NULL,
	[name]                 TEXT(255)  NOT NULL,
	[category]             TEXT(50),
	[applicable]           TEXT(255),
	[unit_of_measurement]  TEXT(15)   NOT NULL,
	[alias]                TEXT(255)  NOT NULL,
	[status]               TEXT(50),
	[created_at]           DATETIME   NOT NULL DEFAULT Now(),
	[updated_at]           DATETIME,
	[deleted_at]           DATETIME,
	[last_synced_at]       DATETIME,

	CONSTRAINT
		[pk_spareparts]
			PRIMARY KEY ([id]),
	CONSTRAINT
		[uq_spareparts_code]
			UNIQUE ([code])
);

CREATE INDEX
	[idx_spareparts_name]
		ON [spareparts] ([name]);
CREATE INDEX
	[idx_spareparts_category]
		ON [spareparts] ([category]);
