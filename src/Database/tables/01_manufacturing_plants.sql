-- =====================================================================
-- Table  : manufacturing_plants
-- Purpose: Master data - one row per plant. Every stock row, user,
--          mobile maintenance and transaction belongs to exactly one plant.
-- Notes  : [type] and [region] are free text, e.g. type "PLANT" /
--          "WAREHOUSE", region "JAWA TIMUR".
-- =====================================================================
CREATE TABLE [manufacturing_plants] (
	[id]          AUTOINCREMENT,
	[code]        TEXT(15)   NOT NULL,
	[name]        TEXT(255)  NOT NULL,
	[address]     TEXT(255),
	[type]        TEXT(50),
	[region]      TEXT(100),
	[created_at]  DATETIME   NOT NULL DEFAULT Now(),
	[updated_at]  DATETIME,
	[deleted_at]  DATETIME,

	CONSTRAINT
		[pk_manufacturing_plants]
			PRIMARY KEY ([id]),
	CONSTRAINT
		[uq_manufacturing_plants_code]
			UNIQUE ([code]),
	CONSTRAINT
		[uq_manufacturing_plants_name]
			UNIQUE ([name])
);
