-- =====================================================================
-- Table  : mobile_maintenances
-- Purpose: Master data - mobile maintenances inside a plant; the destination
--          of every request.
-- Notes  : code and name are unique per plant, not globally, so two
--          plants may both have a mobile maintenance "PM-01".
-- =====================================================================
CREATE TABLE [mobile_maintenances] (
	[id]                      AUTOINCREMENT,
	[manufacturing_plant_id]  LONG       NOT NULL,
	[code]                    TEXT(40)   NOT NULL,
	[name]                    TEXT(255)  NOT NULL,
	[created_at]              DATETIME   NOT NULL DEFAULT Now(),
	[updated_at]              DATETIME,
	[deleted_at]              DATETIME,

	CONSTRAINT
		[pk_mobile_maintenances]
			PRIMARY KEY ([id]),
	CONSTRAINT
		[uq_mobile_maintenances_plant_code]
			UNIQUE ([manufacturing_plant_id], [code]),
	CONSTRAINT
		[uq_mobile_maintenances_plant_name]
			UNIQUE ([manufacturing_plant_id], [name]),
	CONSTRAINT
		[fk_mobile_maintenances_manufacturing_plant]
			FOREIGN KEY ([manufacturing_plant_id])
				REFERENCES [manufacturing_plants] ([id])
);
