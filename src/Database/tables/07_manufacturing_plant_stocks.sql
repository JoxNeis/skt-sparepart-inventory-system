-- =====================================================================
-- Table  : manufacturing_plant_stocks
-- Purpose: One row per (plant, sparepart): the on-hand stock and the
--          reorder point (ROP) planning inputs and outputs.
-- Notes  : current_stock is a cache of the stock ledger
--          (vw_stock_ledger), kept in step by the application.
--            safety_stock  = service_level_z * daily_demand_std_dev
--                            * Sqr(lead_time_days)
--            reorder_point = safety_stock
--                            + average_daily_usage * lead_time_days
--          service_level_z 1.65 = 95 % service level (was hard-coded
--          in update_rop).
-- =====================================================================
CREATE TABLE [manufacturing_plant_stocks] (
	[id]                      AUTOINCREMENT,
	[manufacturing_plant_id]  LONG      NOT NULL,
	[sparepart_id]            LONG      NOT NULL,
	[current_stock]           LONG      NOT NULL DEFAULT 0,
	[reorder_point]           LONG      NOT NULL DEFAULT 0,
	[safety_stock]            LONG      NOT NULL DEFAULT 0,
	[lead_time_days]          LONG      NOT NULL DEFAULT 0,
	[average_daily_usage]     DOUBLE    NOT NULL DEFAULT 0,
	[daily_demand_std_dev]    DOUBLE    NOT NULL DEFAULT 0,
	[service_level_z]         DOUBLE    NOT NULL DEFAULT 1.65,
	[created_at]              DATETIME  NOT NULL DEFAULT Now(),
	[updated_at]              DATETIME,

	CONSTRAINT
		[pk_manufacturing_plant_stocks]
			PRIMARY KEY ([id]),
	CONSTRAINT
		[uq_manufacturing_plant_stocks_plant_sparepart]
			UNIQUE ([manufacturing_plant_id], [sparepart_id]),
	CONSTRAINT
		[fk_manufacturing_plant_stocks_manufacturing_plant]
			FOREIGN KEY ([manufacturing_plant_id])
				REFERENCES [manufacturing_plants] ([id]),
	CONSTRAINT
		[fk_manufacturing_plant_stocks_sparepart]
			FOREIGN KEY ([sparepart_id])
				REFERENCES [spareparts] ([id]),
	CONSTRAINT
		[ck_manufacturing_plant_stocks_current_stock]
			CHECK (
				[current_stock] >= 0
			),
	CONSTRAINT
		[ck_manufacturing_plant_stocks_planning]
			CHECK (
				[reorder_point] >= 0
				AND [safety_stock] >= 0
				AND [lead_time_days] >= 0
				AND [average_daily_usage] >= 0
				AND [daily_demand_std_dev] >= 0
				AND [service_level_z] >= 0
			)
);
