-- =====================================================================
-- Table  : maintenance_work_details
-- Purpose: Lines of a maintenance work - one row per task performed,
--          optionally with the sparepart and amount used for it.
-- Notes  : [sparepart_id] and [amount] are NULL for a task that used no
--          sparepart.
--          [maintenance_task_id] has no FK yet: there is no
--          maintenance_tasks table. Add the constraint once it exists.
-- =====================================================================
CREATE TABLE [maintenance_work_details] (
	[id]                   AUTOINCREMENT,
	[maintenance_work_id]  LONG  NOT NULL,
	[maintenance_task_id]  LONG  NOT NULL,
	[sparepart_id]         LONG,
	[amount]               LONG,

	CONSTRAINT
		[pk_maintenance_work_details]
			PRIMARY KEY ([id]),
	CONSTRAINT
		[fk_mw_details_maintenance_work]
			FOREIGN KEY ([maintenance_work_id])
				REFERENCES [maintenance_works] ([id]),
	CONSTRAINT
		[fk_mw_details_sparepart]
			FOREIGN KEY ([sparepart_id])
				REFERENCES [spareparts] ([id])
);

CREATE INDEX
	[idx_mw_details_work]
		ON [maintenance_work_details] ([maintenance_work_id]);
CREATE INDEX
	[idx_mw_details_task]
		ON [maintenance_work_details] ([maintenance_task_id]);
CREATE INDEX
	[idx_mw_details_sparepart]
		ON [maintenance_work_details] ([sparepart_id]);
