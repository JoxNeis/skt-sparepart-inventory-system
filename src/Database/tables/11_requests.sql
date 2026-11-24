-- =====================================================================
-- Table  : requests
-- Purpose: Detail of a REQUEST transaction - spareparts issued from the
--          plant store to a mobile maintenance. Leaves stock on request_date.
-- =====================================================================
CREATE TABLE [requests] (
	[id]                     AUTOINCREMENT,
	[transaction_id]         LONG       NOT NULL,
	[mobile_maintenance_id]  LONG       NOT NULL,
	[sparepart_id]           LONG       NOT NULL,
	[request_date]           DATETIME   NOT NULL,
	[quantity]               LONG       NOT NULL,
	[remarks]                LONGTEXT,
	[system_remarks]         LONGTEXT,
	[person_in_charge]       TEXT(255),

	CONSTRAINT
		[pk_requests]
			PRIMARY KEY ([id]),
	CONSTRAINT
		[uq_requests_transaction]
			UNIQUE ([transaction_id]),
	CONSTRAINT
		[fk_requests_transaction]
			FOREIGN KEY ([transaction_id])
				REFERENCES [transactions] ([id]),
	CONSTRAINT
		[fk_requests_mobile_maintenance]
			FOREIGN KEY ([mobile_maintenance_id])
				REFERENCES [mobile_maintenances] ([id]),
	CONSTRAINT
		[fk_requests_sparepart]
			FOREIGN KEY ([sparepart_id])
				REFERENCES [spareparts] ([id]),
	CONSTRAINT
		[ck_requests_quantity]
			CHECK ([quantity] > 0)
);

CREATE INDEX
	[idx_requests_request_date]
		ON [requests] ([request_date]);
