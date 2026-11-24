-- =====================================================================
-- Table  : orders
-- Purpose: Detail of an ORDER transaction - spareparts ordered from a
--          vendor. The plant comes from the transaction header.
-- Notes  : [po_number] was po_id. Renamed because *_id is reserved for
--          foreign keys, and this is a free-text SAP PO number.
-- =====================================================================
CREATE TABLE [orders] (
	[id]                AUTOINCREMENT,
	[transaction_id]    LONG       NOT NULL,
	[sparepart_id]      LONG       NOT NULL,
	[vendor_id]         LONG,
	[po_number]         TEXT(50),
	[order_date]        DATETIME   NOT NULL,
	[quantity]          LONG       NOT NULL,
	[remarks]           LONGTEXT,
	[system_remarks]    LONGTEXT,
	[person_in_charge]  TEXT(255),

	CONSTRAINT
		[pk_orders]
			PRIMARY KEY ([id]),
	CONSTRAINT
		[uq_orders_transaction]
			UNIQUE ([transaction_id]),
	CONSTRAINT
		[fk_orders_transaction]
			FOREIGN KEY ([transaction_id])
				REFERENCES [transactions] ([id]),
	CONSTRAINT
		[fk_orders_sparepart]
			FOREIGN KEY ([sparepart_id])
				REFERENCES [spareparts] ([id]),
	CONSTRAINT
		[fk_orders_vendor]
			FOREIGN KEY ([vendor_id])
				REFERENCES [vendors] ([id]),
	CONSTRAINT
		[ck_orders_quantity]
			CHECK ([quantity] > 0)
);

CREATE INDEX
	[idx_orders_order_date]
		ON [orders] ([order_date]);
CREATE INDEX
	[idx_orders_po_number]
		ON [orders] ([po_number]);
