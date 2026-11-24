-- Table: orders
CREATE TABLE [orders] (
    [id] AUTOINCREMENT,
    [transaction_id] LONG NOT NULL,
    [po_id] TEXT(255),
    [order_date] DATETIME,
    [catalogue_id] LONG NOT NULL,
    [amount] LONG NOT NULL DEFAULT 0,
    [manufacturing_plant_id] LONG NOT NULL,
    [vendor_id] LONG,
    [remarks] LONGTEXT,
    [system_remarks] LONGTEXT,
    [person_in_charge] TEXT(255),

    CONSTRAINT [pk_orders] PRIMARY KEY ([id]),
    CONSTRAINT [fk_orders_transaction] FOREIGN KEY ([transaction_id]) REFERENCES [transactions] ([id]),
    CONSTRAINT [fk_orders_catalogue] FOREIGN KEY ([catalogue_id]) REFERENCES [catalogues] ([id]),
    CONSTRAINT [fk_orders_manufacturing_plant] FOREIGN KEY ([manufacturing_plant_id]) REFERENCES [manufacturing_plants] ([id]),
    CONSTRAINT [fk_orders_vendor] FOREIGN KEY ([vendor_id]) REFERENCES [vendors] ([id])
);

CREATE UNIQUE INDEX [uq_orders_transaction] ON [orders] ([transaction_id]);
CREATE INDEX [idx_orders_manufacturing_plant] ON [orders] ([manufacturing_plant_id]);
CREATE INDEX [idx_orders_catalogue] ON [orders] ([catalogue_id]);
CREATE INDEX [idx_orders_vendor] ON [orders] ([vendor_id]);