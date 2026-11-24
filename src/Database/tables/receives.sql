-- Table: receives
-- Source: Sparepart Inventory System Ver 5.accdb (MS Access / ACE SQL dialect)

CREATE TABLE [receives] (
    [id] AUTOINCREMENT,
    [transaction_id] LONG DEFAULT 0,
    [order_id] BIGINT DEFAULT 0,
    [arrival_date] DATETIME,
    [verification_date] DATETIME,
    [total_receive_amount] LONG DEFAULT 0,
    [ok_amount] LONG DEFAULT 0,
    [not_ok_amount] LONG DEFAULT 0,
    [remarks] LONGTEXT,
    [system_remarks] LONGTEXT,
    [person_in_charge] TEXT(255),
    CONSTRAINT [PrimaryKey] PRIMARY KEY ([id])
);

-- Column properties:
--   person_in_charge: AllowZeroLength=true

-- Indexes
CREATE INDEX [id] ON [receives] ([id]);
CREATE INDEX [order_id] ON [receives] ([order_id]);
CREATE UNIQUE INDEX [transaction_id] ON [receives] ([transaction_id]);

-- Foreign keys / relationships
ALTER TABLE [receives] ADD CONSTRAINT [ordersreceives] FOREIGN KEY ([order_id]) REFERENCES [orders] ([id]); -- NOTE: relationship has NO enforced referential integrity in the .accdb
ALTER TABLE [receives] ADD CONSTRAINT [transactionsreceives] FOREIGN KEY ([transaction_id]) REFERENCES [transactions] ([id]);
