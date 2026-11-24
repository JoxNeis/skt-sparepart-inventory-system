-- Table: transaction_update_requests
-- Source: Sparepart Inventory System Ver 5.accdb (MS Access / ACE SQL dialect)

CREATE TABLE [transaction_update_requests] (
    [id] AUTOINCREMENT,
    [original_transaction_id] LONG DEFAULT 0,
    [proposed_transaction_id] LONG DEFAULT 0,
    [requested_by] TEXT(255),
    [requested_at] DATETIME,
    [approved_by] TEXT(255),
    [approved_at] DATETIME,
    [rejected_at] DATETIME,
    CONSTRAINT [PrimaryKey] PRIMARY KEY ([id])
);

-- Column properties:
--   requested_by: AllowZeroLength=true
--   approved_by: AllowZeroLength=true

-- Indexes
CREATE INDEX [id] ON [transaction_update_requests] ([id]);
CREATE UNIQUE INDEX [original_transaction_id] ON [transaction_update_requests] ([original_transaction_id]);
CREATE UNIQUE INDEX [proposed_transaction_id] ON [transaction_update_requests] ([proposed_transaction_id]);

-- Foreign keys / relationships
ALTER TABLE [transaction_update_requests] ADD CONSTRAINT [transactionstransaction_update_requests] FOREIGN KEY ([original_transaction_id], [proposed_transaction_id]) REFERENCES [transactions] ([id], [id]); -- NOTE: relationship has NO enforced referential integrity in the .accdb
