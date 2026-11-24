-- Table: transactions
CREATE TABLE [transactions] (
    [id] AUTOINCREMENT,
    [uuid] TEXT(36) NOT NULL,
    [transaction_type] TEXT(50) NOT NULL,
    [signed_at] DATETIME,
    [digital_signature] LONGTEXT,
    [created_by] TEXT(255) NOT NULL,
    [activated_at] DATETIME,
    [created_at] DATETIME,
    [updated_at] DATETIME,
    [reversed_at] DATETIME,
    [synced_at] DATETIME,
    [synced_by] TEXT(255),

    CONSTRAINT [pk_transactions] PRIMARY KEY ([id]),
    CONSTRAINT [fk_transactions_created_by] FOREIGN KEY ([created_by]) REFERENCES [users] ([username]),
    CONSTRAINT [fk_transactions_synced_by] FOREIGN KEY ([synced_by]) REFERENCES [users] ([username])
);

CREATE UNIQUE INDEX [uq_transactions_uuid] ON [transactions] ([uuid]);
CREATE INDEX [idx_transactions_type] ON [transactions] ([transaction_type]);
CREATE INDEX [idx_transactions_created_by] ON [transactions] ([created_by]);
CREATE INDEX [idx_transactions_synced_by] ON [transactions] ([synced_by]);
CREATE INDEX [idx_transactions_created_at] ON [transactions] ([created_at]);
CREATE INDEX [idx_transactions_synced_at] ON [transactions] ([synced_at]);