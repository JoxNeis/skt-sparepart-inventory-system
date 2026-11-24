-- Table: requests
-- Source: Sparepart Inventory System Ver 5.accdb (MS Access / ACE SQL dialect)

CREATE TABLE [requests] (
    [id] AUTOINCREMENT,
    [transaction_id] LONG DEFAULT 0,
    [mobile_post_id] BIGINT DEFAULT 0,
    [request_date] DATETIME,
    [catalogue_id] BIGINT DEFAULT 0,
    [amount] LONG DEFAULT 0,
    [remarks] LONGTEXT,
    [system_remarks] LONGTEXT,
    [person_in_charge] TEXT(255),
    CONSTRAINT [PrimaryKey] PRIMARY KEY ([id])
);

-- Column properties:
--   person_in_charge: AllowZeroLength=true

-- Indexes
CREATE INDEX [catalogue_id] ON [requests] ([catalogue_id]);
CREATE INDEX [id] ON [requests] ([id]);
CREATE INDEX [mobile_post_id] ON [requests] ([mobile_post_id]);
CREATE UNIQUE INDEX [transaction_id] ON [requests] ([transaction_id]);

-- Foreign keys / relationships
ALTER TABLE [requests] ADD CONSTRAINT [cataloguesrequests] FOREIGN KEY ([catalogue_id]) REFERENCES [catalogues] ([id]); -- NOTE: relationship has NO enforced referential integrity in the .accdb
ALTER TABLE [requests] ADD CONSTRAINT [mobile_postsrequests] FOREIGN KEY ([mobile_post_id]) REFERENCES [mobile_posts] ([ID]); -- NOTE: relationship has NO enforced referential integrity in the .accdb
ALTER TABLE [requests] ADD CONSTRAINT [transactionsrequests] FOREIGN KEY ([transaction_id]) REFERENCES [transactions] ([id]);
