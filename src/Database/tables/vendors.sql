-- Table: vendors
-- Source: Sparepart Inventory System Ver 5.accdb (MS Access / ACE SQL dialect)

CREATE TABLE [vendors] (
    [id] AUTOINCREMENT,
    [name] TEXT(255),
    [company_name] TEXT(255),
    [code] TEXT(255),
    [email] TEXT(255),
    CONSTRAINT [PrimaryKey] PRIMARY KEY ([id])
);

-- Column properties:
--   name: AllowZeroLength=true
--   company_name: AllowZeroLength=true
--   code: AllowZeroLength=true
--   email: AllowZeroLength=true

-- Indexes
CREATE INDEX [code] ON [vendors] ([code]);
