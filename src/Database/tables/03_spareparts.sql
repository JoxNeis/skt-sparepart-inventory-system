-- Table: spareparts
CREATE TABLE [spareparts] (
    [id] AUTOINCREMENT,
    [name] TEXT(255),
    [code] TEXT(15),
    [category] TEXT(31),
    [applicable] TEXT(255),
    [unit_of_measurement] TEXT(15),
    [alias] TEXT(255),
    [display] AS ([code] & " - " & [name]),
    [created_at] DATETIME,
    [updated_at] DATETIME,
    [deleted_at] DATETIME,
    [last_synced_at] DATETIME,
    CONSTRAINT [pk_spareparts] PRIMARY KEY ([id])
);


-- Column properties:
--   name: AllowZeroLength=true
--   alias_name: AllowZeroLength=true
--   code: AllowZeroLength=true
--   category: AllowZeroLength=true
--   unit_of_measurement: AllowZeroLength=true

-- Indexes
CREATE INDEX [id] ON [spareparts] ([id]);
CREATE INDEX [code] ON [spareparts] ([code]);
CREATE INDEX [display] ON [spareparts] ([display]);