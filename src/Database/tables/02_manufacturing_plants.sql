-- Table: manufacturing_plants
CREATE TABLE [manufacturing_plants] (
    [id] AUTOINCREMENT,
    [name] TEXT(255),
    [code] TEXT(15),
    [address] TEXT(255),
    CONSTRAINT [pk_manufacturing_plants] PRIMARY KEY ([id])
);

-- Column properties:
--   name: AllowZeroLength=true
--   code: AllowZeroLength=true
--   address: AllowZeroLength=true

-- Indexes
CREATE INDEX [id] ON [manufacturing_plants] ([id]);
CREATE INDEX [code] ON [manufacturing_plants] ([code]);
