-- Table: users
CREATE TABLE [users] (
    [id] AUTOINCREMENT,
    [username] TEXT(255),
    [first_name] TEXT(255),
    [middle_name] TEXT(255),
    [last_name] TEXT(255),
    [position] TEXT(255),
    [tier] LONG DEFAULT 0,
    [manufacturing_plant_id] LONG DEFAULT 0,
    CONSTRAINT [pk_users] PRIMARY KEY ([id])
);

-- Column properties:
--   username: AllowZeroLength=true
--   first_name: AllowZeroLength=true
--   middle_name: AllowZeroLength=true
--   last_name: AllowZeroLength=true
--   position: AllowZeroLength=true

-- Indexes
CREATE INDEX [id] ON [users] ([id]);
CREATE INDEX [manufacturing_plant_id] ON [users] ([manufacturing_plant_id]);
CREATE UNIQUE INDEX [username] ON [users] ([username]);

-- Foreign keys / relationships
ALTER TABLE [users] ADD CONSTRAINT [manufacturing_plantsusers] FOREIGN KEY ([manufacturing_plant_id]) REFERENCES [manufacturing_plants] ([id]);
