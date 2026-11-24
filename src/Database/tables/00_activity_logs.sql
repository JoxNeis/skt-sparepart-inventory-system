-- Table: activity_logs
CREATE TABLE [activity_logs] (
    [id] AUTOINCREMENT,
    [activity_code] TEXT(255),
    [detail] LONGTEXT,
    [user] TEXT(255),
    [created_at] DATETIME,
    CONSTRAINT [pk_activity_logs] PRIMARY KEY ([id])
);

-- Column properties:
--   activity_code: AllowZeroLength=true
--   user: AllowZeroLength=true

-- Indexes
CREATE INDEX [id] ON [activity_logs] ([id]);
CREATE INDEX [activity_code] ON [activity_logs] ([activity_code]);
