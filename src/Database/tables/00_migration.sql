-- =====================================================================
-- Table  : migrations
-- Purpose: Tracks which schema versions have been applied to this
--          database, so upgrade scripts run once and in order.
-- Notes  : [version] is the migration number (e.g. 1, 2, 3...); the
--          current DB version is MAX([version]).
--          Rows are append-only; never edit or delete an applied row.
-- =====================================================================
CREATE TABLE [migrations] (
    [id]          AUTOINCREMENT,
    [version]     LONG          NOT NULL,
    [name]        TEXT(100)     NOT NULL,
    [applied_by]  TEXT(64),
    [applied_at]  DATETIME      NOT NULL DEFAULT Now(),
    CONSTRAINT
        [pk_migrations]
            PRIMARY KEY ([id])
);

CREATE UNIQUE INDEX
    [uq_migrations_version]
        ON [migrations] ([version]);
