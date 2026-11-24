-- =====================================================================
-- Table  : activity_logs
-- Purpose: Append-only audit trail written by the application.
-- Notes  : [username] (was [user], a reserved word) is deliberately
--          NOT a foreign key: a log row must outlive the user row.
-- =====================================================================
CREATE TABLE [activity_logs] (
    [id]             AUTOINCREMENT,
    [activity_code]  TEXT(50)  NOT NULL,
    [detail]         LONGTEXT,
    [json]         LONGTEXT,
    [username]       TEXT(64)  NOT NULL,
    [created_at]     DATETIME  NOT NULL DEFAULT Now(),
    
    CONSTRAINT 
        [pk_activity_logs] 
            PRIMARY KEY ([id])
);

CREATE INDEX 
    [idx_activity_logs_created_at] 
        ON [activity_logs] ([created_at]);
CREATE INDEX 
    [idx_activity_logs_activity_code] 
        ON [activity_logs] ([activity_code]);
