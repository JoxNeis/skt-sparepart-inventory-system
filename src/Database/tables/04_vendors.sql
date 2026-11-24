-- =====================================================================
-- Table  : vendors
-- Purpose: Master data - suppliers that orders are placed with.
-- =====================================================================
CREATE TABLE [vendors] (
    [id]            AUTOINCREMENT,
    [code]          TEXT(40)  NOT NULL,
    [name]          TEXT(255) NOT NULL,
    [company_name]  TEXT(255),
    [email]         TEXT(255),
    [created_at]    DATETIME  NOT NULL DEFAULT Now(),
    [updated_at]    DATETIME,
    [deleted_at]    DATETIME,
    
    CONSTRAINT 
        [pk_vendors] 
            PRIMARY KEY ([id]),
    CONSTRAINT 
        [uq_vendors_code] 
            UNIQUE ([code])
);

CREATE INDEX 
    [idx_vendors_name] 
        ON [vendors] ([name]);
