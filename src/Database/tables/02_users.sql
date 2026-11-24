-- =====================================================================
-- Table  : users
-- Purpose: Master data - application users. [username] is the Windows
--          login and is the value stored in every *_by column.
-- Notes  : [tier] drives approvals: the approver needs a higher tier
--          than the requester (see user_check_can_approve).
-- =====================================================================
CREATE TABLE [users] (
	[id]                      AUTOINCREMENT,
	[manufacturing_plant_id]  LONG       NOT NULL,
	[username]                TEXT(64)   NOT NULL,
	[first_name]              TEXT(100)  NOT NULL,
	[middle_name]             TEXT(100),
	[last_name]               TEXT(100),
	[jobdesk]                 TEXT(255),
	[tier]                    LONG       NOT NULL DEFAULT 0,
	[created_at]              DATETIME   NOT NULL DEFAULT Now(),
	[updated_at]              DATETIME,
	[deleted_at]              DATETIME,

	CONSTRAINT
		[pk_users]
			PRIMARY KEY ([id]),
	CONSTRAINT
		[uq_users_username]
			UNIQUE ([username]),
	CONSTRAINT
		[fk_users_manufacturing_plant]
			FOREIGN KEY ([manufacturing_plant_id])
				REFERENCES [manufacturing_plants] ([id]),
	CONSTRAINT
		[ck_users_tier]
			CHECK ([tier] >= 0)
);
