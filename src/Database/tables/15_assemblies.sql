-- =====================================================================
-- Table  : assemblies
-- Purpose: Bill of materials - links a parent sparepart to the child
--          spareparts it is assembled from.
-- Notes  : Both [parent_id] and [child_id] reference spareparts.
--          A parent/child pair is unique, and a sparepart cannot be its
--          own child.
-- =====================================================================
CREATE TABLE [assemblies] (
	[id]          AUTOINCREMENT,
	[parent_id]   LONG      NOT NULL,
	[child_id]    LONG      NOT NULL,
	[created_at]  DATETIME  NOT NULL DEFAULT Now(),
	[updated_at]  DATETIME,

	CONSTRAINT
		[pk_assemblies]
			PRIMARY KEY ([id]),
	CONSTRAINT
		[uq_assemblies_parent_child]
			UNIQUE ([parent_id], [child_id]),
	CONSTRAINT
		[fk_assemblies_parent]
			FOREIGN KEY ([parent_id])
				REFERENCES [spareparts] ([id]),
	CONSTRAINT
		[fk_assemblies_child]
			FOREIGN KEY ([child_id])
				REFERENCES [spareparts] ([id]),
	CONSTRAINT
		[ck_assemblies_not_self]
			CHECK ([parent_id] <> [child_id])
);

CREATE INDEX
	[idx_assemblies_child]
		ON [assemblies] ([child_id]);
