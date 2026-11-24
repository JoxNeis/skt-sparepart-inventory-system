-- Table: post_mobiles
CREATE TABLE [post_mobiles] (
    [id] AUTOINCREMENT,
    [manufacturing_plant_id] LONG NOT NULL,
    [name] TEXT(255) NOT NULL,
    [code] TEXT(50) NOT NULL,
    CONSTRAINT [pk_post_mobiles] PRIMARY KEY ([id]),
    CONSTRAINT [fk_post_mobiles_manufacturing_plant] FOREIGN KEY ([manufacturing_plant_id]) REFERENCES [manufacturing_plants] ([id])
);

CREATE INDEX [idx_post_mobiles_plant] ON [post_mobiles] ([manufacturing_plant_id]);
CREATE UNIQUE INDEX [uq_post_mobiles_code] ON [post_mobiles] ([code]);
