-- Table: manufacturing_plant_stocks
CREATE TABLE [manufacturing_plant_stocks] (
    [id] AUTOINCREMENT,
    [sparepart_id] LONG NOT NULL,
    [manufacturing_plant_id] LONG NOT NULL,
    [current_stock]  LONG NOT NULL DEFAULT 0,
    [reorder_point]  LONG NOT NULL DEFAULT 0,
    [safety_stock] LONG NOT NULL DEFAULT 0,
    [lead_time] LONG NOT NULL DEFAULT 0,
    [average_daily_usage] DOUBLE NOT NULL DEFAULT 0,
    [sd_demand] DOUBLE NOT NULL DEFAULT 0,
    [created_at] DATETIME,
    [updated_at] DATETIME,
    CONSTRAINT [pk_manufacturing_plant_stocks]
        PRIMARY KEY ([id]),
    CONSTRAINT [fk_mp_stocks_sparepart]
        FOREIGN KEY ([sparepart_id])
        REFERENCES [spareparts] ([id]),
    CONSTRAINT [fk_mp_stocks_manufacturing_plant]
        FOREIGN KEY ([manufacturing_plant_id])
        REFERENCES [manufacturing_plants] ([id])
);

CREATE INDEX [idx_mp_stocks_sparepart] ON [manufacturing_plant_stocks] ([sparepart_id]);
CREATE INDEX [idx_mp_stocks_plant] ON [manufacturing_plant_stocks] ([manufacturing_plant_id]);
CREATE UNIQUE INDEX [uq_mp_stocks_sparepart_plant] ON [manufacturing_plant_stocks] ([sparepart_id], [manufacturing_plant_id]);