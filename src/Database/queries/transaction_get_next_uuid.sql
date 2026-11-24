-- =====================================================================
-- Query   : transaction_get_next_uuid
-- Type    : SELECT
-- Purpose : Next free document number
--           <PLANT>-<PREFIX>-<yyyymmdd>-<nnn>,
--           e.g. SBY01-ORD-20261124-001.
--           <PLANT> is manufacturing_plants.code (max 15 chars, so the
--           number is at most 32 chars and fits transactions.uuid).
--           Prefixes come from serviceTransaction: ORD, RCV, WRQ, MTW.
--           Numbering restarts at 001 per plant, prefix and day.
--           A range on [uuid] (instead of Left/Mid) lets Access use
--           the unique index. uq_transactions_uuid still rejects a
--           number taken by another user in the meantime.
-- Replaces: select_uuid
-- =====================================================================
PARAMETERS
	[p_manufacturing_plant_code] Text(15),
	[p_uuid_prefix] Text(3),
	[p_transaction_date] DateTime;

SELECT
	UCase(Trim([p_manufacturing_plant_code]))
		& '-'
		& UCase([p_uuid_prefix])
		& '-'
		& Format([p_transaction_date], 'yyyymmdd')
		& '-'
		& Format(
			IIf(
				Max(Val(Right(t.uuid, 3))) IS NULL,
				1,
				Max(Val(Right(t.uuid, 3))) + 1
			),
			'000'
		) AS next_uuid
FROM
	transactions AS t
WHERE
	t.uuid >= UCase(Trim([p_manufacturing_plant_code]))
		& '-'
		& UCase([p_uuid_prefix])
		& '-'
		& Format([p_transaction_date], 'yyyymmdd')
		& '-000'
	AND t.uuid <= UCase(Trim([p_manufacturing_plant_code]))
		& '-'
		& UCase([p_uuid_prefix])
		& '-'
		& Format([p_transaction_date], 'yyyymmdd')
		& '-999';
