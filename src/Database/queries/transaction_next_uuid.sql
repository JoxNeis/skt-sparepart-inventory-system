-- =====================================================================
-- Query   : transaction_next_uuid
-- Type    : SELECT
-- Purpose : Next free document number <PREFIX>-<yyyymmdd>-<nnn>,
--           e.g. ORD-20261124-001. Prefixes come from
--           serviceTransaction: ORD, RCV, WRQ.
--           A range on [uuid] (instead of Left/Mid) lets Access use
--           the unique index. uq_transactions_uuid still rejects a
--           number taken by another user in the meantime.
-- Replaces: select_uuid
-- =====================================================================
PARAMETERS
    [p_uuid_prefix] Text(3),
    [p_transaction_date] DateTime;

SELECT
    UCase([p_uuid_prefix]) & '-' & Format([p_transaction_date], 'yyyymmdd') & '-'
        & Format(IIf(Max(Val(Right(t.uuid, 3))) IS NULL, 1, Max(Val(Right(t.uuid, 3))) + 1), '000') AS next_uuid
FROM
    transactions AS t
WHERE
    t.uuid >= UCase([p_uuid_prefix]) & '-' & Format([p_transaction_date], 'yyyymmdd') & '-000'
    AND t.uuid <= UCase([p_uuid_prefix]) & '-' & Format([p_transaction_date], 'yyyymmdd') & '-999';
