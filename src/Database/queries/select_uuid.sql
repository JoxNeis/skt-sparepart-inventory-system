-- Query: select_uuid
-- Type: SELECT

PARAMETERS [p_transaction_type] Text(255), [p_transaction_date] DateTime;
SELECT UCase(Left([p_transaction_type], 3))
    & "-"
    & Format([p_transaction_date], "yyyymmdd")
    & "-"
    & Right(
        "000" &
        IIf(
            Max(Val(Right([uuid], 3))) Is Null,
            1,
            Max(Val(Right([uuid], 3))) + 1
        ),
        3
    ) AS generated_uuid
FROM transactions
WHERE Left([uuid], 3) = UCase(Left([p_transaction_type], 3))
    AND Mid([uuid], 5, 8) = Format(
        [p_transaction_date],
        "yyyymmdd"
    );
