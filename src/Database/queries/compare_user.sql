-- Query: compare_user
-- Type: SELECT

PARAMETERS [p_requestor_username] Text(255), [p_approver_username] Text(255);
SELECT IIf(
        approver_user.tier > requestor_user.tier,
        True,
        False
    ) AS can_approve
FROM users AS requestor_user, users AS approver_user
WHERE requestor_user.username = [p_requestor_username]
    AND approver_user.username = [p_approver_username];
