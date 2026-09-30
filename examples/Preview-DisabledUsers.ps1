# Preview disabled-user candidates without deleting any accounts.
# Replace the sample ID with the object ID of your Entra ID group.

$groupId = '<group-object-id>'

& "$PSScriptRoot\..\Remove-DisabledUsersFromGroup.ps1" -GroupId $groupId -WhatIf
