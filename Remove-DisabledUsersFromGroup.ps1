[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = 'ById')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'ById')]
    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string]$GroupId,

    [Parameter(Mandatory = $true, ParameterSetName = 'ByDisplayName')]
    [ValidateNotNullOrEmpty()]
    [string]$GroupDisplayName
)

<#
.SYNOPSIS
Finds disabled user accounts that are direct members of an Entra ID group and optionally deletes them.

.DESCRIPTION
Connects to Microsoft Graph, enumerates all direct members of the selected group, and processes only
user objects. Each user's current accountEnabled state is checked before they can become a deletion
candidate. Enabled users and users whose state cannot be verified are never deleted.

The script previews every disabled-user candidate before attempting removal. Supports -WhatIf and
-Confirm through PowerShell ShouldProcess.

.PARAMETER GroupId
Object ID of the Entra ID group to inspect.

.PARAMETER GroupDisplayName
Exact display name of the Entra ID group. It must uniquely identify one group.

.EXAMPLE
.\Remove-DisabledUsersFromGroup.ps1 -GroupId "<groupid>" -WhatIf

.EXAMPLE
.\Remove-DisabledUsersFromGroup.ps1 -GroupId "<groupid>" -Confirm

.EXAMPLE
.\Remove-DisabledUsersFromGroup.ps1 -GroupDisplayName "Former Contractors" -WhatIf

.NOTES
Requires the Microsoft.Graph.Authentication and Microsoft.Graph.Groups modules and the
Microsoft.Graph.Users module. The signed-in identity needs GroupMember.Read.All and
User.ReadWrite.All permissions, with admin consent as required by the tenant.
#>

$ErrorActionPreference = 'Stop'

$totalMembersFound = 0
$userObjectsFound = 0
$enabledUsersSkipped = 0
$disabledUsersEligible = 0
$actuallyDeleted = 0
$processingErrors = 0
$deletionCandidates = [System.Collections.Generic.List[object]]::new()

try {
    # Verify that the Graph SDK commands are available before connecting.
    $requiredCommands = @('Connect-MgGraph', 'Get-MgGroup', 'Get-MgGroupMember', 'Get-MgUser', 'Remove-MgUser')
    foreach ($commandName in $requiredCommands) {
        if (-not (Get-Command -Name $commandName -ErrorAction SilentlyContinue)) {
            throw "Required Microsoft Graph PowerShell command '$commandName' was not found. Install the Microsoft.Graph PowerShell SDK."
        }
    }

    # Request only the delegated scopes needed to read membership and manage users.
    Write-Host 'Connecting to Microsoft Graph...'
    Connect-MgGraph -Scopes @('GroupMember.Read.All', 'User.ReadWrite.All') -NoWelcome -ErrorAction Stop | Out-Null

    # Resolve the group and reject ambiguous display-name matches.
    if ($PSCmdlet.ParameterSetName -eq 'ById') {
        $group = Get-MgGroup -GroupId $GroupId -Property 'Id,DisplayName' -ErrorAction Stop
    }
    else {
        $escapedDisplayName = $GroupDisplayName.Replace("'", "''")
        $matchingGroups = @(Get-MgGroup -Filter "displayName eq '$escapedDisplayName'" -Property 'Id,DisplayName' -All -ErrorAction Stop)

        if ($matchingGroups.Count -eq 0) {
            throw "No Entra ID group was found with display name '$GroupDisplayName'."
        }
        if ($matchingGroups.Count -gt 1) {
            throw "The display name '$GroupDisplayName' matches multiple groups. Specify the group object ID instead."
        }

        $group = $matchingGroups[0]
    }

    Write-Host ("Selected group: {0} (Object ID: {1})" -f $group.DisplayName, $group.Id)

    # Fetch all direct members. The OData type is used to exclude non-user directory objects.
    Write-Host 'Retrieving all direct group members...'
    $members = @(Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop)
    $totalMembersFound = $members.Count

    foreach ($member in $members) {
        $odataType = $null
        if ($member.AdditionalProperties) {
            $odataType = $member.AdditionalProperties['@odata.type']
        }

        if ($odataType -ne '#microsoft.graph.user') {
            continue
        }

        $userObjectsFound++

        try {
            # Re-read each user so the deletion decision uses the current accountEnabled value.
            $user = Get-MgUser -UserId $member.Id -Property 'Id,DisplayName,UserPrincipalName,AccountEnabled' -ErrorAction Stop

            if ($user.AccountEnabled -eq $true) {
                $enabledUsersSkipped++
                Write-Host ("SKIPPED (enabled): {0} | {1} | {2}" -f $user.UserPrincipalName, $user.DisplayName, $user.Id)
                continue
            }

            if ($user.AccountEnabled -ne $false) {
                $processingErrors++
                Write-Warning ("SKIPPED (accountEnabled could not be verified as false): object ID {0}" -f $member.Id)
                continue
            }

            $candidate = [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName       = $user.DisplayName
                ObjectId          = $user.Id
                AccountEnabled    = $user.AccountEnabled
            }
            $deletionCandidates.Add($candidate)
            $disabledUsersEligible++
        }
        catch {
            $processingErrors++
            Write-Warning ("Could not verify user object {0}; it will not be deleted. {1}" -f $member.Id, $_.Exception.Message)
        }
    }

    # Present the complete candidate list before any removal is attempted.
    Write-Host ''
    Write-Host ("Disabled user accounts eligible for deletion: {0}" -f $disabledUsersEligible)
    if ($deletionCandidates.Count -gt 0) {
        $preview = $deletionCandidates |
            Format-Table -Property UserPrincipalName, DisplayName, ObjectId, AccountEnabled -AutoSize |
            Out-String -Width 240
        Write-Host $preview.TrimEnd()
    }
    else {
        Write-Host 'No disabled user accounts were found in the selected group.'
    }

    # ShouldProcess enforces both -WhatIf and -Confirm for each individual account removal.
    foreach ($candidate in $deletionCandidates) {
        $target = '{0} ({1}; {2})' -f $candidate.UserPrincipalName, $candidate.DisplayName, $candidate.ObjectId
        if ($PSCmdlet.ShouldProcess($target, 'Delete disabled Entra ID user')) {
            try {
                # Recheck after confirmation so a user re-enabled since preview is not removed.
                $currentUser = Get-MgUser -UserId $candidate.ObjectId -Property 'Id,DisplayName,UserPrincipalName,AccountEnabled' -ErrorAction Stop
                if ($currentUser.AccountEnabled -ne $false) {
                    if ($currentUser.AccountEnabled -eq $true) {
                        $enabledUsersSkipped++
                        Write-Host ("SKIPPED (re-enabled before deletion): {0} | {1} | {2}" -f $currentUser.UserPrincipalName, $currentUser.DisplayName, $currentUser.Id)
                    }
                    else {
                        $processingErrors++
                        Write-Warning ("SKIPPED (accountEnabled could not be verified as false before deletion): object ID {0}" -f $candidate.ObjectId)
                    }
                    continue
                }

                Remove-MgUser -UserId $candidate.ObjectId -ErrorAction Stop
                $actuallyDeleted++
                Write-Host ("DELETED: {0}" -f $target)
            }
            catch {
                $processingErrors++
                Write-Error -ErrorAction Continue ("Failed to delete {0}. {1}" -f $target, $_.Exception.Message)
            }
        }
    }
}
catch {
    $processingErrors++
    Write-Error -ErrorAction Continue ("The operation could not be completed. {0}" -f $_.Exception.Message)
}
finally {
    # Summarize discovery, safety skips, eligible candidates, and completed deletions.
    Write-Host ''
    Write-Host 'Summary'
    Write-Host ("Total group members found:                  {0}" -f $totalMembersFound)
    Write-Host ("User objects found:                         {0}" -f $userObjectsFound)
    Write-Host ("Enabled users skipped:                      {0}" -f $enabledUsersSkipped)
    Write-Host ("Disabled users eligible for deletion:       {0}" -f $disabledUsersEligible)
    Write-Host ("Users actually deleted:                     {0}" -f $actuallyDeleted)
    Write-Host ("Errors or unverifiable user states:         {0}" -f $processingErrors)
}