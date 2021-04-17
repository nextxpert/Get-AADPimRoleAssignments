#Requires -Module AzureADPreview

<#
    .DESCRIPTION
    This script exports all Privileged Identity Management role assignments to a *.csv file.

#>


function Get-NextLevelAzureADDirectoryRoleAssignments {
    [CmdletBinding()]
    param ()

    process {

        $roleAssignments = @()
        $i = 0
        Write-Verbose -Message "Retrieving Privileged Identity Management Role Assignments"
        $PIMRoleAssignments = Get-AzureADMSPrivilegedRoleAssignment -ProviderId "aadRoles" -ResourceId $token.AccessToken.TenantId
        
        Write-Verbose -Message "Collect information from Assigned Roles collection"
        $PIMRoleAssignments | ForEach-Object {
            $i++
            $ProgressState = [math]::Round($i / $PIMRoleAssignments.count * 100)
            Write-Progress -Activity "Collecting PIM Information" -Status "$($ProgressState)% Complete" -PercentComplete $ProgressState
            # Process all Role Assignments
            $RoleDefinitionId = $_.RoleDefinitionId
            Write-Verbose -Message "Obtaining information for $RoleDefinitionID"
            $AADRole = Get-AzureADDirectoryRole | Where-Object { $_.RoleTemplateID -eq $RoleDefinitionId }

            $AADPrincipal = Get-AzureADObjectByObjectId -ObjectIds $_.SubjectID

            # Distinguish between users, groups and service principals
            if ($AADPrincipal.ObjectType -match "ServicePrincipal") {
                $memberName = "$($AADPrincipal.AppDisplayName) ($($AADPrincipal.AppId))"
            }
            elseif ($AADPrincipal.ObjectType -match "Group") {
                $memberName = "$($AADPrincipal.DisplayName)"
            }
            else {   
                $memberName = "$($AADPrincipal.UserPrincipalName) ($($AADPrincipal.Surname))"
            }

            if(($null -eq $_.EndDateTime) -and ($_.AssignmentState -eq "Active")) { 
                $AssignmentState = "Permanent" }
            else{ 
                $AssignmentState = $_.AssignmentState 
            }

            Write-Verbose -Message "Storing role assigment for $memberName"
            $roleAssignments += [PSCustomObject]@{
                Role       = $AADRole.DisplayName
                Member     = $memberName
                ObjectType = $AADPrincipal.ObjectType
                AssignmentState = $AssignmentState

        }
    }

    Write-Verbose -Message "Done storing $($PIMRoleAssignments.count) Role Assignments"
    return $roleAssignments
    }
}


# $VerbosePreference = "Continue"
# Check Azure AD Access token
Import-Module -Name AzureADPreview
try 
{ 
    Get-AzureADTenantDetail 
} 

catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] { 
    Write-Output "Please login to Active Directory with the Security Reader Role"
    try {
        Connect-AzureAD
    }
    catch [Microsoft.Open.Azure.AD.CommonLibrary.AadAuthenticationFailedException] {
        Write-Host "Authentication Failed"
        Exit        
    }
}


$token = [Microsoft.Open.Azure.AD.CommonLibrary.AzureSession]::AccessTokens
Write-Verbose "Connected to tenant: $($token.AccessToken.TenantId) with user: $($token.AccessToken.UserId)"


#Get all role assignments
$roleAssignments = Get-NextLevelAzureADDirectoryRoleAssignments

$exportPath = Join-Path $PSScriptRoot "AzureADDirectoryRoleAssignments_$(get-date -f yyyy-MM-dd)_$($($token.AccessToken.UserId).split("@")[1]).csv"

# Export as CSV to current directory
Write-Verbose -Message "Storing results in $exportPath"
$roleAssignments | Export-Csv -Path $exportPath -NoTypeInformation -Delimiter ";"

Write-Output "`nExported role assignments to: '$exportPath'"
Write-Output $roleAssignments | Format-Table

<# Snippet to compare two different exports
$firstReport = Import-Csv -Path "AzureADDirectoryRoleAssignments_2020-03-13 - Copy.csv" -Delimiter ";"
$secondReport = Import-Csv -Path "AzureADDirectoryRoleAssignments_2020-03-13.csv" -Delimiter ";"
Compare-Object $firstReport $secondReport
#>

<# snippet to count assignments per role
$roleAssignments | Group-Object -Property Role | Sort-Object -Descending | Select-Object -Property Name,Count
#>