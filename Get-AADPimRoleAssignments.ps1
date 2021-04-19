<#PSScriptInfo

.VERSION 3.0

.GUID 08feb122-7003-4004-868e-8c0c112fe407

.AUTHOR Raymond Comvalius

.COMPANYNAME 

.COPYRIGHT 

.TAGS 

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 
AzuerADPreview
importexcel

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES
   
#>

<#
    .DESCRIPTION
    This script exports all Privileged Identity Management role assignments to a *.xlsx file.

#>


function Get-AADPimRoleAssignments {
    [CmdletBinding()]
    param ()

    process {

        $roleAssignments = @()
        $i = 0
        Write-Verbose -Message "Retrieving Privileged Identity Management Role Assignments"
        Try{
            $PIMRoleAssignments = Get-AzureADMSPrivilegedRoleAssignment -ProviderId "aadRoles" -ResourceId $token.AccessToken.TenantId
        }
        Catch [Microsoft.Open.MSGraphBeta.Client.ApiException]{

            Write-Host "`nError collecting information from Graph API. Please check licensing and privileges." -ForegroundColor Red
            
            Exit
        }
        
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
                $AssignedName = "$($AADPrincipal.AppDisplayName) ($($AADPrincipal.AppId))"
            }
            elseif ($AADPrincipal.ObjectType -match "Group") {
                $AssignedName = "$($AADPrincipal.DisplayName)"
            }
            else {   
                $AssignedName = "$($AADPrincipal.UserPrincipalName) ($($AADPrincipal.Givenname) $($AADPrincipal.Surname))"
            }

            if(($null -eq $_.EndDateTime) -and ($_.AssignmentState -eq "Active")) { 
                $AssignmentState = "Permanent" }
            else{ 
                $AssignmentState = $_.AssignmentState 
            }

            Write-Verbose -Message "Storing role assigment for $AssignedName"
            $roleAssignments += [PSCustomObject]@{
                Role       = $AADRole.DisplayName
                Member     = $AssignedName
                ObjectType = $AADPrincipal.ObjectType
                AssignmentState = $AssignmentState
            }

            if ($AADPrincipal.ObjectType -match "Group") {
                $GroupMembers = Get-AzureADGroupMember -ObjectId $AADPrincipal.objectId -All $true
                $GroupMembers | ForEach-Object{
                    $GroupMemberName = "$($_.UserPrincipalName) ($($_.Givenname) $($_.Surname))"

                    $roleAssignments += [PSCustomObject]@{
                        Role            = $AADRole.DisplayName
                        Member          = $GroupMemberName
                        ObjectType      = "Member of $AssignedName"
                        AssignmentState = $AssignmentState
                    }
                }
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
        Write-Output "Error: Authentication Failed! `n" -ForegroundColor Red
        Exit        
    }
}


$token = [Microsoft.Open.Azure.AD.CommonLibrary.AzureSession]::AccessTokens
Write-Verbose "Connected to tenant: $($token.AccessToken.TenantId) with user: $($token.AccessToken.UserId)"


#Get all role assignments
$roleAssignments = Get-AADPimRoleAssignments

$excelExportPath = Join-Path "$(Get-Location)" "AzureADDirectoryRoleAssignments_$(get-date -f yyyy-MM-dd)_$($($token.AccessToken.UserId).split("@")[1]).xlsx"

# Export to Excel sheet to current directory
if ($roleAssignments.count -gt 0){
    $roleAssignments | Export-Excel -Path $excelExportPath -tablestyle medium16 -AutoSize -title "$($($token.AccessToken.UserId).split("@")[1])" -TitleSize 18 -TitleBold -WorksheetName ("PIM Assignments")

    Write-Output $roleAssignments | Format-Table

    Write-Output "`nExported role assignments to: '$excelExportPath'"
}
