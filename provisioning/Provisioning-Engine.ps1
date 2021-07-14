<#
    Adjusted provisioning script used in Banker Portal POC during period 2021/06/01 to 2021/07/02.

    Note: this script has not been tested or developed in the intended target environment. Additional changes might be required.

    Intended sites need to be created before running this script. 
#>

Import-Module -Name PnP.PowerShell

function AddMetadataToSite {
    param ([string] $siteUrl, [string] $client, [string] $oppId, [System.Management.Automation.PSCredential] $credentials)
    Set-PnPTenantSite -Url $siteUrl -NoScriptSite:$false

    Connect-PnPOnline -Url $siteUrl -Credentials $credentials
    Set-PnPPropertyBagValue -Key "invBankClient" -Value $client -Indexed 
    Set-PnPPropertyBagValue -Key "invBankOpportunityId" -Value $oppId -Indexed
    Set-PnPPropertyBagValue -Key "invBankPortalSolutionId" -Value $oppConstant -Indexed

    Request-PnPReIndexWeb
}

function AddUsers {
    param ([string] $site, [System.Management.Automation.PSCredential] $credentials)
    Connect-PnPOnline -Url $site -Credentials $credentials
    $groupId = "Banker Portal Documents Members"

    New-PnPGroup -Title $groupId -Owner $adminEmail
    Set-PnPSiteGroup -Identity $groupId -PermissionLevelsToAdd "Read"

    Set-PnPList -Identity "Documents" -BreakRoleInheritance
    Set-PnPListPermission -Identity "Documents" -Group $groupId -AddRole "Contribute"

    Disconnect-PnPOnline
}


$adminEmail = "" # Add admin UPN
$adminURL = "https://ubscloudeng-admin.sharepoint.com/"
$oppBaseUrl = "https://ubscloudeng.sharepoint.com/teams/"
$oppConstant = "portalpoc"

$cred = Get-Credential -UserName $adminEmail -Message "Type admin password"

$sites = Import-Csv ".\sites-and-members.csv"

Connect-PnPOnline -Url $adminURL -Credentials $cred

foreach ($item in $sites) {

    $client = $item.Client
    $oppId = $item.OppId

    Write-Host "Adding metadata to opportunity..."
    AddMetadataToSite -siteUrl "$($oppBaseUrl)$($oppConstant)$($oppId.Replace("-", ''))" `
        -client $client `
        -oppId $oppId `
        -credentials $cred
    Write-Host "Metadata added!"

    Write-Host "Adding group and permissions to opportunity..."
    AddUsers -site "$($oppBaseUrl)$($oppConstant)$($oppId.Replace("-", ''))" `
        -oppId $oppId `
        -cred $cred
    Write-Host "Users added!"
}
