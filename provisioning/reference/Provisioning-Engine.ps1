<#
    Provisioning script used in demo environment in Banker Portal POC during period 2021/06/01 to 2021/07/02.

    Added for future reference.
#>

Import-Module -Name PnP.PowerShell

function CreateOpportunity {
    param ([string] $oppId)

    New-PnPSite -Type TeamSite `
        -Title $oppId `
        -Alias "$($oppConstant)-$($oppId.Trim())" `

}

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
    param ([string[]] $users, [string] $site, [string] $oppId, [System.Management.Automation.PSCredential] $credentials)

    Connect-PnPOnline -Url $site -Credentials $credentials
    $groupId = "$($oppId) Portal PoC Members"

    New-PnPGroup -Title $groupId -Owner $adminEmail
    Set-PnPSiteGroup -Identity $groupId -PermissionLevelsToAdd "Read"

    Set-PnPList -Identity "Documents" -BreakRoleInheritance
    Set-PnPListPermission -Identity "Documents" -Group $groupId -AddRole "Contribute"

    foreach ($user in $users) {
        Add-PnPGroupMember -LoginName $user.Trim().ToLower() -Group $groupId
    }
    Disconnect-PnPOnline
}


$adminEmail = "admin@spportaldemo.onmicrosoft.com"
$adminURL = "https://spportaldemo-admin.sharepoint.com/"
$oppBaseUrl = "https://spportaldemo.sharepoint.com/sites/"
$oppConstant = "portalpoc"

$cred = Get-Credential -UserName $adminEmail -Message "Type admin password"

$sites = Import-Csv ".\sites-and-members.csv"

Connect-PnPOnline -Url $adminURL -Credentials $cred

foreach ($item in $sites) {

    $client = $item.Client
    $oppId = $item.OppId
    $members = $item.Members.Split("|")

    Write-Host "Creating site..."
    CreateOpportunity -oppId $oppId
    Write-Host "Site created!"

    Write-Host "Adding metadata to site.."
    AddMetadataToSite -siteUrl "$($oppBaseUrl)$($oppConstant)-$($oppId)" `
        -client $client `
        -oppId $oppId `
        -credentials $cred
    Write-Host "Metadata added!"

    Write-Host "Adding users to site.."
    AddUsers -users $members.Split("|") `
        -site "$($oppBaseUrl)$($oppConstant)-$($oppId)" `
        -oppId $oppId `
        -cred $cred
    Write-Host "Users added!"
}
