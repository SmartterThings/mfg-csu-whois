$SITE_URL = "https://microsoft.sharepoint.com/teams/MFGCustomerSuccessManagers"
$SITE_URL = "https://microsoft.sharepoint.com/teams/Rivendell"
$LIST_NAME = "MFG_CSU_MASTER_ACCOUNT"
$FIELD_PREFIX = "mfg_"

$list = New-PnPList -Title $LIST_NAME -Template GenericList -Url "Lists/$LIST_NAME"

if ( $list -eq $null ) {
    $list = Get-PnPList -Identity $LIST_NAME
}

Set-PnPField -List $list -Identity Title -Values @{ Title = "TPName"}
Add-PnPField -DisplayName "TPID" -InternalName "TPID" -Type Text -List $list -AddToDefaultView 
Add-PnPField -DisplayName ATU -InternalName ATU -List $list -Type Text -AddToDefaultView
Add-PnPField -DisplayName Territory -InternalName Territory -List $list -Type Text -AddToDefaultView
Add-PnPField -DisplayName "State" -InternalName "State" -List $list -Type Text -AddToDefaultView
Add-PnPField -DisplayName "AE" -InternalName "AE" -List $list -Type User -AddToDefaultView
Add-PnPField -DisplayName "ATS" -InternalName "ATS" -List $list -Type User -AddToDefaultView
Add-PnPField -DisplayName "CSM" -InternalName "CSM" -List $list -Type User -AddToDefaultView
Add-PnPField -DisplayName "VCSM" -InternalName "VCSM" -List $list -Type User -AddToDefaultView
Add-PnPField -DisplayName "ACSM" -InternalName "ACSM" -List $list -Type User -AddToDefaultView

# Multi-User fields must use XML for options.
$xml = @'
<Field 
    Type="UserMulti" 
    DisplayName="CSAM" 
    List="UserInfo" 
    Required="FALSE" 
    ID="{3a6091de-45e5-4022-be96-9f78d833d507}" 
    ShowField="Display Name" 
    UserSelectionMode="PeopleOnly" 
    StaticName="CSAM" 
    Name="CSAM" 
    Mult="TRUE" />
'@
Add-PnPFieldFromXml -List $list -FieldXml $xml