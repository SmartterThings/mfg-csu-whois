$SITE_URL = "https://microsoft.sharepoint.com/teams/MFGCustomerSuccessManagers"
$SITE_URL = "https://microsoft.sharepoint.com/teams/Rivendell"
$LIST_NAME = "MFG_CSU_MASTER_ACCOUNT"
$FIELD_PREFIX = "mfg_"

$list = New-PnPList -Title $LIST_NAME -Template GenericList -Url "Lists/$LIST_NAME"

Add-PnPField -DisplayName "TPID" -InternalName ($FIELD_PREFIX + "TPID") -Type Text -List $list -AddToDefaultView 
Add-PnPField -DisplayName "State" -InternalName ($FIELD_PREFIX + "State") -List $list -Type Text -AddToDefaultView
Add-PnPField -DisplayName "AE" -InternalName ($FIELD_PREFIX + "AE" ) -List $list -Type User -AddToDefaultView
Add-PnPField -DisplayName "ATS" -InternalName ($FIELD_PREFIX + "ATS" ) -List $list -Type User -AddToDefaultView
Add-PnPField -DisplayName "CSM" -InternalName ($FIELD_PREFIX + "CSM") -List $list -Type User -AddToDefaultView
Add-PnPField -DisplayName "VCSM" -InternalName ($FIELD_PREFIX + "VCSM") -List $list -Type User -AddToDefaultView
Add-PnPField -DisplayName "ACSM" -InternalName ($FIELD_PREFIX + "ACSM") -List $list -Type User -AddToDefaultView

$xml = @'
<Field 
    Type="UserMulti" 
    DisplayName="CSAM" 
    List="UserInfo" 
    Required="FALSE" 
    ID="{3a6091de-45e5-4022-be96-9f78d833d507}" 
    ShowField="Display Name" 
    UserSelectionMode="PeopleOnly" 
    StaticName="mfg_CSAM" 
    Name="mfg_CSAM" 
    Mult="TRUE" />
'@
Add-PnPFieldFromXml -List $list -FieldXml $xml