$SITE_URL = "https://microsoft.sharepoint.com/teams/MFGCustomerSuccessManagers"
$SITE_URL = "https://microsoft.sharepoint.com/teams/Rivendell"
$LIST_NAME = "MFG_CSU_MASTER_ACCOUNT"
$FIELD_PREFIX = "mfg_"

$list = New-PnPList -Title $LIST_NAME -Template GenericList -Url "Lists/$LIST_NAME"

if ( $list -eq $null ) {
    $list = Get-PnPList -Identity $LIST_NAME
}

# Rename display name for "TItle"
Set-PnPField -List $list -Identity Title -Values @{ Title = "TPName"}

# TPID
Add-PnPField -DisplayName "TPID" -InternalName "TPID" -Type Text -List $list -AddToDefaultView 

# Index TPID and enforce required &uniqueness
Set-PnPField -List $list -Identity TPID -Values @{ Indexed = $true; Required = $true }
Set-PnPField -List $list -Identity TPID -Values @{ EnforceUniqueValues = $true }

# ATU
Add-PnPField -DisplayName ATU -InternalName ATU -List $list -Type Text -AddToDefaultView

# Territory
Add-PnPField -DisplayName Territory -InternalName Territory -List $list -Type Text -AddToDefaultView

# State/Province
Add-PnPField -DisplayName "State" -InternalName "State" -List $list -Type Text -AddToDefaultView

# Account Executive
Add-PnPField -DisplayName "AE" -InternalName "AE" -List $list -Type User -AddToDefaultView

# Account Technology Strategist
Add-PnPField -DisplayName "ATS" -InternalName "ATS" -List $list -Type User -AddToDefaultView

# Primary CSM
Add-PnPField -DisplayName "CSM" -InternalName "CSM" -List $list -Type User -AddToDefaultView

# Primary CSM
Add-PnPField -DisplayName "iCSM" -InternalName "iCSM" -List $list -Type User -AddToDefaultView

# Voice CSM
Add-PnPField -DisplayName "Voice CSM" -InternalName "VoiceCSM" -List $list -Type User -AddToDefaultView

# App/Plat CSM
Add-PnPField -DisplayName "App CSM" -InternalName "AppCSM" -List $list -Type User -AddToDefaultView

# CSAM: Multi-User fields must use XML for options.
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