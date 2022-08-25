Function Add-MultiUserField {
    param(
        $List,
        [string] $Name,
        [bool] $AddToDefaultView = $true
    )

    $FieldXML = "<Field Type='UserMulti' Name='$Name' ID='$([GUID]::NewGuid())' DisplayName='$Name' Required ='FALSE' UserDisplayOptions='NamePhoto' UserSelectionMode='0' IsModern='TRUE' Mult='TRUE' Viewable='TRUE' ></Field>"
    $field = Add-PnPFieldFromXml -List $list -FieldXml $FieldXML

    If ( $AddToDefaultView -eq $true ) {
        $AllItemsView = Get-PnPView -List $list -Identity "All Items"
        $AllItemsView.ViewFields.Add($field.Title)
        $AllItemsView.Update()
        Invoke-PnPQuery
    }
}



$list = Get-PnPList -Identity MFG_CSU_MASTER_ACCOUNT

Add-MultiUserField -List $list -Name "CSAM"

$users = @("tdurham@microsoft.com", "matthic@microsoft.com")
Set-PnPListItem -Identity 404 -List $list -Values @{ "CSAM" = $users }

Exit

$FieldXML = "<Field Type='UserMulti' Name='CSAMTEST' ID='$([GUID]::NewGuid())' DisplayName='CSAMTEST' Required ='FALSE' UserDisplayOptions='NamePhoto' UserSelectionMode='0' IsModern='TRUE' Mult='TRUE' Viewable='TRUE' ></Field>"

$field = Add-PnPFieldFromXml -List $list -FieldXml $FieldXML

