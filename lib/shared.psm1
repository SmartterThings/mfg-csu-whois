Function Get-Settings {
    
    # Read global settings
    $settings = Get-Content -Path "settings.json" | Out-String | ConvertFrom-Json
    Write-Output $settings
}

Function Add-SpoMultiUserField {
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


Function Update-SpoListItemByPrimaryKey {
    param(
        [Parameter(Mandatory = $true)]
        [string] $PrimaryKeyFieldValue,

        [Parameter(Mandatory=$true)]
        $List,

        [Parameter(HelpMessage = "The field that serves as the Primary Key for the row.")]
        $PrimaryKeyFieldName = "TPID",

        [Parameter(Mandatory=$true)]
        $Values
    )

    $Query = "<View><Query><Where><Eq><FieldRef Name='" + $PrimaryKeyFieldName + "'/><Value Type='Text'>" + $PrimaryKeyFieldValue + "</Value></Eq></Where></Query></View>"

    $SpoListItem = Get-PnPListItem -Query $Query -List $List

    Set-PnPListItem -Identity $SpoListItem.Id -List $list -Values $Values | Out-Null
}

Export-ModuleMember Get-Settings
Export-ModuleMember Add-SpoMultiUserField
Export-ModuleMember Update-SpoListItemByPrimaryKey