param(
    [string] $InputFile,
    [array] $Fields = "TPID",
    [string] $PrimaryKeyField = "TPID"
)

Import-Module (Resolve-Path -Relative ".\lib\settings.psm1") -Force -NoClobber

$settings = Get-Settings

$spList = Get-PnPList -Identity $settings.ListName

$count = 0
Get-Content -Path $InputFile | ConvertFrom-Csv  | ForEach-Object {
    
    $pkValue = $null
    $values = @{}
    $row = $_
    $count += 1

    $row.psobject.Properties | ForEach {
        # Primary key
        if ($_.Name -eq $PrimaryKeyField) {
            $pkValue = $_.Value
        }

        # Convert to Hashtable
        if ($Fields.Contains($_.Name) ) {
            $values[$_.Name] = $_.Value
        }
    }

    if ($pkValue -eq $null) {
        Write-Error "A primary key of $PrimaryKeyField could not be found within the dataset!"
        Write-Host (ConvertTo-Json $row)
        exit
    }

    # I can't believe in this day and age I still have to do this to query SharePoint
    $query = "<View><Query><Where><Eq><FieldRef Name='" + $PrimaryKeyField + "'/><Value Type='Text'>" + $pkValue + "</Value></Eq></Where></Query></View>"
    $spListItem = Get-PnPListItem -List $spList -Query $query
    
    if ( $spListItem -eq $null ) {
        Write-Warning ( "Could not find list item with '" + $PrimaryKeyField + "' equal to '" + $pkValue + "'." )
        exit
    } else {
        $id = $spListItem.Id
        Write-Host ( $count.ToString() + ") Updating list item " + $id + " with '" + $PrimaryKeyField + "' equal to '" + $pkValue + "'." )
        Set-PnPListItem -List $spList -Identity $id -Values $values
    }
}

