param(
    [Alias("f")]
    [Parameter(HelpMessage="A comma-delimited (CSV) file containing the data to import.", Position=0, Mandatory=$true)]
    $InputFile,

    [switch] $Overwrite,

    [Parameter(HelpMessage="Array of field names to read from the InputFile. NOTE: Fields NOT specified will not be imported.")]
    [array] $FieldNames = ("TPID", "Title", "State", "Vertical"),

    [Parameter(HelpMessage="The field that will serve as the Primary Key for the row.")]
    $PrimaryKeyFieldName = "TPID",
    $MaxRows = -1
)

Import-Module (Resolve-Path -Relative ".\lib\shared.psm1") -Force -NoClobber
$Settings = Get-Settings

$SpoList = Get-PnPList -Identity $Settings.ListName

If( $null -eq $SpoList ) {
    # List does NOT exist. No need to proceed.
    Write-Error ("List '", $Settings.ListName, "' does NOT exist in the current site. Please create it first." -Join "");
    Exit
}

$RowNumber = 0
Get-Content -Path $InputFile | ConvertFrom-Csv | ForEach-Object {
    $RowNumber += 1
    $PrimaryKeyValue = $null
    $FieldValues = @{}

    $_.PSObject.Properties | ForEach-Object {
        # Primary key
        if ($_.Name -eq $PrimaryKeyField) {
            $PrimaryKeyValue = $_.Value
        }

        # Convert to Hashtable
        if ($FieldNames.Contains($_.Name) ) {
            $FieldValues[$_.Name] = $_.Value
        }
    }

    Write-Host -ForegroundColor Blue "$RowNumber) Creating list item '" -NoNewline
    Write-Host -ForegroundColor White ($_.Title) -NoNewline;
    Write-Host -ForegroundColor Blue "'...";
    
    $Err = $null
    $ListItem = Add-PnPListItem -List $SpoList -Values $FieldValues -ErrorVariable Err -ErrorAction SilentlyContinue
    If ( ( $Err -Join "").Contains("TPID") ) {
        Write-Host -ForegroundColor Yellow "List item '" -NoNewline
        Write-Host -ForegroundColor White ($_.Title) -NoNewline;
        Write-Host -ForegroundColor Yellow "' already exists.";

        if ( $Overwrite -eq $true ) {
            Update-SpoListItemByPrimaryKey -List $SpoList -PrimaryKeyFieldValue $_.TPID -FieldValues $FieldValues
        }
    } else {
        Write-Host -ForegroundColor Green "$RowNumber) Created list item '" -NoNewline
        Write-Host -ForegroundColor White ($_.Title) -NoNewline;
        Write-Host -ForegroundColor Green "'...";
    }

    If ( $MaxRows -gt 0 -and $RowNumber -eq $MaxRows) { 
        Write-Host -ForegroundColor Blue "Exiting because max number of rows ($MaxRows) reached."
        Exit 
    }
}