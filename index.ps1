param(

    [Parameter(HelpMessage = "Array of field names to index.")]
    [array] $FieldNames = ("CSM", "iCSM", "VoiceCSM", "AppCSM"),

    $MaxRows = -1
)

Function Get-IndexedFieldValue {
    Param(
        $FieldValue
    )

    Write-Output ($FieldValue.LookupValue, $FieldValue.Email -Join " - " );

}

Import-Module (Resolve-Path -Relative ".\lib\shared.psm1") -Force -NoClobber
$Settings = Get-Settings

$SpoList = Get-PnPList -Identity $Settings.ListName

If ( $null -eq $SpoList ) {
    # List does NOT exist. No need to proceed.
    Write-Error ("List '", $Settings.ListName, "' does NOT exist in the current site. Please create it first." -Join "");
    Exit
}

$SpoListItems = Get-PnPListItem -List $SpoList

$RowNumber = 0
$SpoListItems | ForEach-Object {
    $RowNumber += 1

    $SysUserIndexValues = @();

    $SpoListItem = $_;

    $FieldNames | ForEach-Object {
        $SysUserIndexValues += Get-IndexedFieldValue -FieldValue $SpoListItem.FieldValues[$_];
    }
     
    $IndexedFieldValue = @{ "SysUserIndex" = $SysUserIndexValues -Join "`n" }
    
    Write-Host -ForegroundColor Blue "$RowNumber) Indexing list item '" -NoNewline
    Write-Host -ForegroundColor White ($SpoListItem.FieldValues["Title"]) -NoNewline;
    Write-Host -ForegroundColor Blue "'...";
    Write-Host ( ConvertTo-Json $IndexedFieldValue )

    Set-PnPListItem -List $SpoList -Identity $SpoListItem.Id -Values $IndexedFieldValue | Out-Null
    

    # For testing/debugging - Stop after N number of updates.
    If ( $MaxRows -gt 0 -and $RowNumber -eq $MaxRows) { 
        Write-Host -ForegroundColor Yellow "🛑 Exiting because max number of rows ($MaxRows) reached."
        Exit 
    }
}

