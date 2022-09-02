param(
    [Parameter(
        HelpMessage = "The file to import."
    )]
    [string]    $CsvInputFile = "import-calc.csv",

    [int] $MaxRows = -1
)

Import-Module (Resolve-Path -Relative ".\lib\shared.psm1") -Force -NoClobber

$LIST_NAME = $Settings.ListName;

$spoList = Get-PnPList -Identity $LIST_NAME;

If ($spoList -eq $null) {
    Write-Host
}

$recordCount = 0
foreach($record in Import-Csv -Path $CsvInputFile) {
    $recordCount++;
    Write-Host "Importing $($recordCount)"
}