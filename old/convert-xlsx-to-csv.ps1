param(
    [Parameter(Position=0, Mandatory=$true)]
    [string] $ExcelFileName,
    [switch] $UseTempDirectory
)

$FilePath = Resolve-Path $ExcelFileName;
$FileInfo = Get-ChildItem $FilePath;
$Excel = New-Object -ComObject Excel.Application;
$Workbook = $Excel.Workbooks.Open($FilePath);

# Do we have multiple worksheets
$HasMultipleWorksheets = ($Workbook.Worksheets.Count -gt 1);

# Keep track of each filepath that was exported
$OutFilePaths = @($Workbook.Worksheets.Count);

foreach ($Worksheet in $Workbook.Worksheets) {
    $Suffix = "";
    $Directory = $FileInfo.Directory;
    If ($UseTempDirectory -eq $true) {
        $Directory = $env:TEMP
    }
    If ($HasMultipleWorksheets -eq $true) {
        # Create a suffix for the filename that includes the worksheet name
        $Suffix = "-", $Worksheet.Name -Join ""
    }
    $CsvFileName = ($FileInfo.BaseName, $Suffix, ".csv" -Join "").ToLower();
    $CsvFilePath = Join-Path -Path $Directory -ChildPath $CsvFileName;

    if (Test-Path -Path $CsvFilePath) {
        # Nuke from orbit, just to be sure
        Remove-Item -Path $CsvFilePath -Force
    }

    $Worksheet.SaveAs($CsvFilePath, 6);
    $OutFilePaths += ($CsvFilePath);
}
$Excel.Quit();

Write-Output $OutFilePaths



