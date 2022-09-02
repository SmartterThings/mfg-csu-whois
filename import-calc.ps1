param(
    [Parameter(
        HelpMessage = "The file to import."
    )]
    [string]    $CsvInputFile = "import-calc.csv",

    [int] $MaxRows = -1,
    [int] $BatchSize = 10
)

Function IsNotNullOrBlank {
    param (
        [Parameter(Position=0)]
        $value
    )

    If ($null -ne $value -and $value.Trim() -ne "") {
        return $true;
    } else {
        return $false;
    }
}

Import-Module (Resolve-Path -Relative ".\lib\shared.psm1") -Force -NoClobber

$Settings = Get-Settings
$LIST_NAME = $Settings.ListName;

$LIST = Get-PnPList -Identity $LIST_NAME;


If ($LIST -eq $null) {
    
}

$recordCount = 0
$records = Import-Csv -Path $CsvInputFile
$totalRecords = $records.Count
$totalBatches = [math]::Round($totalRecords / $BatchSize, 0);
$totalBatches
$batch = New-PnPBatch
foreach($record in $records) {
    $recordCount++;
    
    $values = @{
        TPID = $record.TPID
        Title = $record.TPNAME.ToUpper()
        ATU = $record.ATU.ToUpper()
        Territory = $record.Territory.ToUpper()
        City = $record.City.ToUpper()
        State = $record.StateOrProvince.ToUpper()
        PostalCode = $record.PostalCode.ToUpper()
        AppCSM = $record.AppCSM
        Vertical = ($record.Vertical, $record.VerticalCategory -join ": ").ToUpper()
    }

    $CSM = $record.CSM;
    If (IsNotNullOrBlank $CSM) {
        $values.Add("CSM", $CSM);
    }    

    $AE = $record.AE;
    If (IsNotNullOrBlank $AE) {
        $values.Add("AE", $AE);
    }    

    $ATS = $record.ATS;
    If (IsNotNullOrBlank $ATS) {
        $values.Add("ATS", $ATS);
    }    

    $CSAM = $record.CSAMs;
    If (IsNotNullOrBlank $CSAM) {
        $values.CSAM = $CSAM -split ","
    }    

    $iCSM = $record.iCSM;
    if (IsNotNullOrBlank $iCSM) {
        $values.Add("iCSM", $iCSM);
    }

    $VoiceCSM = $record.VoiceCSM;
    if (IsNotNullOrBlank $iCSM) {
        $values.Add("VoiceCSM", $VoiceCSM);
    }    

    Write-Host "$($recordCount) of $($records.Count)] Importing $($record.TPName)";

    $index = @()
    foreach($key in $values.Keys) {
        $index += $values[$key]
    }

    $values.Add("SysIndex", $index -join "`n")

    Add-PnPListItem -List $LIST -Values $values -Batch $batch

    If ($recordCount % $BatchSize -eq 0) {
        Write-Host "Sending batch $($recordCount / $batch.RequestCount) of $($totalBatches)"
        Invoke-PnPBatch -Batch $batch
        $batch = New-PnPBatch
    }  
}
Invoke-PnPBatch -Batch $batch -ErrorAction SilentlyContinue


