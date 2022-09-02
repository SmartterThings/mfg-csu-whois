param(
    [Parameter(
        HelpMessage="The file to import."
    )]
    [string]    $ExcelFileName = "CSAMs.xlsx",

    [Parameter(
        HelpMessage="The name of the worksheet to import from. The first worksheet will be used if not specified."
    )]
    [string]    $WorksheetName,    
    [int]       $BatchSize = 10,
    [string]    $TPIDColumnName = "MSSalesTPID",
    [string]    $CSAMColumnName = "EmailAlias",
    [string]    $AccountColumnName = "AccountName",
    [string]    $UserIDSuffix = "@microsoft.com",
    [int] $MaxRows = -1
)

Import-Module (Resolve-Path -Relative ".\lib\shared.psm1") -Force -NoClobber
Import-Module (Resolve-Path ".\lib\excel-lib.psm1") -Force 

$Settings = Get-Settings


$CsvFile = Export-Worksheet (Resolve-Path $ExcelFileName);

$RowCount = 0;
$Records = @{}
foreach ($item in Import-Csv $CsvFile) {
    $RowCount++;
    $TPID = $item.PSObject.Properties[$TPIDColumnName].Value;
    $CSAM = $item.PSObject.Properties[$CSAMColumnName].Value;
    $AccountName = $item.PSObject.Properties[$AccountColumnName].Value;

    Write-Host "Preparing record $($RowCount).";

    If ($TPID.Trim().Length -gt 0 -and $CSAM.Trim().Length -gt 0) {
        $CSAM = $CSAM.Trim().ToLower(), $UserIDSuffix -Join ""
        If ( $Records.ContainsKey($TPID)) {
            $Records[$TPID].CSAM += $CSAM;
        } else {
            # I can't believe in this day and age I still have to do this to query SharePoint
            $query = "<View><Query><Where><Eq><FieldRef Name='TPID'/><Value Type='Text'>" + $TPID + "</Value></Eq></Where></Query></View>"
            $spListItem = Get-PnPListItem -List $Settings.ListName -Query $query

            $ht = @{ 
                AccountName = $AccountName
                CSAM = @($CSAM)
                SpoItemId = $spListItem.Id
            }
            $Records.Add($TPID, $ht)
        }
    }

    If($MaxRows -gt -1 -and $RowCount -ge $MaxRows) { break; }
}

$RowCount = 0;
$Batch = New-PnPBatch

$Records.GetEnumerator() | ForEach-Object {
    $RowCount++;
    $SendBatch = ($RowCount % $BatchSize -eq 0 -or $RowCount -ge $Records.Count);
    $SpoItemId = $_.Value.SpoItemId;
    $CSAM = $_.Value.CSAM;
    $AccountName = $_.Value.AccountName;
    
    $Values = @{ CSAM = $CSAM }
    
    Write-Host "$($RowCount) of $($Records.Count)) Updating CSAM for $AccountName";
    Set-PnPListItem -List $Settings.ListName -Identity $SpoItemId -Values $Values -Batch $Batch

    If ( $SendBatch) {
        Write-Host "Sending Batch!"
        Invoke-PnPBatch -Batch $Batch
    }
    #Write-Host ( $_ | ConvertTo-Json )
}





