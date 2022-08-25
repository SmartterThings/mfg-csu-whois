param (
    $InputFile = "csams.csv"
)

function ConvertTo-Hashtable {
    param ([PSObject] $obj)

    $ht = @{}

    $obj.psobject.Properties | ForEach {
        $ht[$_.Name] = $_.Value
    }

    return $ht
}

# Read global settings
$settings = Get-Content -Path "settings.json" | Out-String | ConvertFrom-Json
$LIST_NAME = $settings.ListName

$spList = Get-PnPList -Identity $LIST_NAME

$users = @("tdurham@microsoft.com", "matthic@microsoft.com")
Set-PnPListItem -Identity 404 -List $spList -Values @{ "CSAMs" = $users }


Exit

Get-Content -Path $InputFile | ConvertFrom-Csv | ForEach-Object {

    $caml = "<View><Query><Where><Eq><FieldRef Name='TPID'/><Value Type='Text'>" + $_.TPID + "</Value></Eq></Where></Query></View>"

    $spListItem = Get-PnPListItem -Query $caml -List $spList

    $users = $_.CSAM -Split ";"

    if ($null -eq $spListItem) {
        Write-Host ("Could not find a list item with TPID ", $_.TPID, -Join "")
    } else {
        Write-Output ( "Updating list item ", $spListItem['Title'], $_.CSAM -Join "")
        Set-PnPListItem -Identity $spListItem.Id -List $spList -Values @{ "CSAM" = $users }
    }
}