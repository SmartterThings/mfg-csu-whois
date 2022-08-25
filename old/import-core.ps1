param (
    $InputFile = "data.csv"
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

Get-Content -Path $InputFile | ConvertFrom-Csv | ForEach-Object {
    $ht = ConvertTo-Hashtable $_
    Write-Output ( "Creating list item ", $ht.Title -Join "")
    Add-PnPListItem -List $spList -Values $ht
}







exit


$user = @{
    Title = "DELL"; 
    TPID = "0000"; 
    AE = "tdurham@microsoft.com";
    ATS = "Matt Hickey"
}
Add-PnPListItem -List $LIST_NAME -Values $user