Function ConvertTo-CSVFromExcel {
    param(
        [Parameter(
            Position = 0,             
            Mandatory = $true, 
            HelpMessage="The excel file to convert."
        )]
        [string] $ExcelFilePath,

        [Parameter(
            HelpMessage="Use this switch to generate the CSVs in the current working directory.")]
        [switch] $UseWorkingDirectory
    )

    If ( Test-Path -Path $ExcelFilePath -eq $false) {
        Write-Error "The specified Excel file $ExcelFilePath does not exist!";
        Exit;
    }

    $ExcelFileInfo = Get-ChildItem $ExcelFilePath;

    # Open file in Excel
    $Excel = New-Object -ComObject Excel.Application;
    $Workbook = $Excel.Workbooks.Open($ExcelFilePath);

    # Do we have multiple worksheets?
    $HasMultipleWorksheets = ($Workbook.Worksheets.Count -gt 1);

    # Keep track of each filepath that was exported
    $CsvFilePaths = @($Workbook.Worksheets.Count);

    foreach ($Worksheet in $Workbook.Worksheets) {
        
        # Save to TEMP directory unless otherwise specified
        $Directory = $env:TEMP;
        If ($UseWorkingDirectory -eq $true) {
            $Directory = $ExcelFileInfo.Directory;
        }

        # If the Excel file has multiple worksheets, create a suffix using the 
        # worksheet name. Otherwise, the output file will be the name of the 
        # Excel file.
        $Suffix = "";
        If ($HasMultipleWorksheets -eq $true) {
            $Suffix = "-", $Worksheet.Name -Join ""
        }

        $CsvFileName = ($ExcelFileInfo.BaseName, $Suffix, ".csv" -Join "").ToLower();
        $CsvFilePath = Join-Path -Path $Directory -ChildPath $CsvFileName;

        if (Test-Path -Path $CsvFilePath) {
            # Nuke from orbit, just to be sure
            Remove-Item -Path $CsvFilePath -Force
        }

        $Worksheet.SaveAs($CsvFilePath, 6);
        $CsvFilePaths += ($CsvFilePath); # The caller will want this information
    }
    $Excel.Quit();

    Write-Output $CsvFilePaths
}

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

Export-ModuleMember ConvertTo-CSVFromExcel
Export-ModuleMember Get-Settings
Export-ModuleMember Add-SpoMultiUserField
Export-ModuleMember Update-SpoListItemByPrimaryKey