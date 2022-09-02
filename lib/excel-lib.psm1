Function Export-Worksheet {
    param(
        [Parameter(            
            Mandatory = $true, 
            HelpMessage = "The Excel file to open."
        )]
        [string] $ExcelFilePath,

        [Parameter( HelpMessage="THe name of the worksheet." )]
        $WorksheetName
    )

    $Excel, $Workbook = Get-Workbook $ExcelFilePath
    
    # Save to TEMP directory
    $Directory = $env:TEMP;
    
    $CsvFilePath = $null;

    try {
        foreach ( $Worksheet in $Workbook.Worksheets ) {
            $DoExport = $null -eq $WorksheetName -or $WorksheetName -eq $Worksheet.Name
            $CsvFileName = ("cat-", $Worksheet.Name, ".csv" -Join "").ToLower();
            $CsvFilePath = Join-Path -Path $Directory -ChildPath $CsvFileName;

            
            if (Test-Path -Path $CsvFilePath) {
                # Nuke from orbit, just to be sure
                Remove-Item -Path $CsvFilePath -Force
            }
            
            $Worksheet.SaveAs($CsvFilePath, 6);
            if ($DoExport -eq $true ) { break }
        }
    } catch {

    } finally {
        If ( $null -ne $Excel ) { $Excel.Quit(); }
        Write-Output $CsvFilePath;
    }
}
Function Get-WorksheetNames {
    param(
        [Parameter(            
            Mandatory = $true, 
            HelpMessage = "The Excel file to open."
        )]
        [string] $ExcelFilePath
    )

    $Excel, $Workbook = Get-Workbook $ExcelFilePath
    $WorksheetNames = @();
    
    foreach ($Worksheet in $Workbook.Worksheets) {
        $WorksheetNames += $Worksheet.Name;
    }

    $Excel.Quit();

    Write-Output $WorksheetNames;
}

Function Get-Workbook {
    param(
        [Parameter(      
            Mandatory = $true, 
            HelpMessage = "The Excel file to open."
        )]
        [string] $ExcelFilePath
    )

    Begin {
        $Excel = $null;
        $Workbook = $null;
    }
    
    Process {
        If ( ( Test-Path -Path $ExcelFilePath ) -eq $false) {
            Write-Error "The Excel file $ExcelFilePath does not exist!";
            Write-Output $null;
            Exit;
        }

        $FileInfo = Get-Item -Path $ExcelFilePath
        If ( $FileInfo.Extension.ToUpper() -ne ".XLSX") {
            Write-Error "The specified file ($ExcelFilePath) is NOT an Excel (.XLSX) file."
            Write-Output $null;
            Exit;
        }

        $Excel = New-Object -ComObject Excel.Application;
        $Workbook = $Excel.Workbooks.Open($ExcelFilePath);
    }

    End {
        return $Excel, $Workbook
    }    
}


Export-ModuleMember Get-Workbook;
Export-ModuleMember Get-WorksheetNames;
Export-ModuleMember Export-Worksheet