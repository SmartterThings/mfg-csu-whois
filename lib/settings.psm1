Function Get-Settings {
    
    # Read global settings
    $settings = Get-Content -Path "settings.json" | Out-String | ConvertFrom-Json
    Write-Output $settings
}

Export-ModuleMember Get-Settings