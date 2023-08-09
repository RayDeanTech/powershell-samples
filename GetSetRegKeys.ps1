<#
    Get and or Set reg keys as though you were deploying software to a computer,
    or for endpoint configuration & tracking purposes
#>

$rootPath = "HKLM:\SOFTWARE\_testing1\"

function Set-InstallationFlag {
    param(
        [string]$Project,
        [string]$Version,
        [DateTime]$InstallDate,
        [bool]$Success
    )
    
    $targetPath = join-path $rootPath $Project

    if ((Test-Path $targetPath) -ne $true) {
        New-Item -Path $targetPath -Force
    };


    Set-ItemProperty -Path $targetPath -Name "Version" -Value $Version -Force
    Set-ItemProperty -Path $targetPath -Name "InstallDate" -Value $InstallDate -Force
    Set-ItemProperty -Path $targetPath -Name "Success" -Value $Success -Force
}



function Get-InstallationFlag {
    param(
        [string]$Project
    )
    
    $targetPath = join-path $rootPath $Project
    
    if (Test-Path $targetPath) {
        
        $version = Get-ItemPropertyValue -Path $targetPath 'Version'
        $installDate = Get-ItemPropertyValue -Path $targetPath 'InstallDate'
        $success = Get-ItemPropertyValue -Path $targetPath 'Success'
        
        [PSCustomObject]@{
            Project     = $Project
            Version     = $version
            InstallDate = $installDate
            Success     = $success
        } | Format-Table -AutoSize
    }
    else {
        Write-Host "No information found for $Project."
    }
}


Set-InstallationFlag -Project "Adobe" -Version "1.0.22" -InstallDate (Get-Date) -Success $true

Get-InstallationFlag -Project "Adobe"