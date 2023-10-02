# 32 on 32; 64 on 64
$uninstallKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
# 32 on 64
$uninstallKeyPathWOW6432 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"


# Show all uinstall strings as a table
<#
Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach-Object {
    $subkey = $_
    $displayName = $subkey.GetValue("DisplayName")
    $displayVersion = $subkey.GetValue("DisplayVersion")
    $uninstallString = $subkey.GetValue("UninstallString")
    [PSCustomObject]@{
        DisplayName = $displayName
        DisplayVersion = $displayVersion
        UninstallString = $uninstallString
    }
}  | Format-Table -AutoSize
#>

# Search for Application's uninstall string
# use -like with * or ?
# * match zero or more characters
# ? match one character in that position
# about_Wildcards:  https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_wildcards?view=powershell-7.3
Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | ForEach-Object {
    $subkey = $_
    $displayName = $subkey.GetValue("DisplayName")
    $displayVersion = $subkey.GetValue("DisplayVersion")
    $uninstallString = $subkey.GetValue("UninstallString")
    [PSCustomObject]@{
        DisplayName = $displayName
        DisplayVersion = $displayVersion
        UninstallString = $uninstallString
    }
} | Where-Object { $_.DisplayName -like "*nvidia*" } | Format-Table -AutoSize
