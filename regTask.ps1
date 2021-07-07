[CmdletBinding()]
param ()
[xml]$XMLData = Get-Content "$PSScriptRoot\Update Wallpaper.xml"
$XMLData.Task.Actions.Exec.Arguments = "-File ""$PSScriptRoot\updateWP.ps1"" -WPOnly"
Register-ScheduledTask "Update Wallpaper" -Xml $XMLData.OuterXml
Write-Verbose -Message "$? Register ScheduledTask"