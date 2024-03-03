$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = "test-case\"
$watcher.Filter = "*.xlsx"
$watcher.IncludeSubdirectories = $true
$watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite -bor [System.IO.NotifyFilters]::FileName


$action =
{
	$path = $Event.SourceEventArgs.Name
	if ($path -eq "Test.xlsx" -or $path.StartsWith("~$") -or -not $path.EndsWith(".xlsx"))
	{
		Write-Host "$(Get-Date) $path skip"
		return
	}
	Start-Process -FilePath "test-case\@drop-me.bat" -ArgumentList $path -WorkingDirectory "test-case\" -NoNewWindow -Wait
	Write-Host "$(Get-Date) $path done"
}

Register-ObjectEvent $watcher "Changed" -Action $action
Register-ObjectEvent $watcher "Renamed" -Action $action

while ($true)
{
	Start-Sleep 1
}
