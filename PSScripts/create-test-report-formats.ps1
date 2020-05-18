function Format-TestReport{
    
    # Get root location
	$root = [System.Environment]::GetEnvironmentVariable('RUSTY_HOME','User')
	Write-Host "Root location is RUSTY_HOME : $root"

	# Make an html report
	Get-Content "$root\TestReport\test-report.csv" | ConvertFrom-Csv | ConvertTo-Html -post "<br><I>$(get-date)</I>" | Out-File "$root\TestReport\test-report.htm" -Encoding ascii

	# Make a JSON report
	Get-Content "$root\TestReport\test-report.csv" | ConvertFrom-Csv | ConvertTo-Json | Out-File "$root\TestReport\test-report.json" -Encoding ascii

	# Format as a table
	Get-Content "$root\TestReport\test-report.csv" | ConvertFrom-Csv | Format-Table > "$root\TestReport\test-report.table" 
}

# Call this function
Format-TestReport