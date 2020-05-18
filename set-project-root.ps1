function Set-ProjectRoot{
    
    # Set root location
	[System.Environment]::SetEnvironmentVariable('RUSTY_HOME',$PSScriptRoot,[System.EnvironmentVariableTarget]::User)
	$root = [System.Environment]::GetEnvironmentVariable('RUSTY_HOME','User')
	Write-Host "Root location set to RUSTY_HOME : $root"
}

# Call this function
Set-ProjectRoot