# Refresh-Path.ps1
# 
# Refreshes the path variable in a Powershell session.
# Taken from http://stackoverflow.com/questions/14381650/how-to-update-windows-powershell-session-environment-variables-from-registry 

foreach($level in "Machine","User") {
   [Environment]::GetEnvironmentVariables($level).GetEnumerator() | % {
      # For Path variables, append the new values, if they're not already in there
      if($_.Name -match 'Path$') { 
         $_.Value = ($((Get-Content "Env:$($_.Name)") + ";$($_.Value)") -split ';' | Select -unique) -join ';'
      }
      $_
   } | Set-Content -Path { "Env:$($_.Name)" }
}
