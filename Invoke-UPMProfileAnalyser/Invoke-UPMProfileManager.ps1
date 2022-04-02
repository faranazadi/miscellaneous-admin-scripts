# The directory that the script resides
$scriptRootDir = Split-Path $MyInvocation.InvocationName

# The path at which the NTUSER.dat file resides
$NTUserPath = "\\UPMProfiles\*.dat"

# The amount of time since the NTUSER.dat file has been written to
$monthsInactive = 3

# Do we want to print useful messages to console?
$debugMode = $false

# Recursively get each user's NTUSER.dat and store their user path, last access time and last write time
$allNTUsers = Get-ChildItem -Path $NTUserPath -Recurse -Force | Where-Object {$_.Name -match 'NTUSER.dat'} | Select-Object FullName,LastAccessTime,LastWriteTime

if ($debugMode -eq $true) {
    Write-Host "All users: $allNTUsers" 
}

# Filter the users that have been inactive for 3 months out of $allNTUsers
$inactiveNTUsers = $allNTUsers | Where-Object { $_.LastWriteTime -le ((Get-Date).AddMonths(-$monthsInactive)) } 

if ($debugMode -eq $true) {
    Write-Host "Inactive users: $inactiveNTUsers"
}

# Export both lists of users
$allNTUsers | Export-Csv -Path $scriptRootDir\UPM_Profiles.csv -NoTypeInformation
$inactiveNTUsers | Export-Csv -Path $scriptRootDir\Inactive_UPM_Profiles.csv -NoTypeInformation