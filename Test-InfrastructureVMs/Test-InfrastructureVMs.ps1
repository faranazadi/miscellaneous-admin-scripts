$hosts = @{localhost = "Faran's local machine"; thishostwontwork = "A made up host that will fail"}
[string]$onlineHosts = @()
[string]$offlineHosts = @()
[int]$timeToWait = 60

ForEach ($machine in $hosts) { 
    $response = Test-Connection -ComputerName $($machine.Key) -Quiet -ErrorAction SilentlyContinue 

    Write-Host "Host: $($machine.Key)             Online: $response"

    # if ping succeeded, add host to list of online hosts
    # if ping failed, add host to list of offline hosts
    if ($response -eq $true) {
        $onlineHosts += $($machine.Key)
    } else {
        $offlineHosts += $machine
    }
}

# only proceed with 'pausing' script if there are actually any failed hosts
# wait x seconds (in anticipation of the host/hosts coming back up)
if ($offlineHosts.Length -ge 1) {
    Start-Sleep -Seconds $timeToWait
}

# test each of the offline hosts again after waiting
ForEach ($offlineHost in $offlineHosts) {
    $response = Test-Connection -ComputerName $offlineHost -Quiet -ErrorAction SilentlyContinue 

    Write-Host "$offlineHost was re-tested. Online: $response"

    # no longer offline so lists need updating appropriately 
    if ($response -eq $true) {
        $offlineHosts -= $offlineHost
        $onlineHosts += $offlineHost
    }
}

Write-Host "Number of online hosts: " $onlineHosts.Length
Write-Host "Number of offline hosts: " $offlineHosts.Length

# send e-mail alert out
#Send-MailMessage -To alerts@ebb3.com -From anyone@ebb3.com -Subject "Blah blah blah"