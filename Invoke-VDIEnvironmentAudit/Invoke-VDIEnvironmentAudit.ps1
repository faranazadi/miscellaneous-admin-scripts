# MAKE SURE XENSERVER SDK MODULE HAS BEEN DOWNLOADED BEFORE RUNNING
# C:\Program Files\WindowsPowerShell\Modules

# Add the Citrix snap ins
Add-PSSnapIn Citrix*

# Import XenServer SDK
# To list all XenServer SDK cmdlets, use: Get-Module -Name XenServerPSModule | Select -ExpandProperty ExportedCommands 
Import-Module XenServerPSModule

# Init/declare variables
$xenHosts = @()
$primaryDDC = ""
$emailSubject = "Infrastructure Audit Results"
$emailRecipient = ""
$emailSender = ""
$smtpServer = ""
$auditResultsPath = "$PSScriptRoot\Infrastructure Audit Results.csv"


function main() {
    $allMachines = Get-AllMachines -DDC $primaryDDC

    Get-GoldImageUsed -machines $allMachines

    # Export to CSV to same directory the script is executed from
    $allMachines | Export-Csv -Force -NoTypeInformation -Path $auditResultsPath

    Send-MailMessage -Subject $emailSubject -To $emailRecipient -From $emailSender -Attachments $auditResultsPath -SmtpServer $smtpServer
}

# Go through each machine object in $allMachines to grab all UNIQUE XenServer names
ForEach ($machine in $allMachines) {
    $currentXenHost = $machine.HostingServerName
    ForEach ($host in $xenHosts) {
        For (int i = 0; i > $xenHosts.Count; i++) {
            if ($currentXenHost -ne $host) {
                $xenHosts += $currentXenHost
            }
        }
        
    }
}

# Connect to each of the Xen hosts
ForEach ($xenHost in $xenHosts) {
    Connect-XenServer -Server $xenHost -UserName '' -Password ''
}



# Helper functions

# Iterate over all of the machines in the environment, grab the fields we're interested in, then store them in $allMachines
function Get-AllMachines([string]$DDC) {
    return $allMachines = @(Get-BrokerMachine -AdminAddress '' -MaxRecordCount 1000 | Select-Object HostedMachineName,IPAddress,AgentVersion,CatalogName,DesktopGroupName,AllocationType,HostingServerName,HypervisorConnectionName,ImageOutOfDate,InMaintenanceMode,IsAssigned,IsPhysical,IsReserved,ProvisioningType,OSType,OSVersion,VMToolsState) 
}

# Gets the gold image used by each machine
function Get-GoldImageUsed([array]$machines) {
    ForEach ($machine in $machines) {
        $imageUsed = Get-ProvScheme -AdminAddress '' -ProvisioningSchemeUid (Get-BrokerCatalog -Name ($machine.CatalogName)).ProvisioningSchemeId | Select-Object MasterImageVM
        
        $imageUsedSubStrings = @($imageUsed.MasterImageVM.Split("\"))

        ForEach ($string in $imageUsedSubStrings) {
            Write-Host "Machine: $($machine.HostedMachineName)" 
            Write-Host "Image used: $($imageUsedSubStrings[-1])"
        }

        $machine = $machine | Add-Member -MemberType NoteProperty -Name GoldImage -Value $($imageUsedSubStrings[-1])
    }
}

function Connect-XenHost([String]$server, [String]$user, [String]$password) {
    Write-Host ("Connecting to host '{0}'" -f $server)
    $session = Connect-XenServer -Server $server -UserName $user -Password $password -PassThru

    if ($null -eq $session) {
        return $false
    }
    return $true

}

function Disconnect-XenHost([String]$server) {
    Write-Host ("Disconnecting from host '{0}'" -f $server)
    Get-XenSession -Server $server | Disconnect-XenServer

    if ($null -eq (Get-XenSession -Server $server)) {
        return $true
    }
    return $false
}

function Get-PoolMaster() {
    $pool = Get-XenPool
    return Get-XenHost -Ref $pool.master
}


# Script entry point
main
