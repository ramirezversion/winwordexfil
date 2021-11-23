# winwordexfil
A kind of malware to exfiltrate Word documents text written

Write code for Windows 10 that:

    - Is persistent (survives reboots).
    - It is as file-less as possible (should leave a minimal amount of code on disk, for as little time as possible).
    - Requires as little privileges as possible.
    - Searches for Word processes and extracts all text written (both for already open documents and for new Word instances).
    - Send the gathered data to a server.
    
