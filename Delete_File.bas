Sub SelfDeleteWorkbook()
    Dim wsShell As Object
    Dim tempFilePath As String
    Dim script As String
    Dim wbPath As String
    Dim tempFile As Integer

    ' Get the full path of the workbook
    wbPath = ThisWorkbook.FullName
    
    ' Generate a temporary file path for the script
    tempFilePath = Environ("Temp") & "\DeleteWorkbook.bat"
    
    ' Create the script content
    script = "@echo off" & vbCrLf & _
             "ping 127.0.0.1 -n 5 > nul" & vbCrLf & _
             "del """ & wbPath & """" & vbCrLf & _
             "del """ & tempFilePath & """"
    
    ' Write the script to the temporary file
    tempFile = FreeFile
    Open tempFilePath For Output As tempFile
    Print #tempFile, script
    Close tempFile
    
    ' Create a new instance of the WScript.Shell object
    Set wsShell = CreateObject("WScript.Shell")
    
    ' Ensure the script file exists before running it
    If Dir(tempFilePath) <> "" Then
        wsShell.Run """" & tempFilePath & """", 0       ' Run the batch file invisibly
    Else
        'MsgBox "The script file was not created.", vbCritical
        Exit Sub
    End If

    ' Close the workbook without saving changes
    ThisWorkbook.Close SaveChanges:=False
End Sub

