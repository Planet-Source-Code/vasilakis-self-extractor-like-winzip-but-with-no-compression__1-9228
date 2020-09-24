Attribute VB_Name = "modSelfExtract"

Public iFilez As Integer
Sub SelfExtract()

On Error Resume Next

Dim Size As String
Dim iFreeFile As Integer
Dim iName As String
Dim rPath As String
Dim TheFile As String
Dim rWelcome As String
Dim rAbout As String

iFreeFile = FreeFile
rPath = App.Path
If Mid(rPath, Len(rPath)) <> "\" Then rPath = rPath & "\"

curPOS = 0
i = 0
Do
i = i + 1
    Open rPath & App.EXEName & ".exe" For Binary As iFreeFile
    Seek #iFreeFile, LOF(iFreeFile) - (256 * 2) - 5 - 41 - 10 + curPOS
    iName = String(40, Chr(0))
    Get iFreeFile, , iName
    
    DoEvents
    iName = Replace$(iName, vbCr, "")
    frmSelfExtract.lblFiles.Caption = "Extracting " & iName & "..."
    frmSelfExtract.lblFiles.Refresh
    
    Seek #iFreeFile, LOF(iFreeFile) - (256 * 2) - 5 - 11 + curPOS
    Size = String(10, Chr(0))
    Get iFreeFile, , Size
    DoEvents
    Size = CCur(Size)
   DoEvents
    Seek #iFreeFile, LOF(iFreeFile) - 51 - Size - (256 * 2) - 5 + curPOS
    TheFile = String(Size, Chr(0))
    Get iFreeFile, , TheFile
    DoEvents
    Close iFreeFile
    FFile = FreeFile
    Open iName For Binary Access Write As #FFile
        Put #FFile, , TheFile
    DoEvents
    Close #FFile
    DoEvents
    curPOS = curPOS - Size - 50
DoEvents
Loop Until i >= iFilez

End
Exit Sub

Err:

Result = MsgBox("An error occured. Header may be damaged." _
    & vbCrLf & "Do you want to abort/retry?", _
    vbAbortRetryIgnore + vbExclamation, "Error")

If Result = vbRetry Then
    Resume
ElseIf Result = vbIgnore Then
    Resume Next
ElseIf Result = vbAbort Then
    End
End If

End Sub


