Attribute VB_Name = "modMake"
Function AddToSelfExtract(SelfExtract As String, WhatFile As ListBox, SaveAs As String) As Boolean
On Error GoTo Er

Dim iFreeFile As Integer
Dim iFreeFile2 As Integer
Dim sBuffer As String
Dim sBefore As String
Dim iFile As String

iFreeFile = FreeFile

Open SelfExtract For Binary As iFreeFile
    sBefore = String(LOF(iFreeFile), Chr(0))
    Get iFreeFile, , sBefore
Close iFreeFile

Open SaveAs For Output As iFreeFile
    wholePrint = sBefore
    For iTMP = 0 To WhatFile.ListCount - 1
        iName = frmMakeSelExtract.OnlyFileName(WhatFile.List(iTMP))
        iFreeFile2 = FreeFile
        frmMakeSelExtract.Caption = "Reading " & frmMakeSelExtract.OnlyFileName(WhatFile.List(iTMP)) & "..."
        frmMakeSelExtract.Refresh
        DoEvents
        Open WhatFile.List(iTMP) For Binary As iFreeFile2
        DoEvents
            sBuffer = String(LOF(iFreeFile2), Chr(0))
            Get iFreeFile2, , sBuffer
            Size = LOF(iFreeFile2)
            iName = String(40 - Len(iName), vbCr) & iName
            Size = String(10 - Len(Size), "0") & Size
            wholePrint = wholePrint & sBuffer & iName & Size
        DoEvents
        Close iFreeFile2
    Next iTMP
    
    rText = frmMakeSelExtract.txtWelcome.Text
    rText = String(256 - Len(rText), vbTab) & rText
    rAbout = frmMakeSelExtract.txtAbout.Text
    rAbout = String(256 - Len(rAbout), vbTab) & rAbout
    iFiles = WhatFile.ListCount
    iFilez = String(5 - Len(iFiles), vbCr) & iFiles
    frmMakeSelExtract.Caption = "Writing X-Tractor..."
    frmMakeSelExtract.Refresh
    Print #iFreeFile, wholePrint & iFilez & rText & rAbout
Close iFreeFile
AddToSelfExtract = True
Exit Function
Er:
MsgBox "An error occured while creating self extractor. Aborting...", vbCritical, "Error"
AddToSelfExtract = False
Exit Function
End Function
