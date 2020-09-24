VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMakeSelExtract 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "X-Tractor"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   Icon            =   "frmMakeSelExtract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      Picture         =   "frmMakeSelExtract.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      Picture         =   "frmMakeSelExtract.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2760
      Picture         =   "frmMakeSelExtract.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      Picture         =   "frmMakeSelExtract.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   795
   End
   Begin VB.TextBox txtAbout 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "frmMakeSelExtract.frx":154A
      Top             =   2280
      Width           =   3495
   End
   Begin VB.TextBox txtWelcome 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "frmMakeSelExtract.frx":1587
      Top             =   840
      Width           =   3495
   End
   Begin VB.ListBox lstFiles 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3720
      Picture         =   "frmMakeSelExtract.frx":15B7
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   915
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      Picture         =   "frmMakeSelExtract.frx":19F9
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   915
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to choose what files are going to be included at the x-tractor."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose where the new file will be writen."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   405
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose the X-tractor."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmMakeSelExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function OnlyFileName(file) As String
If InStr(file, "\") = 0 Then OnlyFileName = file: Exit Function
rTMP = 1
Do
    rTMP0 = rTMP
    rTMP = InStr(rTMP + 1, file, "\")
Loop Until rTMP = 0
OnlyFileName = Right(file, Len(file) - Len(Left(file, rTMP0)))
End Function


Private Sub Command1_Click()
On Error GoTo UsrCancel
CommonDialog1.CancelError = True
CommonDialog1.Filter = "Executable Files|*.exe|"
CommonDialog1.Flags = cdlOFNFileMustExist
CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then Exit Sub
Text1 = CommonDialog1.FileName
UsrCancel:
End Sub

Private Sub Command2_Click()
On Error GoTo UsrCancel
CommonDialog1.CancelError = True
CommonDialog1.Filter = "All Files|*.*|"
CommonDialog1.Flags = cdlOFNFileMustExist
CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then Exit Sub
For i = 0 To lstFiles.ListCount - 1
    If LCase$(OnlyFileName(CommonDialog1.FileName)) = LCase$(OnlyFileName(lstFiles.List(i))) Then MsgBox "A file with the same name exists!", vbExclamation, "Oops!": Exit Sub
Next i
lstFiles.AddItem CommonDialog1.FileName
UsrCancel:

End Sub

Private Sub Command3_Click()
'check something first...
rTMP = Caption
If Text1.Text = "" Then MsgBox "You must choose the X-tractor!", vbExclamation, "Oops!": Text1.SetFocus: Exit Sub
If Text3.Text = "" Then MsgBox "You must choose the output filename!", vbExclamation, "Oops!": Text3.SetFocus: Exit Sub
If lstFiles.ListCount = 0 Then MsgBox "You must add files!", vbExclamation, "Oops!": Exit Sub
'if everything is ok continue...
If AddToSelfExtract(Text1, Me.lstFiles, Text3) = True Then
    MsgBox "Done!", vbInformation, "Done!"
End If
Caption = rTMP



End Sub

Private Sub Command4_Click()
On Error GoTo UsrCancel
CommonDialog1.CancelError = True
CommonDialog1.Filter = "Executable Files|*.exe|"
CommonDialog1.Flags = cdlOFNCreatePrompt Or cdlOFNOverwritePrompt
CommonDialog1.ShowSave

If CommonDialog1.FileName = "" Then Exit Sub
Text3 = CommonDialog1.FileName
UsrCancel:

End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
On Error Resume Next
lstFiles.RemoveItem lstFiles.ListIndex
End Sub

Private Sub Form_Load()
Show
Text1.SetFocus
Text1.Text = Command$
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub txtAbout_Change()
If Len(txtAbout.Text) > 256 Then txtAbout.Text = Left(txtAbout.Text, 256)
End Sub

Private Sub txtWelcome_Change()
If Len(txtWelcome.Text) > 256 Then txtWelcome.Text = Left(txtWelcome.Text, 256)
End Sub


