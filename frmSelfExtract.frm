VERSION 5.00
Begin VB.Form frmSelfExtract 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "+ X-Tractor +"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelfExtract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdView 
      Caption         =   "Files"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblFiles 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   5775
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmSelfExtract.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to X-tractor!"
      Height          =   810
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3720
   End
End
Attribute VB_Name = "frmSelfExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAbout_Click()
If cmdAbout.Tag = "" Then
    MsgBox "X-Tractor by Vasilis Sagonas." & vbCr & vbCr & "Contact - vsag@forthnet.gr", vbInformation, "About..."
Else
    MsgBox cmdAbout.Tag & vbCr & vbCr & "X-Tractor by Vasilis Sagonas." & vbCr & "Contact - vsag@forthnet.gr", vbInformation, "About..."
End If
End Sub

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdExtract_Click()
cmdView.Enabled = False
cmdExtract.Enabled = False
SelfExtract
End Sub

Private Sub cmdView_Click()
frmFiles.Show 1
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim rWelcome As String
Dim rAbout As String
Dim iFiles As String
Dim iName As String
Dim Size As String
iFreeFile = FreeFile
curPOS = 0
i = 0
Open rPath & App.EXEName & ".exe" For Input As iFreeFile
Close iFreeFile
iFreeFile = FreeFile
Open rPath & App.EXEName & ".exe" For Binary As iFreeFile
    
    Seek #iFreeFile, LOF(iFreeFile) - 6 - (256 * 2)
    iFiles = String(5, Chr(0))
    Get iFreeFile, , iFiles

    iFiles = Replace$(iFiles, vbCr, "")
    iFilez = Val(iFiles)
    lblFiles.Caption = "This file contains " & iFilez & " files."
    
Close iFreeFile
rWelcome = Replace(rWelcome, vbTab, "")
If rWelcome <> "" Then lblWelcome.Caption = rWelcome

rAbout = Replace(rAbout, vbTab, "")
If rAbout <> "" Then cmdAbout.Tag = rAbout

Do
i = i + 1
    Open rPath & App.EXEName & ".exe" For Binary As iFreeFile

    Seek #iFreeFile, LOF(iFreeFile) - (256 * 2) - 5 - 41 - 10 + curPOS
    iName = String(40, Chr(0))
    Get iFreeFile, , iName
    
    
    Seek #iFreeFile, LOF(iFreeFile) - (256 * 2) - 5 - 11 + curPOS
    Size = String(10, Chr(0))
    Get iFreeFile, , Size
        
    Size = CCur(Size)
    
    Close iFreeFile
    FFile = FreeFile
    iName = Replace$(iName, vbCr, "")
    
    frmFiles.lstFiles.AddItem iName
    
    curPOS = curPOS - Size - 50

Loop Until i >= iFilez

Show
Refresh
Exit Sub
Err:
MsgBox "This file is damaged or it doesn't include any files.", vbCritical, "Error"
End
End Sub

