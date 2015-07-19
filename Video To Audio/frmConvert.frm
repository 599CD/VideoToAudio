VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConvert 
   Caption         =   "Convert"
   ClientHeight    =   2190
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7470
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6720
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmAudio 
      Caption         =   "Audio"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   6855
      Begin VB.CommandButton cmdOpenFolder 
         Height          =   375
         Left            =   6360
         Picture         =   "frmConvert.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdConvert 
         Caption         =   "Convert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtFilePath_A 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   5295
      End
      Begin VB.TextBox txtFileName_A 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblFilePath_A 
         Caption         =   "File Path"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblFileExt_A 
         Caption         =   ".mp3"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblFileName_A 
         Caption         =   "File Name"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame frmVideo 
      Caption         =   "Video"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton cmdChoose 
         Caption         =   "Choose"
         Height          =   255
         Left            =   6360
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtFilePath_V 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   5295
      End
      Begin VB.TextBox txtFileName_V 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblFilePath_V 
         Caption         =   "File Path"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblFileName_V 
         Caption         =   "File Name"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblDot 
         Caption         =   "."
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   600
         Width           =   135
      End
      Begin VB.Label lblFileExt_V 
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChoose_Click()
    'FileDialog
    
    CommonDialog.Filter = "MP4 (*.mp4)|*.mp4|MOV (*.mov)|*.mov|M4V (*.m4v)|*.m4v|All files (*.*)|*.*"
    CommonDialog.DefaultExt = "txt"
    CommonDialog.DialogTitle = "Select a Video File"
    CommonDialog.ShowOpen
    
    txtFilePath_V = CommonDialog.FileName
    
    If Len(txtFilePath_V) <> 0 Then
        
        If FileExists(txtFilePath_V) = False Then
            MsgBox "No such File exists", vbOKOnly + vbInformation, "No File Exists"
            Exit Sub
        End If
        
        cmdConvert.Visible = True
        cmdOpenFolder.Visible = True
        
        Dim strFile() As String
        strFile() = GetFileName(txtFilePath_V)
        
        txtFileName_V = strFile(1)
        lblFileExt_V = strFile(2)
        
        txtFileName_A = strFile(1)
        txtFilePath_A = App.Path & "\" & txtFileName_A & ".mp3"
    End If
    
End Sub

Sub ClearChoice()
    txtFilePath_V = ""
    txtFileName_V = ""
    lblFileExt_V = ""
    txtFilePath_A = ""
    txtFileName_A = ""
End Sub

Sub Convert()
    
    'BuildFile BuildConversion
    ' now run the batchfile
    'Shell "RUN.BAT", vbNormalFocus
    
    Shell "cmd /c " & BuildConversion, vbNormalFocus

End Sub

Function BuildConversion() As String
    '_ffmpeg -i ShesNotThere.mp4 -q:a 0 -map a ShesNotThere.mp3
    Dim strArg As String
    Dim strVideo As String
    Dim strAudio As String
    Dim strCurrentPath As String

    strCurrentPath = App.Path & "\"
    'strVideo = txtFileName_V & "." & lblFileExt_V
    strVideo = AddQuotes(txtFilePath_V)
    'strAudio = txtFileName_A & lblFileExt_A
    strAudio = AddQuotes(txtFilePath_A)
    
    strArg = "_ffmpeg -i " & strVideo & " -q:a 0 -map a " & strAudio
    
    BuildConversion = strArg
    
End Function

Sub BuildFile(strArg As String)
    
    '_ffmpeg -i ShesNotThere.mp4 -q:a 0 -map a ShesNotThere.mp3
    
    Dim FF
    FF = FreeFile
    Open "RUN.BAT" For Output As #FF
    
    'Print #FF, "_ffmpeg -y -i " & myAVI & " -b:v 2000k -r 30 " & myMPG
    Print #FF, strArg
    'Print #FF, "del output.mpg"
    Print #FF, "del run.bat"
        
    Close #FF
    
End Sub

Private Sub cmdConvert_Click()
        
    If ProgramsExist = False Then
        MsgBox "FFMPEG doesn't exist. Please add this to the folder.", vbOKOnly + vbInformation, "Missing App"
        Exit Sub
    End If
    
    Convert
    MsgBox "Done", vbOKOnly + vbInformation, "Conversion Completed"
    ClearChoice
    cmdConvert.Visible = False
End Sub

Private Sub cmdOpenFolder_Click()
    OpenFolder App.Path
End Sub
