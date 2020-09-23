VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Win-CAB [Windows Cabinet File Creation Utility]"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2160
      Top             =   2760
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2640
      Top             =   2760
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3720
      Top             =   2760
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3240
      Top             =   2760
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "Default"
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog MergeOutput 
      Left            =   4680
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin MSComDlg.CommonDialog AddFolder 
      Left            =   4680
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add File"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog AddFile 
      Left            =   4680
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.ListBox lstFiles 
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Status 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   675
   End
   Begin VB.Line Line1 
      X1              =   4440
      X2              =   120
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "CAB Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function FileExists(strFile As String) As String
'I looked at a lot of different examples of
'These, this turned out the best to see if
'My user settings were there
On Error Resume Next
FileExists = Dir(strFile, vbHidden) <> ""
End Function
Private Sub CheckDaWin()

Timer1.Enabled = False
End Sub
Private Sub Command1_Click()
    AddFile.filename = "" 'To clear the filepath
        AddFile.ShowOpen 'Shows the Open common dialog
        If AddFile.filename = "" Then Exit Sub 'If no file is selected or user clicks cancel, end it
    lstFiles.AddItem AddFile.filename 'Adds the selected file's path from the common dialog
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
Status.Caption = "Creating Diamond Directive File..."
If lstFiles.ListCount = 0 Then Exit Sub 'If no files are in the listbox, end sub
Dim x As Integer
On Error GoTo endit 'End of the sub
 Open App.Path & "\Direct.ddf" For Output As #1 ' Open the path to write to it

        Print #1, ".Option Explicit"
        Print #1, ".Set Cabinet=on"
        Print #1, ".set Compress=on"
        Print #1, ".Set MaxDiskSize=CDRom"
        Print #1, ".set ReservePerCabinetSize=6144"
        Print #1, ".Set DiskDirectoryTemplate="
        Print #1, ".Set CompressionType=MSZip"
        Print #1, ".Set CompressionLevel=7"
        Print #1, ".Set CompressionMemory=21"
        Print #1, ".Set CabinetNameTemplate=" & Text1 & ".cab"
Dim i
For i = 0 To 100
    lstFiles.ListIndex = i
    Write #1, lstFiles.Text
Next i
             'Close the DOS window after it finishes
            'This generates a batch file to merge all the files fromthe lsitbox into one file
            'These are DOS commands that will merge the files from the listbox in the Batch file
Close #1
Timer1.Enabled = True
    '********
    'Shell App.Path & "\MergeFiles.bat", vbMinimizedNoFocus
    '********
    'Open the batch file and it will then minimize and merge.
    'When it finishes, it will close automatically
endit:
Close #1
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
lstFiles.Clear
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Timer1_Timer()
Status.Caption = "Creating\Running Batch Directive..."
Dim x As String
x = App.Path
Open App.Path & "\Create.bat" For Output As #1
Print #1, "@Echo Off"
Print #1, "cd\"
Print #1, "cd " & x
Print #1, "Makecab /f Direct.ddf"
Print #1, "Del Setup.inf"
Print #1, "Del Setup.rpt"
Print #1, "Del Direct.ddf"
Print #1, "@CLS"
Print #1, "@CLS"
Close #1
Shell App.Path & "\Create.bat", vbHide
Timer1.Enabled = False
Timer2.Enabled = True
Timer3.Enabled = True
End Sub

Private Sub Timer2_Timer()
Kill App.Path & "\Create.bat"
Timer2.Enabled = False
lstFiles.Clear
Command2.Enabled = True
Timer4.Enabled = True
End Sub

Private Sub Timer3_Timer()
Status.Caption = "Deleting Un-needed directives..."
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Status.Caption = "Done!"
Timer4.Enabled = False
End Sub
