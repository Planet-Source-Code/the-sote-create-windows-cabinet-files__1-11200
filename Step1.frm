VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Step1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Win-CAB [Windows Cabinet File Creation Utility]"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   Icon            =   "Step1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Huh?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Step 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      Begin VB.ListBox lstFiles 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         ItemData        =   "Step1.frx":0442
         Left            =   120
         List            =   "Step1.frx":0444
         TabIndex        =   4
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Add all the files you would like to be in the CAB file"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2520
      Width           =   975
   End
   Begin MSComDlg.CommonDialog MergeOutput 
      Left            =   4680
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog AddFolder 
      Left            =   4680
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog AddFile 
      Left            =   4680
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Step1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Win-CAB
'by SoTe (Thee_SoTe@Hotmail.com)
'*******************************
'You can use this code freely as long as
'you don't claim it to be your own. This program
'was made so that .cab files can be made easy.
'This program makes the directive files (which is
'needed for multiple files in a .cab file) and then
'runs a program and uses the directive file. Yea I
'know it kinda sucks because I need to use an external
'program to acutally MAKE the .cab file, but that is
'soon to come.

Public Function FileExists(strFile As String) As String

End Function
Private Sub CheckDaWin()
'This was originally something else... But
'changed at the last miniute.
Timer1.Enabled = False
End Sub
Private Sub Command1_Click()
    AddFile.filename = "" 'To clear the filepath
        AddFile.ShowOpen 'Shows the Open common dialog
        If AddFile.filename = "" Then Exit Sub 'If no file is selected or user clicks cancel, end it
    lstFiles.AddItem AddFile.filename 'Adds the selected file's path from the common dialog
End Sub

Private Sub Command2_Click()
Dim i
On Error GoTo error
'Stores the listbox data temporarily
Open App.Path & "\Break.dll" For Output As #1
For i = 0 To 100
    lstFiles.ListIndex = i
    Print #1, lstFiles.Text
Next i
error:
Close #1
Step2.Show
Unload Me
End Sub

Private Sub Command3_Click()
lstFiles.Clear
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Command4_Click()
Me.Hide
MsgBox "As you can see, included with this program is 'Makecab.exe'. This is actual program that makes the .cab files. If you tried to use file (Makecab.exe) alone, you could'nt do multiple files. But if a directive file is included then you can include multiple files. So this program is mainly made to create the directive file for 'Makecab.exe'. This is done because most people don't know how to make a directive file. That is why this simple program is made.", vbInformation, "But why?"
Me.Show
End Sub

Private Sub Timer1_Timer()
'Yep, thats right I used timers. I needed them
'because I had to time it just right to make/delete files.
Status.Caption = "Creating\Running Batch Directive..."
Dim X As String
X = App.Path
Open App.Path & "\Create.bat" For Output As #1
Print #1, "@Echo Off"
Print #1, "cd\"
Print #1, "cd " & X
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

