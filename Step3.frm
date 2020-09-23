VERSION 5.00
Begin VB.Form Step3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Win-CAB [Windows Cabinet File Creation Utility]"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Step 3"
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
      TabIndex        =   4
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command2 
         Caption         =   "Create"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Status 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   840
         TabIndex        =   7
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Create the .cab File"
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
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Done"
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2520
      Width           =   975
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4320
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4800
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   9000
      Left            =   5880
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   5400
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "Default"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add File"
      Height          =   255
      Left            =   7680
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstFiles 
      Height          =   1230
      Left            =   1560
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Win-CAB by SoTe (Thee_SoTe@Hotmail.com)"
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
      TabIndex        =   9
      Top             =   2520
      Width           =   3255
   End
End
Attribute VB_Name = "Step3"
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

Timer1.Enabled = False
End Sub
Private Sub Command2_Click()
Command2.Enabled = False
Command4.Enabled = False
Status.Caption = "Creating Diamond Directive File..."
If lstFiles.ListCount = 0 Then Exit Sub 'If no files are in the listbox, end sub
Dim X As Integer
On Error GoTo endit 'End of the sub
 Open App.Path & "\Direct.ddf" For Output As #1 ' Open the path to write to it

 'THIS is the directive file. The main purpose of this program.
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
    'Put the file names/paths in the directive file...
Next i
Close #1
Timer1.Enabled = True
endit:
Close #1
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
Unload Step1
Unload Step2
Unload Step3
End
End Sub

Private Sub Form_Load()
Dim i
Dim X
On Error GoTo error
Open App.Path & "\Break.dll" For Input As #1


For i = 0 To 100
    Input #1, X
    lstFiles.AddItem X
Next i
error:
Close #1
Kill App.Path & "\Break.dll"
Unload Step1
Unload Step2
End Sub

Private Sub Timer1_Timer()
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
Command4.Enabled = True
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
