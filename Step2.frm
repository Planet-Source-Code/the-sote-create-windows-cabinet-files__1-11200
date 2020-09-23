VERSION 5.00
Begin VB.Form Step2 
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
   Begin VB.ListBox lstFiles 
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Step 2"
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
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Text            =   "CabFile1"
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   ".cab File Name:"
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
         TabIndex        =   4
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select a name for the finished CAB file."
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
         TabIndex        =   3
         Top             =   360
         Width           =   2805
      End
   End
   Begin VB.CommandButton Command4 
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
End
Attribute VB_Name = "Step2"
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



Private Sub Command4_Click()

Step3.Show
Step3.Text1.Text = Step2.Text1.Text
Unload Me
End Sub

Private Sub Form_Load()
Dim i
Dim X
On Error GoTo error
'Get info from the listbox file
Open App.Path & "\Break.dll" For Input As #1


For i = 0 To 100
    Input #1, X
    lstFiles.AddItem X
Next i
error:
Close #1
End Sub

