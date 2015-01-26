VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Forecast File Direction"
   ClientHeight    =   4995
   ClientLeft      =   6690
   ClientTop       =   3000
   ClientWidth     =   7350
   LinkTopic       =   "Form4"
   ScaleHeight     =   4995
   ScaleWidth      =   7350
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fcst_All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3000
         Width           =   6135
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Left            =   2280
         TabIndex        =   5
         Top             =   480
         Width           =   4095
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
Text1.Text = Dir1.Path
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Text1.Text = File1.Path
End Sub

Private Sub file1_click()
Text1.Text = File1.Path + "\" + File1.FileName
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "Please choose a file.", vbOKOnly, "Error"
Else
    Form1.fcst_path = Text1.Text
    Form5.Show
    Unload Me
End If
End Sub
