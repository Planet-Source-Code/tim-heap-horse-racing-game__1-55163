VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3135
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Back to the Game"
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   " Immortal_Cowpat@Hotmail.com"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   3165
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "© 2004 Heap Productons."
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Made by Timothy Heap."
      Height          =   255
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Thats Me!!!"
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Horse Race"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
frmMain.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = frmMain.Top + (frmMain.Height / 2) - (Me.Height / 2)
Me.left = frmMain.left + (frmMain.Width / 2) - (Me.Width / 2)

If Me.Top < 0 Then Me.Top = 0
If Me.left < 0 Then Me.left = 0
If Me.Top + Me.Height > Screen.Height Then Me.Top = Screen.Height - Me.Height
If Me.left + Me.Width > Screen.Width Then Me.left = Screen.Width - Me.Width

End Sub
