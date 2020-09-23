VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2295
   ClientLeft      =   2115
   ClientTop       =   1815
   ClientWidth     =   3825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   3795
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Good Luck."
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   3795
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmHelp.frx":0000
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3795
   End
End
Attribute VB_Name = "frmHelp"
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
Command1.Width = Me.ScaleWidth
Command1.left = 0
End Sub
