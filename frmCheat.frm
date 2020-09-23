VERSION 5.00
Begin VB.Form frmCheat 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCheat 
      Caption         =   "Saboutage"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   960
   End
   Begin VB.CommandButton cmdCheat 
      Caption         =   "Drug"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   960
   End
   Begin VB.CommandButton cndCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   960
   End
   Begin VB.ComboBox lstHorseNumbers 
      Height          =   315
      Left            =   0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Which horse will win?"
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "frmCheat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCheat_Click(Index As Integer)
If lstHorseNumbers.Text = "" Then '... theres no horse selected
    MsgBox "Please select a horse", , "Horse Race"
    Exit Sub
End If
If IsNumeric(lstHorseNumbers.Text) = False Then '... if the horse number isnt a number?
    MsgBox "Please enter a number in the Betting box", , "Horse Race"
    Exit Sub
End If
If 100 > Money Then '... if you bet more than you have
    MsgBox "You dont have enough money", , "Horse Race"
    Exit Sub
End If
Money = Money - 100
frmMain.lblMoney = "$" & Money
mynum = Int(Rnd * ChanceOfBeingCaught)
ChanceOfBeingCaught = ChanceOfBeingCaught - 1
If mynum = 2 Then
    MsgBox "You were found out!! You were forced to pay a fine of $100,000, and go to jail for a year. You loose"
    End
End If
If Index = 0 Then
    horse(lstHorseNumbers.Text).cheat = 10
Else
    horse(lstHorseNumbers.Text).cheat = -20
End If
frmMain.Enabled = True
Unload Me
End Sub

Private Sub cndCancel_Click()
frmMain.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
lstHorseNumbers.Clear
For i = 1 To 9
    lstHorseNumbers.AddItem (i)
    horse(i).cheat = 0
Next i
MsgBox "The chance of being caught is 1:" & ChanceOfBeingCaught & "."
End Sub
