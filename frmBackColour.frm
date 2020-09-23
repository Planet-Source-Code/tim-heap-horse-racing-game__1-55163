VERSION 5.00
Begin VB.Form frmBackColour 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   735
   ClientLeft      =   1230
   ClientTop       =   2265
   ClientWidth     =   1335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   1335
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.HScrollBar hscBlue 
      Height          =   255
      LargeChange     =   20
      Left            =   0
      Max             =   255
      SmallChange     =   5
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.HScrollBar hscGreen 
      Height          =   255
      LargeChange     =   20
      Left            =   0
      Max             =   255
      SmallChange     =   5
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.HScrollBar hscRed 
      Height          =   255
      LargeChange     =   20
      Left            =   0
      Max             =   255
      SmallChange     =   5
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmBackColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub hscRed_Change()
frmMain.BackColor = GetFormColour
PaintBackground
End Sub

Private Sub hscGreen_Change()
frmMain.BackColor = GetFormColour
PaintBackground
End Sub

Private Sub hscBlue_Change()
frmMain.BackColor = GetFormColour
PaintBackground
End Sub

Function GetFormColour()
GetFormColour = RGB(hscRed.Value, hscGreen.Value, hscBlue.Value)
End Function

Sub PaintBackground()
For i = 0 To 20 * 2
    frmMain.PaintPicture frmMain.picFinishline, frmMain.lnefinish.X1, i * frmMain.picFinishline.ScaleHeight - 50, , , , , , , vbSrcCopy
Next i

For i = 1 To 9  'Loop 9 times
                'Paint the horse on the form
    frmMain.PaintPicture frmMain.picHorseMask(horse(i).cell).Picture, horse(i).left, (i - 1) * HorseYPos, , , , , , , vbMergePaint
    frmMain.PaintPicture frmMain.picHorseCell(horse(i).cell).Picture, horse(i).left, (i - 1) * HorseYPos, , , , , , , vbSrcAnd

Next i  'Loop for the next horse
End Sub

