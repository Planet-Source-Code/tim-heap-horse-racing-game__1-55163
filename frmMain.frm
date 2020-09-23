VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horse Race"
   ClientHeight    =   6795
   ClientLeft      =   645
   ClientTop       =   630
   ClientWidth     =   10740
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBackGround 
      AutoSize        =   -1  'True
      Height          =   975
      Left            =   1260
      ScaleHeight     =   915
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame framBetting 
      Caption         =   "Betting"
      Height          =   2775
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   10725
      Begin VB.Frame framButtons 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2655
         Left            =   9240
         TabIndex        =   19
         Top             =   120
         Width           =   1380
         Begin VB.CheckBox chkBackGround 
            Caption         =   "Background"
            Height          =   255
            Left            =   0
            TabIndex        =   25
            ToolTipText     =   "Turning this on makes the game look better, but run slower."
            Top             =   2055
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   375
            Left            =   0
            TabIndex        =   23
            ToolTipText     =   "Quit"
            Top             =   1680
            Width           =   1380
         End
         Begin VB.CommandButton cmdRace 
            Caption         =   "Race"
            Height          =   495
            Left            =   0
            TabIndex        =   22
            ToolTipText     =   "Race!"
            Top             =   120
            Width           =   1380
         End
         Begin VB.CommandButton cmdHelp 
            Caption         =   "Help"
            Height          =   375
            Left            =   0
            TabIndex        =   21
            ToolTipText     =   "Shows the Help"
            Top             =   720
            Width           =   1380
         End
         Begin VB.CommandButton cmdAbout 
            Caption         =   "About"
            Height          =   375
            Left            =   0
            TabIndex        =   20
            ToolTipText     =   "Who made this"
            Top             =   1200
            Width           =   1380
         End
      End
      Begin VB.Frame framWinner 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2655
         Left            =   6930
         TabIndex        =   16
         Top             =   120
         Width           =   2325
         Begin VB.PictureBox picWinner 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   0
            ScaleHeight     =   1275
            ScaleWidth      =   2160
            TabIndex        =   17
            ToolTipText     =   "Who won"
            Top             =   480
            Width           =   2220
            Begin VB.Label lblWinner 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Winner"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   555
               Left            =   315
               TabIndex        =   26
               Top             =   360
               Visible         =   0   'False
               Width           =   1530
            End
         End
         Begin VB.Label lblRaces 
            Caption         =   "Races Left:"
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label lblWinningHorse 
            Alignment       =   2  'Center
            Caption         =   "The Winning Horse is..."
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame framBet 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2655
         Left            =   3360
         TabIndex        =   8
         Top             =   120
         Width           =   3060
         Begin VB.CommandButton cmdCheat 
            Caption         =   "Cheat"
            Height          =   255
            Left            =   1050
            TabIndex        =   27
            Top             =   1800
            Width           =   1275
         End
         Begin VB.TextBox txtBet 
            Height          =   285
            Left            =   1050
            TabIndex        =   10
            Text            =   "10"
            ToolTipText     =   "How much you are betting"
            Top             =   1080
            Width           =   1905
         End
         Begin VB.ComboBox lstHorseNumbers 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   "Which horse will win?"
            Top             =   720
            Width           =   1905
         End
         Begin VB.Label lblMoney 
            Caption         =   "Label1"
            Height          =   255
            Left            =   1080
            TabIndex        =   15
            ToolTipText     =   "How much money you have"
            Top             =   1440
            Width           =   1905
         End
         Begin VB.Label lblHorseNumber 
            Alignment       =   1  'Right Justify
            Caption         =   "Horse Number"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   780
            Width           =   1065
         End
         Begin VB.Label lblPlaceBet 
            Alignment       =   2  'Center
            Caption         =   "Place Your Bet"
            Height          =   315
            Left            =   1065
            TabIndex        =   13
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label lblBet 
            Alignment       =   1  'Right Justify
            Caption         =   "Your Bet:   $"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   1110
            Width           =   1065
         End
         Begin VB.Label lblYourMoney 
            Alignment       =   1  'Right Justify
            Caption         =   "Your Money"
            Height          =   210
            Left            =   30
            TabIndex        =   11
            Top             =   1440
            Width           =   1020
         End
      End
      Begin VB.Frame framOdds 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2655
         Left            =   630
         TabIndex        =   5
         Top             =   120
         Width           =   2640
         Begin VB.Label lblOdds 
            Height          =   255
            Index           =   1
            Left            =   630
            TabIndex        =   7
            ToolTipText     =   "That horses Odds"
            Top             =   0
            Width           =   1275
         End
         Begin VB.Label lblOdd 
            Alignment       =   1  'Right Justify
            Caption         =   "Odds:"
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   540
         End
      End
   End
   Begin VB.PictureBox picFinishline 
      AutoSize        =   -1  'True
      Height          =   975
      Left            =   1260
      ScaleHeight     =   915
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3120
      Top             =   1680
   End
   Begin VB.PictureBox picHorseMask 
      AutoSize        =   -1  'True
      Height          =   1005
      Index           =   0
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   1170
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox picHorseCell 
      AutoSize        =   -1  'True
      Height          =   1005
      Index           =   0
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   1170
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Line lnefinish 
      Visible         =   0   'False
      X1              =   8820
      X2              =   8820
      Y1              =   0
      Y2              =   4080
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Horse Race
'Made by Timothy Heap
'Please E-Mail any comments and suggestions to:
'Immortal_cowpat@hotmail.com
'This code is copyright. Please ask me before using it



Option Explicit

Dim CellCounter As Integer
Dim i As Integer, a As Integer
Dim win As String
Dim winners As String
Dim msg As String
Dim race As Integer
Dim result
Dim background As Boolean

Private Sub chkBackGround_Click()
If chkBackGround.Value = 1 Then
    background = True
    frmBackColour.Visible = False
Else
    background = False
    frmBackColour.Visible = True
    frmBackColour.Top = Me.Top + Me.Height
    frmBackColour.left = framButtons.left + Me.left
End If
End Sub

'Help, Quit and About Buttons
'Makes the other form visible, and locks off this form
'exept for Quit button, which Quits!!

Private Sub cmdAbout_Click()
frmAbout.Show
Me.Enabled = False
End Sub

Private Sub cmdCheat_Click()
frmCheat.Show
Me.Enabled = False
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdHelp_Click()
frmHelp.Show
Me.Enabled = False
End Sub


    'The button that Starts the race
Private Sub cmdRace_Click()
    'Quit if...
    
If lstHorseNumbers.Text = "" Then '... theres no horse selected
    MsgBox "Please select a horse", , "Horse Race"
    Exit Sub
End If
If IsNumeric(txtBet) = False Then '... if the horse number isnt a number?
    MsgBox "Please enter a number in the Betting box", , "Horse Race"
    Exit Sub
End If
If txtBet.Text > Money Then '... if you bet more than you have
    MsgBox "You dont have enough money", , "Horse Race"
    Exit Sub
End If

If txtBet.Text < 0 Then Exit Sub 'if you bet minus money

    'set the winning horse box to the blank horse
picWinner.Cls
picWinner.Picture = picHorseMask(1).Picture

tmrMain.Enabled = True           'start the timer

Money = Money - txtBet          'take off how much you bet from your money
lblMoney.Caption = "$" & Money  'Set the money caption

framBetting.Enabled = False     'lock off the betting frame

For i = 1 To 9                  'Reset the horses, and calculate their speed
    horse(i).left = 0           'put them at the side
                                    'Random number * their odds of winning
    horse(i).speed = Rnd * horse(i).odds + 30
    horse(i).speed = horse(i).speed + horse(i).cheat
    'horse(i).speed = 50
Next i
race = race - 1
lblRaces = "Races Left: " & race
lblWinner.Visible = False
End Sub

Private Sub Form_Load()
HorseYPos = RaceHeight / 9      'Where the horses top is

Me.Visible = True
Randomize
Money = 100                     'Set your Money
race = 10
background = True
lblRaces = "Races Left: " & race
lblMoney.Caption = "$" & Money  'Reset the money caption
ChanceOfBeingCaught = 5
        'Load the grass picture in
picBackGround.Picture = LoadPicture(App.Path & "/Grass.gif")
        'Load the finishline in
picFinishline.Picture = LoadPicture(App.Path & "/finishline.gif")
        'Paint the grass
Me.PaintPicture picBackGround.Picture, 0, 0, Me.ScaleWidth, , , , , , vbSrcCopy
        'Paint the finishline
For i = 0 To RaceHeight / HorseYPos * 2
    Me.PaintPicture picFinishline, lnefinish.X1, i * picFinishline.ScaleHeight - 50, , , , , , , vbSrcCopy
Next i
        'Set the background picture wiht the finishline as well
picBackGround.Picture = Me.Image
        'load the horse picture in
For i = 1 To 4
    Load picHorseCell(i)
    picHorseCell(i).Picture = LoadPicture(App.Path & "/Horse" & i & ".gif")
    Load picHorseMask(i)
    picHorseMask(i).Picture = LoadPicture(App.Path & "/Horse" & i & "Mask.gif")
Next i

        'Load the HorseMask picture in
        'Set the winner's picture to the blank horse
picWinner.Picture = picHorseMask(1).Picture


lstHorseNumbers.Clear           'clear the list of horses
For i = 1 To 9
    If i <> 1 Then
        Load lblOdds(i)         'Load the Odds Lables
        lblOdds(i).Top = lblOdds(i - 1).Top + lblOdds(i - 1).Height
        lblOdds(i).Visible = True
    End If
    horse(i).cell = (Rnd * 3) + 1
        'Paint the HorseMask then the Horse(i) pictures onto the form
    Me.PaintPicture picHorseMask(horse(i).cell).Picture, 0, (i - 1) * HorseYPos, , , , , , , vbMergePaint
    Me.PaintPicture picHorseCell(horse(i).cell).Picture, 0, (i - 1) * HorseYPos, , , , , , , vbSrcAnd
    horse(i).odds = Int(Rnd * 20)   'Set the horses odds
        'Set the lblOdds caption to the horses odds
    lblOdds(i).Caption = "Horse " & i & ":     1:" & 20 - horse(i).odds
        'Add the horse to the list
    lstHorseNumbers.AddItem i & " " & GenerateNames
Next i
    'Positioning the Betting frame
framBetting.Top = (HorseYPos * 10) + (5 * 15)
framBetting.left = -100
framBetting.Width = Me.ScaleWidth + 200
framBetting.Height = lblOdds(9).Top + lblOdds(9).Height + 100
    'Setting the colours of the other frames to normal
framOdds.BackColor = framBetting.BackColor
framBet.BackColor = framBetting.BackColor
framWinner.BackColor = framBetting.BackColor
framButtons.BackColor = framBetting.BackColor
    'Resize the form
Me.Height = RaceHeight + framBetting.Height + 915 - 70
'cell = 1
End Sub



Private Sub Form_Unload(Cancel As Integer)
Unload frmBackColour
Unload frmHelp
Unload frmAbout
End Sub

Private Sub tmrMain_Timer()
win = ""        'Reset win
CellCounter = CellCounter + 1
Me.Cls          'Clear the form
                'Paint on the background
If background = True Then
    Me.PaintPicture picBackGround.Picture, 0, 0, Me.ScaleWidth, , , , , , vbSrcCopy
Else
    For i = 0 To RaceHeight / HorseYPos * 2
        Me.PaintPicture picFinishline, lnefinish.X1, i * picFinishline.ScaleHeight - 50, , , , , , , vbSrcCopy
    Next i
End If
                'Move the horses and draw them

For i = 1 To 9  'Loop 9 times
    horse(i).left = horse(i).left + horse(i).speed  '"Move" the horse
                'Paint the horse on the form
    Me.PaintPicture picHorseMask(horse(i).cell).Picture, horse(i).left, (i - 1) * HorseYPos, , , , , , , vbMergePaint
    Me.PaintPicture picHorseCell(horse(i).cell).Picture, horse(i).left, (i - 1) * HorseYPos, , , , , , , vbSrcAnd
                'If the horse has crossedthe FinishLine
    If horse(i).left + picHorseCell(1).ScaleWidth - 400 > lnefinish.X1 Then
        win = win & i   'Set win to the horses number
    End If
Next i  'Loop for the next horse
If CellCounter > 5 Then
    For i = 1 To 9
        horse(i).cell = horse(i).cell + 1
        If horse(i).cell > 4 Then horse(i).cell = 1
    Next i
    CellCounter = 0
End If
If win <> "" Then   'If a horse has won
    If CDbl(win) < 10 Then  'If its not a draw
                            'If you bet on the right horse
        If win = left$(lstHorseNumbers.Text, 1) Then
                            'Give you your bet + your bet * the horses odds
            Money = Money + CInt(txtBet) + (CInt(txtBet) * (20 - horse(win).odds))
        End If
        lblMoney.Caption = "$" & Money  'Update money
        MsgBox "Horse " & win & " Won!", , "Horse Race" 'Message box with which horse won
            'Paint the winning horse straight from the form into the winning box
        picWinner.PaintPicture Me.Image, 0, 0, , , horse(win).left, (HorseYPos * (win - 1)), picHorseCell(1).ScaleWidth, picHorseCell(1).ScaleHeight, vbSrcCopy
        lblWinner.Caption = win
    Else            'If it was a draw...
        'Searches 'win' to se if you bet on the right horse
        If InStr(1, win, lstHorseNumbers.Text, vbTextCompare) Then
            'Give you your money / how many horses won
            Money = Money + CInt(txtBet) + ((CInt(txtBet) * horse(lstHorseNumbers.Text).odds) / Len(win))
            lblMoney.Caption = "$" & Money 'Update the money
        End If
        picWinner.Cls   'Clear PicWinner
        picWinner.Picture = picHorseMask(1).Picture 'PicWinners pic to the blank horse
        lblWinner.Caption = "Draw"
        
        winners = ""    'Reset Winners
        For i = 1 To Len(win)
            'sets 'winners' to win, but with commas between the numbers
            winners = winners & Mid$(win, i, 1) & ", "
        Next i
            'get rid of the last comma
        winners = left$(winners, Len(winners) - 2)
            'Put and between the last two numbers
        winners = left$(winners, Len(winners) - 2) & " and " & Right$(winners, 1)
            'Writes msg with 'winners' in the middle
        msg = "It was a draw between Horses " & winners & "."
            'message box saying who drew
        MsgBox msg, , "Horse Race"
    End If
    For i = 1 To 9
        horse(i).cheat = 0
        horse(i).odds = Int(Rnd * 20)   'Set the horses odds
            'Set the lblOdds caption to the horses odds
        lblOdds(i).Caption = "Horse " & i & ":     1:" & 20 - horse(i).odds

    Next i
    lblWinner.Visible = True
    lblWinner.left = picWinner.left + (picWinner.Width / 2) - (lblWinner.Width / 2)
    tmrMain.Enabled = False      'stop timer1
    framBetting.Enabled = True  'Enable the Betting frame again
    If race <= 0 Then
        msg = "You made $" & Money - 100 & ". Congratulations. Play again?"
        result = MsgBox(msg, vbYesNo, "Horse Race")
        If result = vbYes Then
            Call ResetStats
        Else
            End
        End If
    End If
End If


End Sub

Sub ResetStats()
Money = 100                     'Set your Money
lblMoney.Caption = "$" & Money  'Reset the money caption
race = 10
lblRaces = "Races Left: " & race
Me.Cls
For i = 1 To 9
    Me.PaintPicture picHorseMask(horse(i).cell).Picture, 0, (i - 1) * HorseYPos, , , , , , , vbMergePaint
    Me.PaintPicture picHorseCell(horse(i).cell).Picture, 0, (i - 1) * HorseYPos, , , , , , , vbSrcAnd
    horse(i).odds = Int(Rnd * 20)   'Set the horses odds
        'Set the lblOdds caption to the horses odds
    lblOdds(i).Caption = "Horse " & i & ":     1:" & 20 - horse(i).odds
Next i
ChanceOfBeingCaught = 5
End Sub
