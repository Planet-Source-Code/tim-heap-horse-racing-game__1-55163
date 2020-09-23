Attribute VB_Name = "Module1"
Public Type horses
    left As Integer
    speed As Single
    odds As Integer
    cell As Integer
    cheat As Integer
End Type

Public horse(9) As horses
Public Money As Integer
Public Const RaceHeight As Integer = 4000
Public HorseYPos As Integer
Public ChanceOfBeingCaught



    
Function GenerateNames()
MyNumRoll:
mynum = CInt(Rnd * 10)
If mynum < 3 Then GoTo MyNumRoll
For i = 0 To mynum          'Makes 'MyNum' letters in the word
lowercase:                      'Can return to here to make another letter
        mynum = Int((122 - 97 + 1) * Rnd + 97)  'Get a random number between
        MyChar = Chr(mynum)                         '97 and 122 and uses it in
        If i <> 0 Then
            Select Case MyChar                          'the Chr(character) command
                Case "a", "e", "i", "o", "u"     'If the letter is a vowel...
                    If vowel = True Then           'If letter before was a vowel...
                        GoTo lowercase              'Start again (make new letter)
                    Else                           'If not...
                        vowel = True                'Set vowel to true
                    End If                         'End If
                Case Else                        'Not a Vowel...
                    If vowel = False Then          'If letter before not a vowel...
                        GoTo lowercase              'Start again (make new letter)
                    Else                           'If it was...
                        vowel = False               'Set vowel to true
                    End If                         'End if
            End Select                           'End Select
        Else
            Select Case MyChar
                Case "a", "e", "i", "o", "u"     'If the letter is a vowel...
                    vowel = True
            End Select
        End If
        HorseName = HorseName & MyChar    'Add letter to Label
        Select Case MyChar
            Case "q"                                 'If letter is a q...
                HorseName = HorseName & "u"       'add a u
                                                     'Hey it rhymes!
            Case "s", "c", "t" 'If s,c or t
                mynum = Rnd * 4             'give random number between 0 and 3
                If mynum = 2 Then               '(one in three chance)
                    HorseName = HorseName & "h"   'Add a h
                End If
        
            Case "g", "r", "p" 'same as above
                mynum = Rnd * 7             'give random number between 0 and 6
                If mynum = 2 Then               '(one in 6 chance)
                    HorseName = HorseName & "h"   'Add a h
                End If
            
            Case "e", "a" 'same as above
                mynum = Rnd * 7             'give random number between 0 and 6
                If mynum = 2 Then               '(one in 6 chance)
                    HorseName = HorseName & "e"   'Add an e
                End If
        End Select
        mynum = Rnd * 11
        If mynum = 2 Then
            HorseName = HorseName & MyChar
        End If
    Next i              'Next letter

    If accent = True Then
        mynum = Rnd * 2     'Random number
        If mynum = 1 Then   '(one chance in two)
        
weirdletter:

            mynum = Int((255 - 192 + 1) * Rnd + 192)    'Random accented letter
            Select Case mynum
                Case 215, 222, 247, 223     'dont want letters 215,222,247 & 223
                    GoTo weirdletter        'Make another accented letter
            End Select
            MyChar = Chr(mynum)             'Set MyChar to Chr(MyNum)
            
            'This bit gets a random number from the length of lblWord(a)
                'and inserts the MyChar letter into that position in lblWord(a)
            mynum = Rnd * Len(HorseName)
            namebitleft = left$(HorseName, mynum) & MyChar '
            namebitright = Right$(HorseName, Len(HorseName) - mynum)
            HorseName = namebitleft & namebitright 'Puts the lblWord(a)
        End If                                                  'back together
    End If
                'Make the word all lowercase, the make the first letter Uppercase

    HorseName = Format(left$(HorseName, 1), ">") & Format(Right$(HorseName, Len(HorseName) - 1), "<")

        'This makes sure the word does not exceede 10 letters long.
    HorseName = left$(HorseName, 10)
    GenerateNames = HorseName
End Function
