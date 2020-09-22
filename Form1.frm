VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CrossWord Puzzle"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   631
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkIncludeInverted 
      Caption         =   "Include inverted"
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CheckBox chkSolutions 
      Caption         =   "Show solutions"
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Height          =   495
      Left            =   8520
      TabIndex        =   6
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtSize 
      Height          =   285
      Left            =   7440
      TabIndex        =   5
      Text            =   "10"
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   6840
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtfield 
      Height          =   285
      Left            =   6840
      TabIndex        =   2
      Top             =   3240
      Width           =   2535
   End
   Begin VB.ListBox lstitems 
      Height          =   2985
      Left            =   6840
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.PictureBox picfield 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6750
      Left            =   0
      ScaleHeight     =   446
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   446
      TabIndex        =   0
      Top             =   0
      Width           =   6750
   End
   Begin VB.Label Label1 
      Caption         =   "Size:"
      Height          =   255
      Left            =   6840
      TabIndex        =   7
      Top             =   4080
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Holds the ascii for each cell
Dim field() As Byte

'Holds if the cell is used by a word, and if
'it's used, with what type (see constants below)
Dim used() As Byte

'Array to hold all solutions
Dim solutions() As solution

'Solution consists of starting cell and ending cell (2 points)
Private Type solution
    x1 As Integer
    y1 As Integer
    x2 As Integer
    y2 As Integer
End Type

'Constants for direction to place a word
Const VERT = 1
Const HORZ = 2
Const NONE = 0
Const DIAG1 = 3
Const DIAG2 = 4

Private Sub chkSolutions_Click()
    If chkSolutions = Checked Then
        'draw solution lines
        Call DrawSolution
    Else
        picfield.Cls
        DrawGrid
        DrawContents
    End If
    
End Sub

Private Sub cmdAdd_Click()
    'if field doesn't contain anything, don't add
    If txtfield.Text = "" Then
        Exit Sub
    End If
    
    'word have to be at least 2 letters
    If Len(txtfield.Text) < 2 Then
        Exit Sub
    End If
    
    'check if the entry is already listed
    Dim val As Integer
    val = IsInList(txtfield.Text)
    
    'not listed, then add
    If val <> -1 Then
        txtfield.Text = ""
        Exit Sub
    Else
        lstitems.AddItem txtfield.Text
    End If
    
    txtfield.Text = ""
End Sub

Function IsInList(str As String) As Integer
    'Check if the str is already listed in the lstitems
    Dim i As Integer
    For i = 0 To lstitems.ListCount - 1
        'not case sensitive
        If LCase(str) = LCase(lstitems.List(i)) Then
            IsInList = i
            Exit Function
        End If
    Next
    
    IsInList = -1
End Function

Private Sub cmdDraw_Click()

    'redim the field, used to the specified size
    ReDim field(val(txtSize.Text), val(txtSize.Text))
    ReDim used(val(txtSize.Text), val(txtSize.Text))
    ReDim solutions(0) As solution
    
    
    'fill field with spaces
    Dim i As Integer
    Dim j As Integer
    For j = 0 To val(txtSize.Text)
        For i = 0 To val(txtSize.Text)
            field(i, j) = 32
        Next
    Next
    
    'build the puzzle
    CreatePuzzle
    
    FillRemWithRandom
    
    picfield.Cls
    DrawGrid
    DrawContents
    If chkSolutions.Value = Checked Then
        DrawSolution
    End If
End Sub

Sub DrawGrid()
    'Draws a vertical line every n pixels
    'where n = total width / specified size
    Dim i As Integer
    For i = 1 To val(txtSize.Text)
        'draw vertical line
        picfield.Line ((picfield.Width \ val(txtSize.Text)) * i, 0)-((picfield.Width \ val(txtSize.Text)) * i, picfield.Height)
        'draw horizontal line
        picfield.Line (0, (picfield.Height \ val(txtSize.Text)) * i)-(picfield.Width, (picfield.Height \ val(txtSize.Text)) * i)
    Next
    picfield.Refresh
    
End Sub

Sub DrawContents()
    
    'change the fontsize if the size gets bigger
    
    Dim fontsize As Integer
    fontsize = 10 + (10 - val(txtSize.Text)) * 2
    If fontsize < 6 Then
        fontsize = 6
    End If
    
    'set fontsize on picturebox
    picfield.fontsize = fontsize
    
    
    'now print the letters from the field on the picturebox
    Dim i As Integer
    Dim j As Integer
    For j = 0 To val(txtSize.Text)
        For i = 0 To val(txtSize.Text)
            ' currentx = cell.left + cell.width / 2 - characterwidth / 2
            ' currenty = cell.top + cell.height / 2 - characterheight /2
            picfield.CurrentX = i * (picfield.Width \ val(txtSize.Text)) + ((picfield.Width \ val(txtSize.Text)) \ 2) - (picfield.TextWidth(Chr(field(i, j))) \ 2)
            picfield.CurrentY = j * (picfield.Height \ val(txtSize.Text)) + ((picfield.Width \ val(txtSize.Text)) \ 2) - (picfield.TextHeight(Chr(field(i, j))) \ 2)
            picfield.Print Chr(field(i, j))
        Next
    Next
    
End Sub

Private Sub cmdRemove_Click()
    'Remove an entry from the list
    If lstitems.ListIndex >= 0 And lstitems.ListIndex < lstitems.ListCount Then
        lstitems.RemoveItem (lstitems.ListIndex)
    End If
End Sub

Private Sub Form_Load()
    ReDim field(val(txtSize.Text), val(txtSize.Text))
    ReDim used(val(txtSize.Text), val(txtSize.Text))
    ReDim solutions(0)
    
End Sub

Private Sub lstitems_Click()
    txtfield.Text = lstitems.List(lstitems.ListIndex)
End Sub

Sub CreatePuzzle()
    Dim i As Integer
    
    'check if words are too long
    Dim res As Integer
    res = wordsTooLong
    If res <> -1 Then
        'notify user that a word is too long
        MsgBox "Word: '" & lstitems.List(res) & "' is too long, increase size or remove word"
        'set index to that word
        lstitems.ListIndex = res
        Exit Sub
    End If
    
    'for each word
    For i = 0 To lstitems.ListCount - 1
        ' pick a random cell
        Dim xpos As Integer
        Dim ypos As Integer
        xpos = Int(Rnd() * val(txtSize.Text))
        ypos = Int(Rnd() * val(txtSize.Text))
        
        'if inverted words are included, pick inverted words random
        Dim invert As Integer
        If chkIncludeInverted = Checked Then
            invert = Int(Rnd() * 2)
        Else
            invert = 0
        End If
        
        'if we need to invert the word, flip the string
        Dim curstr As String
        If invert = 1 Then
            curstr = flipString(lstitems.List(i))
        Else
            curstr = lstitems.List(i)
        End If
        
        
        'pick a random direction
        Dim testdir As Integer
        testdir = Int(Rnd() * 4) + 1
        
        'check if direction is possible
        Dim v As Boolean
        v = isDirPossible(testdir, curstr, xpos, ypos)
        
        'Holds the number of tries to place a word
        Dim count As Long
        count = 0
        Do While Not v
            'pick a random cell
            xpos = Int(Rnd() * val(txtSize.Text))
            ypos = Int(Rnd() * val(txtSize.Text))
            
            'try all directions, if a correct one is found, keep it
            Dim dirdone(4) As Boolean
            Erase dirdone
            
            Do While Not dirdone(1) Or Not dirdone(2) Or Not dirdone(3) Or Not dirdone(4)
                'test all directions, keep picking random directions until all have been picked
                testdir = Int(Rnd() * 4) + 1
                v = isDirPossible(testdir, curstr, xpos, ypos)
                dirdone(testdir) = True
                
                'if the current direction is possible, stop the loop
                If v Then Exit Do
            Loop
            DoEvents
            count = count + 1
            
            'unable to find a correct placement after 30000 tries
            If count > 30000 Then
                MsgBox "After 30000 retries still no possible move on " & curstr & " :s"
                Exit Sub
            End If
        Loop
        'Debug.Print "Word: " & curstr & " set after " & count & " retries"
        
        'put the current word on the cell with the found direction
        Call PutWord(curstr, xpos, ypos, testdir)
    Next
End Sub

Function flipString(str As String) As String
    'Flip a string backwards, e.g "hello" becomes "olleh"
    Dim ret As String
    Dim i As Integer
    For i = Len(str) To 1 Step -1
        ret = ret & Mid$(str, i, 1)
    Next
    flipString = ret
    
End Function

Function wordsTooLong() As Integer
    'Checks if words are too long
    Dim i As Integer
    For i = 0 To lstitems.ListCount - 1
        If Len(lstitems.List(i)) > val(txtSize.Text) Then
            wordsTooLong = i
            Exit Function
        End If
    Next
    
    wordsTooLong = -1
End Function

Function isDirPossible(dir As Integer, str As String, xpos As Integer, ypos As Integer) As Boolean
    Dim ispossible As Boolean
    
    'First test: check if there is enough space to place the words
    '----------
    
    If val(txtSize.Text) - xpos >= Len(str) And _
       val(txtSize.Text) - ypos >= Len(str) Then
        'DIAG1, VERT or HORZ are possible, enough space right and below
        If dir = HORZ Or dir = VERT Or dir = DIAG1 Then
            ispossible = True
        ElseIf dir = DIAG2 Then
            'diag2 is '/' so check if we have enough space above the given cell coordinates
            ispossible = (ypos - (Len(str) - 1) >= 0)
        End If
    ElseIf val(txtSize.Text) - xpos >= Len(str) And _
           ypos - (Len(str) - 1) >= 0 Then
            'enough space right and enough space above
        If dir = HORZ Or dir = DIAG2 Then
            ispossible = True
        ElseIf dir = VERT Or dir = DIAG1 Then
            'check if there is enough space below
            ispossible = (val(txtSize.Text) - ypos >= Len(str))
        End If
    ElseIf val(txtSize.Text) - xpos >= Len(str) Then
            'enough space right only
        If dir = HORZ Then
            ispossible = True
        Else
            ispossible = False
        End If
    ElseIf val(txtSize.Text) - ypos >= Len(str) Then
            'enough space below only
        If dir = VERT Then
            ispossible = True
        Else
            ispossible = False
        End If
    Else
        'no space right and below
        ispossible = False
    End If
    
    'if first test isn't possible return false and exit function
    If Not ispossible Then
        isDirPossible = False
        Exit Function
    End If
    
    'Second Test: Check if we place it, that we don't overwrite a cell
    'that is already being used be another word (unless the cell's letter is the same)
    'also encountered other words can't be in the same direction
    Dim i As Integer
    Dim count As Integer
    If dir = HORZ Then
        
        For i = xpos To xpos + Len(str) - 1
        
            If used(i, ypos) <> NONE Then
                'if we encounter a word thats placed HORZ, while we also want to put a word HORZ, add count +1
                If used(i, ypos) = HORZ Then
                    count = count + 1
                End If

                
                If Asc(UCase(Mid$(str, i - xpos + 1, 1))) <> field(i, ypos) Then
                    ' no match with char from other word, so it isn't possible
                    isDirPossible = False
                    Exit Function
                End If
                        
            End If
        Next
        'if we encountered multiple words that are placed HORZ, then placing it isn't possible
        If count > 0 Then
            isDirPossible = False
            Exit Function
        End If
    ElseIf dir = VERT Then
        For i = ypos To ypos + Len(str) - 1
            If used(xpos, i) <> NONE Then
                'if we encounter a word thats placed VERT, while we also want to put a word VERT, add count +1
                If used(xpos, i) = VERT Then
                    count = count + 1
                End If
                
                If Asc(UCase(Mid$(str, i - ypos + 1, 1))) <> field(xpos, i) Then
                    ' no match with char from other word, so it isn't possible
                    isDirPossible = False
                    Exit Function
                End If
            End If
        Next
        'if we encountered multiple words that are placed VERT, then placing it isn't possible
        If count > 0 Then
            isDirPossible = False
            Exit Function
        End If
    ElseIf dir = DIAG1 Then
        For i = 0 To Len(str) - 1
            If used(xpos + i, ypos + i) <> NONE Then
                If used(xpos + i, ypos + i) = DIAG1 Then
                    count = count + 1
                End If
                
                If Asc(UCase(Mid$(str, i + 1, 1))) <> field(xpos + i, ypos + i) Then
                    ' no match with char from previous word
                    isDirPossible = False
                    Exit Function
                End If
            End If
        Next
        If count > 0 Then
            isDirPossible = False
            Exit Function
        End If
    ElseIf dir = DIAG2 Then
        For i = 0 To Len(str) - 1
            If used(xpos + i, ypos - i) <> NONE Then
                If used(xpos + i, ypos - i) = DIAG2 Then
                    count = count + 1
                End If
                
                If Asc(UCase(Mid$(str, i + 1, 1))) <> field(xpos + i, ypos - i) Then
                    ' no match with char from previous word
                    isDirPossible = False
                    Exit Function
                End If
                
            End If
        Next
        If count > 0 Then
            isDirPossible = False
            Exit Function
        End If
    End If
    
    
    'Third test: check begin and ending of the word, and check if any other word next to it
    'is in the same direction
    'e.g if you have HELLO and WORLD, that you don't have  HELLOWORLD on the board
    
    If dir = HORZ Then
        'check begin and end if they are near another horz
        
        'begin
        If xpos - 1 >= 0 Then
        If used(xpos - 1, ypos) = HORZ Then
            isDirPossible = False
            Exit Function
        End If
        End If
        
        'end
        If (xpos + Len(str) - 1) + 1 < val(txtSize.Text) Then
        If used(xpos + Len(str), ypos) = HORZ Then
            isDirPossible = False
            Exit Function
        End If
        End If
    ElseIf dir = VERT Then
        'begin
        If ypos - 1 >= 0 Then
        If used(xpos, ypos - 1) = VERT Then
            isDirPossible = False
            Exit Function
        End If
        End If
        
        'end
        If (ypos + Len(str) - 1) + 1 < val(txtSize.Text) Then
        If used(xpos, ypos + Len(str)) = VERT Then
            isDirPossible = False
            Exit Function
        End If
        End If
    ElseIf dir = DIAG1 Then
        'begin
        If ypos - 1 >= 0 And xpos - 1 >= 0 Then
        If used(xpos - 1, ypos - 1) = DIAG1 Then
            isDirPossible = False
            Exit Function
        End If
        End If
        
        'end
        If (xpos + Len(str) - 1) + 1 < val(txtSize.Text) And _
           (ypos + Len(str) - 1) + 1 < val(txtSize.Text) Then
        If used(xpos + Len(str), ypos + Len(str)) = DIAG1 Then
            isDirPossible = False
            Exit Function
        End If
        End If
    ElseIf dir = DIAG2 Then
        'begin
        If ypos + 1 < val(txtSize.Text) And xpos - 1 >= 0 Then
        If used(xpos - 1, ypos + 1) = DIAG2 Then
            isDirPossible = False
            Exit Function
        End If
        End If
        
        'end
        If (xpos + Len(str) - 1) + 1 < val(txtSize.Text) And _
           ypos - 1 >= 0 Then
        If used(xpos + Len(str), ypos - 1) = DIAG2 Then
            isDirPossible = False
            Exit Function
        End If
        End If
    End If
    
    'after all test have been completed, its possible to put the word
    isDirPossible = True
    
End Function

Sub PutWord(str As String, xpos As Integer, ypos As Integer, dir As Integer)
    Dim i As Integer
    Dim sol As solution
    
    If dir = HORZ Then
        
        'put letters in the field array
        'and flag used to HORZ
        For i = xpos To xpos + Len(str) - 1
            field(i, ypos) = Asc(UCase(Mid$(str, i - xpos + 1, 1)))
            used(i, ypos) = HORZ
        Next
        
        ' add positions to solution
        sol.x1 = xpos
        sol.x2 = xpos + Len(str) - 1
        sol.y1 = ypos
        sol.y2 = ypos
        
    ElseIf dir = VERT Then
        
        'put letters in the field array
        'and flag used to VERT
        For i = ypos To ypos + Len(str) - 1
            field(xpos, i) = Asc(UCase(Mid$(str, i - ypos + 1, 1)))
            used(xpos, i) = VERT
        Next
        
        ' add positions to solution
        sol.x1 = xpos
        sol.x2 = xpos
        sol.y1 = ypos
        sol.y2 = ypos + Len(str) - 1
    ElseIf dir = DIAG1 Then
    
        'put letters in the field array
        'and flag used to DIAG1
        For i = 0 To Len(str) - 1
            field(xpos + i, ypos + i) = Asc(UCase(Mid$(str, i + 1, 1)))
            used(xpos + i, ypos + i) = DIAG1
        Next
        
        ' add positions to solution
        sol.x1 = xpos
        sol.x2 = xpos + Len(str) - 1
        sol.y1 = ypos
        sol.y2 = ypos + Len(str) - 1
    ElseIf dir = DIAG2 Then

        'put letters in the field array
        'and flag used to DIAG2
        For i = 0 To Len(str) - 1
            field(xpos + i, ypos - i) = Asc(UCase(Mid$(str, i + 1, 1)))
            used(xpos + i, ypos - i) = DIAG2
        Next
        
        ' add positions to solution
        sol.x1 = xpos
        sol.x2 = xpos + Len(str) - 1
        sol.y1 = ypos
        sol.y2 = ypos - (Len(str) - 1)
    End If

    'add current solution to the solutions array
    solutions(UBound(solutions)) = sol
    ReDim Preserve solutions(UBound(solutions) + 1)
    
    
End Sub

Private Sub txtfield_Change()
    Dim ascii As Byte
    If Len(txtfield.Text) <= 0 Then
        Exit Sub
    End If
    
    'check if typed letter is between [A...Z] or [a...z]
    
    ascii = Asc(Mid$(txtfield.Text, Len(txtfield.Text), 1))
    
    If (ascii < 65 Or ascii > 90) And (ascii < 97 Or ascii > 122) Then 'And (ascii > 90 Or ascii < 122) Then
        txtfield.Text = Mid$(txtfield.Text, 1, Len(txtfield.Text) - 1)
        txtfield.SelStart = Len(txtfield.Text)
    End If
    
End Sub

Private Sub txtfield_KeyPress(KeyAscii As Integer)
    'if enter is pressed when typing a word to add, add it
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
        Call cmdAdd_Click
    End If
End Sub

Sub DrawSolution()
    Dim i As Integer
    Dim sol As solution
    
    picfield.ForeColor = vbRed
    For i = 0 To UBound(solutions)
        sol = solutions(i)
        
        'center x and y in the cell
         'pseudo code:
         'x1 = cell1.Left + cell1.Width / 2
         'y1 = cell1.Top + cell1.Height / 2
        sol.x1 = sol.x1 * (picfield.Width \ val(txtSize.Text)) + ((picfield.Width \ val(txtSize.Text)) \ 2)
        sol.y1 = sol.y1 * (picfield.Width \ val(txtSize.Text)) + ((picfield.Width \ val(txtSize.Text)) \ 2)
        sol.x2 = sol.x2 * (picfield.Width \ val(txtSize.Text)) + ((picfield.Width \ val(txtSize.Text)) \ 2)
        sol.y2 = sol.y2 * (picfield.Width \ val(txtSize.Text)) + ((picfield.Width \ val(txtSize.Text)) \ 2)
        
        'draw a red line
        picfield.Line (sol.x1, sol.y1)-(sol.x2, sol.y2)

    Next
    picfield.ForeColor = vbBlack
    
End Sub

Sub FillRemWithRandom()
    'Fill all cells that haven't been used with random letters
    Dim i As Integer
    Dim j As Integer
    
    For j = 0 To val(txtSize.Text)
        For i = 0 To val(txtSize.Text)
            If used(i, j) = NONE Then
                field(i, j) = Rnd() * 25 + 65
            End If
        Next
    Next
End Sub
