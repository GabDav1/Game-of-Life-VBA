Attribute VB_Name = "GameOfLife"
Public nrColoane As Integer
Public nrLinii As Integer
Public colCel As Variant
Public colBgr As Variant
Public minNeighb As Integer
Public maxNeighb As Integer

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Settings_show()
Settings.Show
End Sub

    ''This checks the color scheme option and sets the color scheme parameters.
Sub Colorize()
If Sheets(2).Cells(4, 1).Value = 1 Then
    colCel = rgbRed
    colBgr = rgbBlue
ElseIf Sheets(2).Cells(4, 1).Value = 2 Then
    colCel = rgbWhite
    colBgr = rgbBlack
ElseIf Sheets(2).Cells(4, 1).Value = 3 Then
    colCel = rgbGreen
    colBgr = rgbBrown
End If
End Sub
    
    ''Game logic goes here.
Sub game_life()
Dim parcurgere As Integer
Dim vecini As Integer
Dim iplus As Integer
Dim jplus As Integer
Dim iminus As Integer
Dim jminus As Integer
Dim cellcounter As Integer

        ''Sets the variables for size of board (number of columns and lines).
        ''Also sets the variables that define the rules of the game:
        ''minimum neighbors a cell needs to survive(minNeighb)
        ''maximum neighbors a cell can have(maxNeighb)
nrColoane = Sheets(2).Cells(1, 1)
nrLinii = Sheets(2).Cells(2, 1)
minNeighb = Sheets(2).Cells(5, 1)
maxNeighb = Sheets(2).Cells(6, 1)

        ''Colors the board according to user selected color scheme.
Call Colorize
Range("A2", Cells(Rows.Count, Columns.Count)).Font.Color = colBgr
parcurgere = 1
If parcurgere = 1 Then
    For j = 1 To nrColoane
        For i = 2 To nrLinii
            If Sheets(1).Cells(i, j).Value = "|" Then
                Sheets(1).Cells(i, j).Font.Color = colCel
                Sheets(1).Cells(i, j).Interior.Color = colCel
            End If
        Next i
    Next j
End If

''Runs the game for a number of rounds. The number of rounds is defined by the user.
Do While parcurgere < Sheets(2).Cells(3, 1).Value
    For j = 1 To nrColoane
        For i = 2 To nrLinii
        
    '' Transforms the "X" markings into terrain/dead cells and the "O" markings into living cells
            If Sheets(1).Cells(i, j).Value = "X" Then
                Sheets(1).Cells(i, j).Font.Color = colBgr
                Sheets(1).Cells(i, j).Value = "_"
            End If
        
            If Sheets(1).Cells(i, j).Value = "O" Then
                Sheets(1).Cells(i, j).Font.Color = colCel
                Sheets(1).Cells(i, j).Interior.Color = colCel
                Sheets(1).Cells(i, j).Value = "|"
            End If
           
    '' The jplus/iplus and jminus/iminus variables are used for the circular edges: the board always wraps around itself.
            If j = nrColoane Then
                jplus = 1
            Else
                jplus = j + 1
            End If
            
            If j = 1 Then
                jminus = nrColoane
            Else
                jminus = j - 1
            End If
            
    '' The ipjp and imjm variables have been implemented to repair a bug that occured when counting the neighbors of the cells
    '' that were located at the upper or lower bounds of the board (due to the board's circular edges). 
            If i = nrLinii Then
                iplus = 2
                ipjp = "X"
            Else
                iplus = i + 1
                ipjp = "O"
            End If
            
            If i = 2 Then
                iminus = nrLinii
                imjm = "O"
            Else
                iminus = i - 1
                imjm = "X"
            End If
           
    '' Counts neighbours on all 8 positions that surround a cell. In my implementation I had to take into account if
    '' the neighbour that is being counted has been traversed or not. This combined with the circular edges proved to be
    '' an interesting challenge. I commented on each line if it represents a neighbour that has been traversed or not.
            If (Sheets(1).Cells(i, jminus).Value = "|" Or Sheets(1).Cells(i, jminus).Value = "X") Then 'TRAVERSED
                vecini = vecini + 1
            End If
            
            If (Sheets(1).Cells(iplus, jminus).Value = "|" Or Sheets(1).Cells(iplus, jminus).Value = "X") Then 'TRAVERSED
                vecini = vecini + 1
            End If
           
            If (Sheets(1).Cells(iminus, jminus).Value = "|" Or Sheets(1).Cells(iminus, jminus).Value = "X") Then 'TRAVERSED
                vecini = vecini + 1
            End If
            
            If (Sheets(1).Cells(iminus, j).Value = "|" Or Sheets(1).Cells(iminus, j).Value = imjm) Then 'TRAVERSED
                vecini = vecini + 1
            End If
            
            If (Sheets(1).Cells(iminus, jplus).Value = "|" Or Sheets(1).Cells(iminus, jplus).Value = "O") Then 'NOT TRAVERSED
                vecini = vecini + 1
            End If
            
            If (Sheets(1).Cells(iplus, jplus).Value = "|" Or Sheets(1).Cells(iplus, jplus).Value = "O") Then 'NOT TRAVERSED
                vecini = vecini + 1
            End If
            
            If (Sheets(1).Cells(i, jplus).Value = "|" Or Sheets(1).Cells(i, jplus).Value = "O") Then 'NOT TRAVERSED
                vecini = vecini + 1
            End If
            
            If (Sheets(1).Cells(iplus, j).Value = "|" Or Sheets(1).Cells(iplus, j).Value = ipjp) Then 'NOT TRAVERSED
                vecini = vecini + 1
            End If
            
    '' This marks what cells will be born and what cells will die at the start of the next round.
            If (vecini < minNeighb Or vecini > maxNeighb) And Sheets(1).Cells(i, j).Value = "|" Then
                Sheets(1).Cells(i, j).Font.Color = colBgr
                Sheets(1).Cells(i, j).Interior.Color = colBgr
                Sheets(1).Cells(i, j).Value = "X"
            End If
            
            If Sheets(1).Cells(i, j).Value = "_" And vecini = maxNeighb Then
                Sheets(1).Cells(i, j).Font.Color = colCel
                Sheets(1).Cells(i, j).Interior.Color = colCel
                Sheets(1).Cells(i, j).Value = "O"
            End If
            
            If Cells(i, j).Interior.Color = colCel Then
                cellcounter = cellcounter + 1
            End If
                 
            vecini = 0
                        
        Next i
    Next j
    
    parcurgere = parcurgere + 1
    Sheets(1).Cells(1, 1).Value = "R: " & parcurgere
    Sheets(1).Range("H1").Value = "C: " & cellcounter
    cellcounter = 0

Loop

End Sub

'' Adds a glider cell-formation around the selected cell.
Sub add_glider()
Selection.Offset(1, -1).Font.Color = rgbBlack
Selection.Offset(1, -1).Value = "|"
Selection.Offset(1, 0).Font.Color = rgbBlack
Selection.Offset(1, 0).Value = "|"
Selection.Offset(1, 1).Font.Color = rgbBlack
Selection.Offset(1, 1).Value = "|"
Selection.Offset(0, 1).Font.Color = rgbBlack
Selection.Offset(0, 1).Value = "|"
Selection.Offset(1, -1).Font.Color = rgbBlack
Selection.Offset(1, -1).Value = "|"
Selection.Offset(-1, 0).Font.Color = rgbBlack
Selection.Offset(-1, 0).Value = "|"
End Sub

''A simple function to wipe the board clean.
Sub clear_board()
Dim x As Integer
Dim y As Integer

nrColoane = Sheets(2).Cells(1, 1)
nrLinii = Sheets(2).Cells(2, 1)
coldesters = Cells(2, Columns.Count).End(xlToLeft).Column
lindesters = Range("A" & Rows.Count).End(xlUp).Row

For x = 1 To coldesters
        For y = 2 To lindesters
            Sheets(1).Cells(y, x).Interior.Color = rgbWhite
            Sheets(1).Cells(y, x).Value = " "
        Next y
Next x
Call Colorize

For x = 1 To nrColoane
        For y = 2 To nrLinii
                Sheets(1).Cells(y, x).Interior.Color = colBgr
                Sheets(1).Cells(y, x).Font.Color = colBgr
                Sheets(1).Cells(y, x).Value = "_"
        Next y
Next x
End Sub

'' This function adds a circle on the board around the selected cell. 
'' The circle's radius is the distance from the selected cell to the closest margin of the board. 
Sub cerc()
Dim x As Integer
Dim y As Integer
Dim raza As Integer
Dim distancetobottom As Integer
Dim distancetotop As Integer
Dim distancetoleft As Integer
Dim distancetoright As Integer

distancetoleft = Selection.Column - 1
distancetoright = nrColoane - Selection.Column
distancetotop = Selection.Row - 2
distancetobottom = nrLinii - Selection.Row

raza = WorksheetFunction.Min(distancetotop, distancetobottom, distancetoleft, distancetoright)

For x = 1 To raza
y = Sqr(raza * raza - x * x)
Selection.Offset(x, y).Font.Color = rgbBlack
Selection.Offset(x, y).Value = "|"
Selection.Offset(x, -y).Font.Color = rgbBlack
Selection.Offset(x, -y).Value = "|"
Selection.Offset(-x, y).Font.Color = rgbBlack
Selection.Offset(-x, y).Value = "|"
Selection.Offset(-x, -y).Font.Color = rgbBlack
Selection.Offset(-x, -y).Value = "|"
Next x
End Sub

''This function adds a custom number of cells in random locations on the board.
Sub addRandom()
Dim nrpoints As Integer
nrpoints = InputBox("How many cells do you want to place?")

nrColoane = Sheets(2).Cells(1, 1)
nrLinii = Sheets(2).Cells(2, 1)

For i = 1 To nrpoints
    x = (nrLinii - 1) * Rnd + 1
    y = nrColoane * Rnd
    
''The rnd function can produce floating point values of x and y that are less than 1.
''Our board's coordinates are defined by integers that begin at 1 so we will transform those numbers.
    If x < 1 Then
        x = 1
    End If
    If y < 1 Then
        y = 1
    End If
    
    Cells(x, y).Font.Color = rgbBlack
    Cells(x, y).Value = "|"
Next i
End Sub
