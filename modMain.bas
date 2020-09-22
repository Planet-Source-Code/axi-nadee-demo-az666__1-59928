Attribute VB_Name = "modMain"
Public X, Y As Long
Public BrushColor As Double
Type point
    X As Long
    Y As Long
End Type
Public SelectedNums() As point
Public NumsNumber As Long
Public MouseClick As Boolean
Public DrawBrushNum As Long
Public strText As String
Public intCharindex As Integer
Public corX, corY As Long

Function DrawSquare(x1, y1, x2, y2 As Long, SquareColor As Double)
    Dim i As Long
    For i = y1 + 10 To y2 Step 10
        formMain.picTarget.Line (x1, i)-(x2, i), SquareColor
    Next i
End Function
Function DrawTriangle(x1, y1, x2, y2 As Long, SquareColor As Double)
    Dim i As Long
    For i = y1 To y2
        formMain.picTarget.Line (x1, y1)-(x2, i), SquareColor
    Next
End Function

Function ResetAll()
    formMain.picTarget.Cls
    X = 0
    Y = 0
    formMain.cmdStart.Enabled = True
    formMain.cmdStop.Enabled = False
    formMain.cmdClear.Enabled = False
End Function

Function GetRandomNum(xMax As Long, yMax As Long, ByRef xSource, ySource As Long)
    Dim i As Long
Loop1:
    xSource = Int(Rnd * xMax)
    ySource = Int(Rnd * yMax)
    For i = 0 To UBound(SelectedNums) - 1
        If SelectedNums(i).X = xSource And SelectedNums(i).Y = ySource Then
            GoTo Loop1
         End If
    Next
    NumsNumber = NumsNumber + 1
    ReDim Preserve SelectedNums(UBound(SelectedNums) + 1)
    SelectedNums(UBound(SelectedNums) - 1).X = xSource
    SelectedNums(UBound(SelectedNums) - 1).Y = ySource
End Function

Function RandomInIt()
    ReDim SelectedNums(1)
    NumsNumber = 1
End Function

Function color_val(f_color As Double, value As Double) As Double
    Dim Red     As Integer
    Dim Green   As Integer
    Dim Blue    As Integer
    Red = Green = Blue = 0
    Red = f_color Mod 256
    If Red + (value) < 0 Then
        Red = Red + (Red + (value))
    Else
        If Red + (value) > 255 Then
            Red = Red + (value) - 255
        Else
            Red = Red + (value)
        End If
    End If
    Green = ((f_color - Red) / 256) Mod 256
    If Green + (value) < 0 Then
        Green = Green + (Green + (value))
    Else
        If Green + (value) > 255 Then
            Green = Green + (value) - 255
        Else
            Green = Green + (value)
        End If
    End If
    Blue = (f_color - Red - CDbl(256) * Green) / (CDbl(256) * 256)
    If Blue + (value) < 0 Then
        Blue = Blue + (Blue + (value))
    Else
        If Blue + (value) > 255 Then
            Blue = Blue + (value) - 255
        Else
            Blue = Blue + (value)
        End If
    End If
    color_val = RGB(Red, Green, Blue)
End Function

Function FillAll()
    With formMain
    .prgbarMain.Max = (.picSource.Width / Int(.txtBrushSize)) * (.picSource.Height / Int(.txtBrushSize))
    .prgbarMain.value = 0
    While Not (Y > .picSource.Height + .txtBrushSize)
        BrushColor = .picSource.point(X, Y)
        If .chkUseColorValue.value = 1 Then ' change color
            BrushColor = color_val(BrushColor, .txtColorValue)
        End If
        Select Case (.cmdBrushStyle.ListIndex)
        Case 0 ' Triangle
            Call DrawTriangle(X, Y, X + .txtBrushSize, Y + .txtBrushSize, BrushColor)
        Case 1 ' circle
            .picTarget.Circle (X, Y), .txtCircleRadius, BrushColor
        Case 2 ' squares
            Call DrawSquare(X, Y, X + .txtBrushSize, Y + .txtBrushSize, BrushColor)
        Case 4 ' lines
            .picTarget.Line (X, Y)-(X + .txtBrushSize, Y + .txtBrushSize), BrushColor
            .picTarget.Line (X + .txtBrushSize, Y)-(X, Y + .txtBrushSize), BrushColor
        End Select
        X = X + .txtBrushSize
        If X > .picSource.Width + .txtBrushSize Then
            Y = Y + .txtBrushSize
            X = 0
        End If
        If .prgbarMain.value + 1 < .prgbarMain.Max Then
            .prgbarMain.value = .prgbarMain.value + 1
        End If
    Wend
    End With
End Function
Function PaintPhoto()
    formMain.prgbarMain.Max = (formPaint.picTarget.Width / formMain.txtBrushSize) * (formPaint.picTarget.Height / formMain.txtBrushSize)
    While Y < formPaint.picTarget.Height
        BrushColor = formPaint.picTarget.point(X, Y)
        If BrushColor = 0 Then ' black
            X = X + formMain.txtBrushSize
        Else
            BrushColor = formMain.picSource.point(X, Y)
            Select Case (formMain.cmdBrushStyle.ListIndex)
            Case 0 ' Triangle
                Call DrawTriangle(X, Y, X + formMain.txtBrushSize, Y + formMain.txtBrushSize, BrushColor)
            Case 1 ' circle
                formMain.picTarget.Circle (X, Y), formMain.txtCircleRadius, BrushColor
            Case 2 ' squares
                Call DrawSquare(X, Y, X + formMain.txtBrushSize, Y + formMain.txtBrushSize, BrushColor)
            Case 4 ' lines
                formMain.picTarget.Line (X, Y)-(X + formMain.txtBrushSize, Y + formMain.txtBrushSize), BrushColor
                formMain.picTarget.Line (X + formMain.txtBrushSize, Y)-(X, Y + formMain.txtBrushSize), BrushColor
            Case 5 ' ABC...
                picTarget.CurrentX = X
                picTarget.CurrentY = Y
                picTarget.ForeColor = Abs(BrushColor)
                picTarget.Print GiveLetter
            End Select
            X = X + formMain.txtBrushSize
        End If
        If X >= formPaint.picTarget.Width Then
            Y = Y + formMain.txtBrushSize
            X = 0
        End If
        If formMain.prgbarMain.value + 1 < formMain.prgbarMain.Max Then
            formMain.prgbarMain.value = formMain.prgbarMain.value + 1
        End If
    Wend
End Function

Function GiveLetter() As String
    intCharindex = intCharindex + 1
    If strText <> "" Then
        GiveLetter = Mid(strText, intCharindex, 1)
        If intCharindex + 1 > Len(strText) Then
            intCharindex = 0
        End If
    Else
        GiveLetter = Chr(intCharindex)
        If intCharindex + 1 > 122 Then
            intCharindex = 64
        End If
    End If
End Function
