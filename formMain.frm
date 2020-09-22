VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Axi's Demo AZ666"
   ClientHeight    =   6510
   ClientLeft      =   2010
   ClientTop       =   2475
   ClientWidth     =   8175
   Icon            =   "formMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8175
   Begin VB.CommandButton cmdEditText 
      Caption         =   "Edit"
      Height          =   375
      Left            =   6360
      TabIndex        =   27
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdChooseABCFont 
      Caption         =   "Font"
      Height          =   375
      Left            =   6360
      TabIndex        =   26
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar prgbarMain 
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   3720
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdShowHidePaint 
      Caption         =   ">"
      Height          =   1575
      Left            =   7560
      TabIndex        =   24
      Top             =   4080
      Width           =   495
   End
   Begin VB.CheckBox chkFillAll 
      Caption         =   "Do Not Use The Timer"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   4920
      Width           =   2295
   End
   Begin VB.HScrollBar scrollColorValue 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2880
      Max             =   360
      Min             =   1
      TabIndex        =   19
      Top             =   5220
      Value           =   1
      Width           =   3375
   End
   Begin VB.TextBox txtColorValue 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1037
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      MaxLength       =   5
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   4920
      Width           =   735
   End
   Begin VB.CheckBox chkUseColorValue 
      Caption         =   "Change Color by Value"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      ToolTipText     =   "Active Only While Target Image is Empty"
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Frame frmImages 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   7455
      Begin VB.PictureBox picSource 
         AutoRedraw      =   -1  'True
         Height          =   3360
         Left            =   0
         MousePointer    =   2  'Cross
         Picture         =   "formMain.frx":08CA
         ScaleHeight     =   3184.211
         ScaleMode       =   0  'User
         ScaleWidth      =   3540
         TabIndex        =   14
         ToolTipText     =   "Click The Source Image to Choose a Different one..."
         Top             =   0
         Width           =   3600
         Begin VB.Timer TimerOpenPaintWindow 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   2760
            Top             =   720
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   2760
            Top             =   240
         End
         Begin VB.Timer TimerRandomSquares 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   2760
            Top             =   1680
         End
      End
      Begin VB.PictureBox picTarget 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   3360
         Left            =   3840
         ScaleHeight     =   3300
         ScaleWidth      =   3540
         TabIndex        =   13
         ToolTipText     =   "Left Click This Image To Save It or Right Click to change Image Background's Color (only after reset)..."
         Top             =   0
         Width           =   3600
      End
   End
   Begin VB.TextBox txtCircleRadius 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1037
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   2880
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtBrushSize 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1037
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   4440
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   3600
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSpeed 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1037
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   6240
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4140
      Width           =   735
   End
   Begin VB.HScrollBar scrollSpeed 
      Height          =   255
      Left            =   4080
      Max             =   100
      Min             =   1
      TabIndex        =   4
      Top             =   4440
      Value           =   1
      Width           =   3375
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Reset"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   3
      Top             =   5760
      Width           =   2415
   End
   Begin VB.ComboBox cmdBrushStyle 
      Height          =   315
      ItemData        =   "formMain.frx":8AE7
      Left            =   240
      List            =   "formMain.frx":8AFD
      TabIndex        =   2
      Text            =   "Circle"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Pause"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label lblT 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Target"
      Height          =   255
      Index           =   6
      Left            =   3960
      TabIndex        =   22
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label lblT 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Source"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label lblT 
      BackStyle       =   0  'Transparent
      Caption         =   "Color Change Value:"
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   20
      Top             =   4980
      Width           =   1575
   End
   Begin VB.Label lblRandomCounter 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   16
      Top             =   4500
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblRandomCounter 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   15
      Top             =   4500
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblT 
      BackStyle       =   0  'Transparent
      Caption         =   "Circle Radius:"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   10
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblT 
      BackStyle       =   0  'Transparent
      Caption         =   "Steps Size:"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   8
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblT 
      BackStyle       =   0  'Transparent
      Caption         =   "Brush Style:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblT 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed (1000 = 1sec):"
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
   End
End
Attribute VB_Name = "formMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkFillAll_Click()
    If chkFillAll.value = 1 Then
        txtSpeed.Enabled = False
        scrollSpeed.Enabled = False
    Else
        txtSpeed.Enabled = True
        scrollSpeed.Enabled = True
    End If
    End Sub

Private Sub chkUseColorValue_Click()
    If chkUseColorValue.value = 1 Then ' checked
        scrollColorValue.Enabled = True
        txtColorValue.Enabled = True
    Else
        scrollColorValue.Enabled = False
        txtColorValue.Enabled = False
    End If
End Sub

Private Sub cmdBrushStyle_Click()
    lblT(2).Visible = True
    txtBrushSize.Visible = True
    If cmdBrushStyle.ListIndex = 1 Then ' circle
        lblT(3).Visible = True
        txtCircleRadius.Visible = True
    Else
        lblT(3).Visible = False
        txtCircleRadius.Visible = False
        lblRandomCounter(1).Visible = False
        lblRandomCounter(0).Visible = False
        If cmdBrushStyle.ListIndex = 3 Then ' random
            lblT(2).Visible = False
            txtBrushSize.Visible = False
            chkFillAll.value = 0
            chkFillAll.Visible = False
        Else
            chkFillAll.value = 0
            chkFillAll.Visible = True
            If cmdBrushStyle.ListIndex = 5 Then ' ABC...
                cmdChooseABCFont.Visible = True
                cmdEditText.Visible = True
            Else
                cmdChooseABCFont.Visible = False
                cmdEditText.Visible = False
            End If
        End If
    End If
    Call ResetAll
End Sub

Private Sub cmdChooseABCFont_Click()
    On Error Resume Next
    cmdlg.Flags = cdlCFBoth
    cmdlg.FontName = picTarget.FontName
    cmdlg.ShowFont
    txtBrushSize.Text = cmdlg.FontSize * 20
    picTarget.FontBold = cmdlg.FontBold
    picTarget.FontItalic = cmdlg.FontItalic
    picTarget.FontName = cmdlg.FontName
    picTarget.FontSize = cmdlg.FontSize
    picTarget.FontStrikethru = cmdlg.FontStrikethru
    picTarget.FontUnderline = cmdlg.FontUnderline
End Sub

Private Sub cmdClear_Click()
    X = 0
    Y = 0
    ReDim SelectedNums(1)
    picTarget.Cls
    cmdClear.Enabled = False
    cmdStart.Enabled = True
    lblRandomCounter(0) = 0
    lblRandomCounter(1).Visible = False
    lblRandomCounter(0).Visible = False
    prgbarMain.value = 0
    txtColorValue.Enabled = True
    scrollColorValue.Enabled = True
    chkUseColorValue.Enabled = True
    If strText <> "" Then
        intCharindex = 0
    Else
        intCharindex = 64
    End If
End Sub

Private Sub cmdEditText_Click()
    formText.Show (1)
End Sub

Private Sub cmdShowHidePaint_Click()
    Me.Enabled = False
    formPaint.Enabled = False
    If cmdShowHidePaint.Caption = ">" Then
        formPaint.Top = Me.Top
        formPaint.Left = Me.Left + Me.Width - formPaint.Width
        cmdStart.Enabled = False
        cmdStop.Enabled = False
        formPaint.Show
        Me.Show
        formMain.TimerOpenPaintWindow = True
        Timer1 = False
    Else
        X = 0
        Y = 0
        cmdStart.Enabled = True
        cmdStop.Enabled = True
        prgbarMain.value = 0
        picTarget.Cls
        TimerOpenPaintWindow = True
    End If
End Sub

Private Sub cmdStart_Click()
    On Error GoTo errorRun
    If chkFillAll.value = 0 Then
        If cmdClear.Enabled <> True Then ' pause
            X = 0
            Y = 0
            prgbarMain.value = 0
        End If
        lblRandomCounter(1).Visible = False
        lblRandomCounter(0).Visible = False
        If (cmdBrushStyle.ListIndex) = 3 Then
            Call RandomInIt
            lblRandomCounter(1).Visible = True
            lblRandomCounter(0).Visible = True
            lblRandomCounter(1) = " - " & Int((picSource.Width / 120) * ((picSource.Height / 120) + 1))
            TimerRandomSquares = True
        Else
            prgbarMain.Max = (picSource.Width / txtBrushSize) * (picSource.Height / txtBrushSize)
            Timer1 = True
        End If
        frmImages.Enabled = False
        cmdStart.Enabled = False
        cmdStop.Enabled = True
        cmdClear.Enabled = False
        cmdBrushStyle.Enabled = False
        txtColorValue.Enabled = False
        chkUseColorValue.Enabled = False
        chkFillAll.Enabled = False
        scrollColorValue.Enabled = False
        txtBrushSize.Enabled = False
        txtCircleRadius.Enabled = False
        cmdEditText.Enabled = False
        cmdChooseABCFont.Enabled = False
    Else
        prgbarMain.value = 0
        X = 0
        Y = 0
        cmdStart.Enabled = False
        cmdClear.Enabled = True
        Call FillAll
    End If
    Exit Sub
errorRun:
    MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Error !!!"
End Sub

Private Sub cmdStop_Click()
    frmImages.Enabled = True
    Timer1 = False
    cmdStart.Enabled = True
    cmdStop.Enabled = False
    cmdClear.Enabled = True
    cmdBrushStyle.Enabled = True
    chkUseColorValue.Enabled = True
    chkFillAll.Enabled = True
    If chkUseColorValue = 1 Then
        scrollColorValue.Enabled = True
        txtColorValue.Enabled = True
    End If
    txtBrushSize.Enabled = True
    txtCircleRadius.Enabled = True
    TimerRandomSquares = False
    cmdEditText.Enabled = True
    cmdChooseABCFont.Enabled = True
End Sub

Private Sub Command1_Click()
    formText.Show (1)
End Sub

Private Sub Form_Load()
    X = 0
    Y = 0
    picSource.CurrentX = 120
    picSource.CurrentY = picSource.Height / 2 - 360
    picSource.FontSize = 18
    picSource.Print "Axi from USHASOFT"
    picSource.FontSize = 24
    picSource.ForeColor = &HC00000
    picSource.CurrentX = 1200
    picSource.CurrentY = picSource.Height / 2
    picSource.Print "2005"
    scrollSpeed.value = 20
    scrollColorValue.value = 30
    cmdBrushStyle.ListIndex = 1
    txtBrushSize = 120
    txtCircleRadius = 40
    intCharindex = 64
    strText = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub picSource_Click()
    On Error GoTo CancelError
    cmdlg.Filter = "Pictures (*.bmp;*.ico;*.jpg;*.gif)|*.bmp;*.ico;*.jpg;*.gif"
    cmdlg.CancelError = True
    cmdlg.ShowOpen
    Call picSource.PaintPicture(LoadPicture(cmdlg.FileName), 0, 0, picSource.Width, picSource.Height)
    Call cmdClear_Click
    Exit Sub
CancelError:
End Sub


Private Sub picTarget_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo CancelError
    If Button = vbLeftButton Then   ' left click
        cmdlg.Filter = "Pictures (*.bmp;*.ico;*.jpg;*.gif)|*.bmp;*.ico;*.jpg;*.gif"
        cmdlg.Flags = cdlOFNFileMustExist And cdlOFNOverwritePrompt
        cmdlg.CancelError = True
        cmdlg.ShowSave
        SavePicture picTarget.Image, cmdlg.FileName
    Else                            ' right click
        If cmdClear.Enabled = False And cmdStop.Enabled = False Then
            cmdlg.color = picTarget.BackColor
            cmdlg.ShowColor
            picTarget.BackColor = cmdlg.color
        End If
    End If
    Exit Sub
CancelError:
End Sub

Private Sub scrollColorValue_Change()
    txtColorValue = scrollColorValue.value
End Sub

Private Sub scrollColorValue_Scroll()
    txtColorValue = scrollColorValue.value
End Sub

Private Sub scrollSpeed_Change()
    txtSpeed = scrollSpeed.value
    Timer1.Interval = txtSpeed
End Sub

Private Sub scrollSpeed_Scroll()
    txtSpeed = scrollSpeed.value
    Timer1.Interval = txtSpeed
End Sub

Private Sub Timer1_Timer()
    BrushColor = picSource.point(X, Y)
    If chkUseColorValue.value = 1 Then ' change color
        BrushColor = color_val(BrushColor, txtColorValue)
    End If
    Select Case (cmdBrushStyle.ListIndex)
    Case 0 ' Triangle
        Call DrawTriangle(X, Y, X + txtBrushSize, Y + txtBrushSize, BrushColor)
    Case 1 ' circle
        picTarget.Circle (X, Y), txtCircleRadius, BrushColor
    Case 2 ' squares
        Call DrawSquare(X, Y, X + txtBrushSize, Y + txtBrushSize, BrushColor)
    Case 4 ' lines
        picTarget.Line (X, Y)-(X + txtBrushSize, Y + txtBrushSize), BrushColor
        picTarget.Line (X + txtBrushSize, Y)-(X, Y + txtBrushSize), BrushColor
    Case 5 ' ABC...
        picTarget.CurrentX = X
        picTarget.CurrentY = Y
        picTarget.ForeColor = Abs(BrushColor)
        picTarget.Print GiveLetter
    End Select
    X = X + txtBrushSize
    If X + txtBrushSize > picSource.Width Then
        Y = Y + txtBrushSize
        X = 0
    End If
    If Y + txtBrushSize > picSource.Height Then
        Call cmdStop_Click
    End If
    If formMain.prgbarMain.value < formMain.prgbarMain.Max Then
        formMain.prgbarMain.value = formMain.prgbarMain.value + 1
    End If
End Sub

Private Sub TimerOpenPaintWindow_Timer()
    With formPaint
        If formMain.cmdShowHidePaint.Caption = ">" Then
            .Left = .Left + 60
            If .Left >= formMain.Left + formMain.Width Then
                formMain.cmdShowHidePaint.Caption = "<"
                TimerOpenPaintWindow = False
                formMain.Enabled = True
                formPaint.Enabled = True
                formPaint.SetFocus
            End If
        Else
            .Left = .Left - 60
            If .Left <= formMain.Left + formMain.Width - .Width Then
                formMain.cmdShowHidePaint.Caption = ">"
                TimerOpenPaintWindow = False
                formMain.Enabled = True
                formPaint.Enabled = True
                formMain.SetFocus
                .Hide
            End If
        End If
    End With
End Sub

Private Sub TimerRandomSquares_Timer()
    Call GetRandomNum(picSource.Width / 120, (picSource.Height / 120) + 1, X, Y)
    lblRandomCounter(0) = NumsNumber
    X = X * 120
    Y = Y * 120
    BrushColor = picSource.point(X, Y)
    If chkUseColorValue.value = 1 Then ' change color
        BrushColor = color_val(BrushColor, txtColorValue)
    End If
    Call DrawSquare(X, Y, X + 120, Y + 120, BrushColor)
    X = X + 120
    If NumsNumber = Int((picSource.Width / 120) * (picSource.Height / 120 + 1)) Then
        BrushColor = picSource.point(0, 0)
        Call DrawSquare(0, 0, 120, 120, BrushColor)
        cmdStop_Click
    End If
End Sub

Private Sub txtBrushSize_Change()
    Call ResetAll
End Sub


Private Sub txtCircleRadius_Change()
    Call ResetAll
End Sub

Private Sub txtSpeed_Change()
    scrollSpeed.value = txtSpeed
End Sub

