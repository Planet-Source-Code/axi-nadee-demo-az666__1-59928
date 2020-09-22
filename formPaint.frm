VERSION 5.00
Begin VB.Form formPaint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paint"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   Icon            =   "formPaint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar scrollBrushSize 
      Height          =   255
      Left            =   240
      Max             =   100
      Min             =   1
      TabIndex        =   10
      Top             =   4560
      Value           =   4
      Width           =   3375
   End
   Begin VB.CommandButton cmdNewImage 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      TabIndex        =   9
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdFillCircle 
      Height          =   435
      Left            =   720
      Picture         =   "formPaint.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   450
   End
   Begin VB.CommandButton cmdFillSquare 
      Height          =   435
      Left            =   1920
      Picture         =   "formPaint.frx":0C56
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   450
   End
   Begin VB.CommandButton cmdLine 
      Height          =   435
      Left            =   2520
      Picture         =   "formPaint.frx":0D75
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   450
   End
   Begin VB.CommandButton cmdSquare 
      Height          =   435
      Left            =   1320
      Picture         =   "formPaint.frx":10E3
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   450
   End
   Begin VB.CommandButton cmdCircle 
      Height          =   435
      Left            =   120
      Picture         =   "formPaint.frx":1324
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   450
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
      TabIndex        =   1
      Top             =   5760
      Width           =   3615
   End
   Begin VB.PictureBox picTarget 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   4
      FillColor       =   &H00FFFFFF&
      Height          =   3360
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   3300
      ScaleWidth      =   3540
      TabIndex        =   0
      ToolTipText     =   "Left Click This Image To Save It or Right Click to change Image Background's Color (only after reset)..."
      Top             =   240
      Width           =   3600
      Begin VB.Shape shpCircle 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   2  'Dash
         DrawMode        =   7  'Invert
         FillColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   960
         Shape           =   3  'Circle
         Top             =   1800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Shape shpSquare 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   2  'Dash
         DrawMode        =   7  'Invert
         FillColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   960
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Label lblT 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Paint"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label lblBrushSize 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblT 
      BackStyle       =   0  'Transparent
      Caption         =   "Brush Size:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
End
Attribute VB_Name = "formPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCircle_Click()
    DrawBrushNum = 0
End Sub

Private Sub cmdFillCircle_Click()
    DrawBrushNum = 1
End Sub

Private Sub cmdFillSquare_Click()
    DrawBrushNum = 3
End Sub

Private Sub cmdLine_Click()
    DrawBrushNum = 4
End Sub

Private Sub cmdNewImage_Click()
    picTarget.Cls
End Sub

Private Sub cmdSquare_Click()
    DrawBrushNum = 2
End Sub

Private Sub cmdStart_Click()
    X = 0
    Y = 0
    formMain.picTarget.Cls
    formMain.prgbarMain.value = 0
    PaintPhoto
End Sub



Private Sub Form_Unload(Cancel As Integer)
    X = 0
    Y = 0
    formMain.cmdStart.Enabled = True
    formMain.picTarget.Cls
End Sub

Private Sub picTarget_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseClick = True
    Select Case DrawBrushNum
    Case 0
        shpCircle.Left = X
        shpCircle.Top = Y
        shpCircle.Width = 0
        shpCircle.Height = 0
        shpCircle.Visible = True
        shpCircle.BorderWidth = picTarget.DrawWidth
    Case 1
        shpCircle.Left = X
        shpCircle.Top = Y
        shpCircle.Width = 0
        shpCircle.Height = 0
        shpCircle.Visible = True
        shpCircle.BorderWidth = picTarget.DrawWidth
        shpCircle.FillStyle = 0
        picTarget.FillStyle = 0
    Case 2
        shpSquare.Left = X
        shpSquare.Top = Y
        shpSquare.Width = 0
        shpSquare.Height = 0
        shpSquare.Visible = True
        shpSquare.BorderWidth = picTarget.DrawWidth
    Case 3
        shpSquare.Left = X
        shpSquare.Top = Y
        shpSquare.Width = 0
        shpSquare.Height = 0
        shpSquare.Visible = True
        shpSquare.BorderWidth = picTarget.DrawWidth
        shpSquare.FillStyle = 0
        picTarget.FillStyle = 0
    Case 4
        picTarget.CurrentX = X
        picTarget.CurrentY = Y
    End Select
    corX = X
    corY = Y
End Sub

Private Sub picTarget_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    If MouseClick Then
            Select Case DrawBrushNum
            Case 0, 1
                shpCircle.Width = Abs(X - corX)
                shpCircle.Height = Abs(Y - corY)
                If X - corX < 0 Then
                    shpCircle.Left = X
                End If
                If Y - corY < 0 Then
                    shpCircle.Top = Y
                End If
            Case 2, 3
                shpSquare.Width = Abs(X - corX)
                shpSquare.Height = Abs(Y - corY)
                If X - corX < 0 Then
                    shpSquare.Left = X
                End If
                If Y - corY < 0 Then
                    shpSquare.Top = Y
                End If
            Case 4
                If Button = vbLeftButton Then
                    picTarget.Line (picTarget.CurrentX, picTarget.CurrentY)-(X, Y), &HFFFFFF
                Else
                    picTarget.Line (picTarget.CurrentX, picTarget.CurrentY)-(X, Y), &H0
                End If
            End Select
        End If
End Sub

Private Sub picTarget_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim color As Double
    If Button = vbLeftButton Then
        color = &HFFFFFF
    Else
        color = &H0
    End If
    picTarget.FillColor = color
    MouseClick = False
    Select Case DrawBrushNum
            Case 0, 1
                picTarget.Circle (shpCircle.Left + (shpCircle.Width / 2), shpCircle.Top + (shpCircle.Height / 2)), Abs(X - corX) / 2, color
                shpCircle.Visible = False
                shpCircle.FillStyle = 1
                picTarget.FillStyle = 1
            Case 2, 3
                picTarget.Line (corX, corY)-Step(X - corX, Y - corY), color, B
                shpSquare.Visible = False
                shpSquare.FillStyle = 1
                picTarget.FillStyle = 1
            Case 4
                For i = 0 To 5
                    picTarget.Line (X, Y + i)-(X + 10, Y + i), &HFFFFFF
                    picTarget.Line (X, Y + i + 5)-(X + 10, Y + i + 5), &HFFFFFF
                Next
    End Select
End Sub

Private Sub scrollBrushSize_Change()
    lblBrushSize.Caption = scrollBrushSize.value
    picTarget.DrawWidth = scrollBrushSize.value
End Sub

Private Sub scrollRederingQuality_Change()
    lblRederingQuality.Caption = scrollRederingQuality.value
End Sub
