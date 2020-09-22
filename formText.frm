VERSION 5.00
Begin VB.Form formText 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Costume Text"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdUse 
      Caption         =   "Use"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtText 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "formText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdClear_Click()
    txtText = ""
End Sub

Private Sub cmdUse_Click()
    If txtText <> "" Then
        strText = txtText
        intCharindex = 0
    Else
        strText = ""
        intCharindex = 64
    End If
    Me.Hide
End Sub
