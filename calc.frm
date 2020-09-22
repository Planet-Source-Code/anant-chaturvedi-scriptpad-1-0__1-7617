VERSION 5.00
Begin VB.Form frmcalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XTREME CALCULATOR"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4215
   Icon            =   "calc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton sqrt 
      Caption         =   "_/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   20
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton plusminus 
      Caption         =   "&+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   18
      ToolTipText     =   "CLEAR"
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton over 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   17
      ToolTipText     =   "SQUARE ROOT(_/)"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton dotbttn 
      Caption         =   "&."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   16
      ToolTipText     =   "DECIMAL(.)"
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton equals 
      Caption         =   "&="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   15
      ToolTipText     =   "OUTPUT(=)"
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton div 
      Caption         =   "&/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   14
      ToolTipText     =   "DIVISION(/)"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton clearbttn 
      Caption         =   "&C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      ToolTipText     =   "CLEAR"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton minus 
      Caption         =   "&-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      ToolTipText     =   "SUBTRACTION(-)"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton times 
      Caption         =   "&*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      ToolTipText     =   "MULTIPLICATION(*)"
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton plus 
      Caption         =   "&+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      ToolTipText     =   "PLUS(+)"
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton digits 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   " 0 "
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton digits 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   1560
      TabIndex        =   8
      ToolTipText     =   " 9 "
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton digits 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   840
      TabIndex        =   7
      ToolTipText     =   " 8 "
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton digits 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   " 7 "
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton digits 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   1560
      TabIndex        =   5
      ToolTipText     =   " 6 "
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton digits 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   840
      TabIndex        =   4
      ToolTipText     =   " 5 "
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton digits 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   " 4 "
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton digits 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   " 3 "
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton digits 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   1
      ToolTipText     =   " 2 "
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton digits 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " 1 "
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label display 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   3975
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufilesend 
         Caption         =   "&Send To Current Document"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmcalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim operand1 As Double, operand2 As Double
Dim operator As String
Dim cleardisplay As Boolean

Private Sub clearbttn_Click()
display.Caption = ""
End Sub

Private Sub digits_Click(Index As Integer)
If cleardisplay Then
display.Caption = ""
cleardisplay = False
End If
display.Caption = display.Caption + digits(Index).Caption
End Sub

Private Sub div_Click()
operand1 = Val(display.Caption)
operator = "/"
display.Caption = ""
End Sub

Private Sub dotbttn_Click()
If InStr(display.Caption, ".") Then
Exit Sub
Else
display.Caption = display.Caption + "."
End If
End Sub

Private Sub equals_Click()
Dim result As Double
operand2 = Val(display.Caption)
If operator = "+" Then result = operand1 + operand2
If operator = "-" Then result = operand1 - operand2
If operator = "*" Then result = operand1 * operand2
If operator = "/" And operand2 <> 0 Then result = operand1 / operand2
display.Caption = result
End Sub

Private Sub minus_Click()
operand1 = Val(display.Caption)
operator = "-"
display.Caption = ""
End Sub

Private Sub mnufilesend_Click()
frmscript.txttext.SelText = display.Caption
End Sub

Private Sub over_Click()
If Val(display.Caption) <> 0 Then display.Caption = 1 / Val(display.Caption)
End Sub

Private Sub plus_Click()
operand1 = Val(display.Caption)
operator = "+"
display.Caption = ""
End Sub

Private Sub plusminus_Click()
display.Caption = -Val(display.Caption)
End Sub

Private Sub times_Click()
operand1 = Val(display.Caption)
operator = "*"
display.Caption = ""
End Sub
