VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcrypt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HYPERCRYPT 1.0 (DEMO)"
   ClientHeight    =   3195
   ClientLeft      =   360
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "HYPERCRYPT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlgb 
      Left            =   4560
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Choose File To Encrypt..."
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   -360
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   4575
   End
   Begin RichTextLib.RichTextBox s 
      Height          =   1005
      Left            =   -360
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1773
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"HYPERCRYPT.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrtime 
      Left            =   2040
      Top             =   4680
   End
   Begin MSComctlLib.ProgressBar prgproga 
      Height          =   345
      Left            =   2780
      TabIndex        =   9
      Top             =   6120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
      Max             =   10
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2820
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4886
            MinWidth        =   4886
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4886
            MinWidth        =   4886
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton order 
      Caption         =   "&ORDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Order Xtreme Tool Kit."
      Top             =   5760
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&HELP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   5
      ToolTipText     =   "Get Help."
      Top             =   5760
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&QUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      MousePointer    =   12  'No Drop
      TabIndex        =   4
      ToolTipText     =   "Quit Hypercrypt."
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   9015
      Begin VB.CommandButton Command3 
         Caption         =   "&BROWSE"
         Height          =   285
         Left            =   7680
         TabIndex        =   17
         Top             =   660
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&ENCRYPT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         ToolTipText     =   "Encrypt"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&DECRYPT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   12
         ToolTipText     =   "Decrypt"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtfile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Top             =   660
         Width           =   4935
      End
      Begin VB.TextBox txtpass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2640
         MaxLength       =   10
         MousePointer    =   3  'I-Beam
         PasswordChar    =   "*"
         TabIndex        =   7
         ToolTipText     =   "Sets Your Password[Numeric]"
         Top             =   1320
         Width           =   4935
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "KEY:(NUMERIC)"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FILENAME:"
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "syntax(drive:\filename.ext)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   3
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "syntax(your password[###])"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   2
         Top             =   1680
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      HYPERCRYPT 1.0(DEMO)(r)"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   705
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "HYPERCRYPT 1.0(DEMO VERSION)PLEASE REGISTER"
      Top             =   360
      Width           =   8940
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmcrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
s.LoadFile txtfile.Text, rtfText
Text1.Text = s.Text
Encrypt Text1, Text1 'Run Encrypt Function
s.Text = Text1.Text
s.SaveFile txtfile.Text, rtfText
status.Panels(1).Text = "PROGRESS:"
tmrtime.Interval = 10
status.Panels(1).Text = "ENCRYPTED"
status.Panels(3).Text = "DONE"
End Sub

Private Sub Command2_Click()
s.LoadFile txtfile.Text, rtfText
Text1.Text = s.Text
Decrypt Text1, Text1 'Run Encrypt Function
s.Text = Text1.Text
s.SaveFile txtfile.Text, rtfText
status.Panels(1).Text = "PROGRESS:"
tmrtime.Interval = 10
status.Panels(1).Text = "ENCRYPTED"
status.Panels(3).Text = "DONE"
End Sub

Private Function Encrypt(Text As String, Output As TextBox)
On Error GoTo Break 'Error Trap
Dim iLen As Long
Dim i As Long
Dim Char()
Dim Out
iLen = Len(Text) 'Get Length of Text
ReDim Char(1 To iLen) 'Redim Char(1 To length of text)
Text = StrReverse(Text) 'Reverse Text
For i = 1 To iLen
    Char(i) = Mid(Text, i, 1) 'Get each bit of text letter by letter
    Char(i) = Asc(Char(i)) Xor txtpass.Text + 7 'Convert to ASCII and Xor
    Char(i) = Chr(Char(i)) 'Convert encrypted ASCII back to character
    Out = Out & Char(i) 'Store Character into out
Next i
Output.Text = Out & "E" 'Return the encrypted text and add "E" onto the end
Exit Function
Break:
MsgBox "Unknown Error!"
End Function

Private Function Decrypt(Text As String, Output As TextBox)
On Error GoTo Break
Dim iLen As Long
Dim i As Long
Dim Char()
Dim Out
Text = Mid(Text, 1, Len(Text) - 1) 'Get Text but without the "E" on the end
iLen = Len(Text) 'Get Length of Text
ReDim Char(1 To iLen) 'Redim
Text = StrReverse(Text) 'Reverse Text
For i = 1 To iLen
    Char(i) = Mid(Text, i, 1) 'Get Character one by one
    Char(i) = Asc(Char(i)) Xor txtpass.Text + 7 'Convert to ASCII and Xor
    Char(i) = Chr(Char(i)) 'COnvert back to character
    Out = Out & Char(i) 'Add encrypted character to Out
Next i
Output.Text = Out
Exit Function
Break:
MsgBox "Unknown Error!"
End Function

Private Sub Command3_Click()
dlgb.ShowOpen
txtfile.Text = dlgb.FileName
End Sub

Private Sub Command5_Click()
frmcrypt.Hide
End Sub

Private Sub tmrtime_Timer()
Static progress As Integer
progress = (progress + 1)
prgproga.Value = progress
If (progress = 10) Then
progress = 0
prgproga.Value = 0
tmrtime.Interval = 0
End If
End Sub
