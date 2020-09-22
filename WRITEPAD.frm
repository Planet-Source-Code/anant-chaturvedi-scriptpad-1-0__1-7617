VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmscript 
   AutoRedraw      =   -1  'True
   Caption         =   "Script - XTREME SCRIPTPAD 1.0"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "WRITEPAD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Timer tmrproga 
      Left            =   6120
      Top             =   5880
   End
   Begin MSComctlLib.ProgressBar prgproga 
      Height          =   375
      Left            =   5450
      TabIndex        =   1
      Top             =   5880
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   10
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2820
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4886
            MinWidth        =   4886
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8556
            MinWidth        =   8556
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3122
            MinWidth        =   3122
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   0
      TabIndex        =   3
      Top             =   -120
      Width           =   9615
      Begin MSComctlLib.ImageList IL 
         Left            =   240
         Top             =   1320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   24
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":0442
               Key             =   "new"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":0556
               Key             =   "open"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":066A
               Key             =   "save"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":077E
               Key             =   "print"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":0892
               Key             =   "find"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":09A6
               Key             =   "cut"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":0ABA
               Key             =   "copy"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":0BCE
               Key             =   "paste"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":0CE2
               Key             =   "undo"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":0DF6
               Key             =   "bold"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":0F0A
               Key             =   "italic"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":101E
               Key             =   "underline"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":1132
               Key             =   "strike"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":1246
               Key             =   "color"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":1A1A
               Key             =   "left"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":1B2E
               Key             =   "center"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":1C42
               Key             =   "right"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":1D56
               Key             =   "email"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":21AA
               Key             =   "paint"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":25FE
               Key             =   "encryption"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":2A52
               Key             =   "font"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":31A6
               Key             =   "time"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":3CDA
               Key             =   "calc"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "WRITEPAD.frx":3FF6
               Key             =   "bullets"
            EndProperty
         EndProperty
      End
      Begin RichTextLib.RichTextBox txttext 
         Height          =   5535
         Left            =   0
         TabIndex        =   0
         Top             =   480
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   9763
         _Version        =   393217
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"WRITEPAD.frx":47CA
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
      Begin MSComDlg.CommonDialog dlgopen 
         Left            =   240
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "txt"
         Filter          =   "Rich Text Files|*.rtf|Text Files|*.txt|HTML Files|*.html;*.htm|Data Files|*.dat|All Files|*.*"
         InitDir         =   "c:\My Documents"
         MaxFileSize     =   32767
      End
      Begin MSComDlg.CommonDialog dlgcolor 
         Left            =   240
         Top             =   2280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog dlgfont 
         Left            =   240
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "ttf"
      End
      Begin MSComDlg.CommonDialog dlgprint 
         Left            =   240
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog dlgsave 
         Left            =   240
         Top             =   1800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   ".rtf"
         Filter          =   "Rich Text Files|*.rtf|Text Files|*.txt|HTML Files|*.html;*.htm|Data Files|*.dat|All Files|*.*"
      End
      Begin MSComctlLib.Toolbar TB 
         Height          =   330
         Left            =   30
         TabIndex        =   6
         Top             =   180
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         ImageList       =   "IL"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   33
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "new"
               Object.ToolTipText     =   "New"
               ImageKey        =   "new"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "open"
               Object.ToolTipText     =   "Open"
               ImageKey        =   "open"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "save"
               Object.ToolTipText     =   "Save"
               ImageKey        =   "save"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "print"
               Object.ToolTipText     =   "Print"
               ImageKey        =   "print"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "find"
               Object.ToolTipText     =   "Find"
               ImageKey        =   "find"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cut"
               Object.ToolTipText     =   "Cut"
               ImageKey        =   "cut"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "copy"
               Object.ToolTipText     =   "Copy"
               ImageKey        =   "copy"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "paste"
               Object.ToolTipText     =   "Paste"
               ImageKey        =   "paste"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "undo"
               Object.ToolTipText     =   "Undo"
               ImageKey        =   "undo"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "time"
               Object.ToolTipText     =   "Date/Time"
               ImageKey        =   "time"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "bold"
               Object.ToolTipText     =   "Bold"
               ImageKey        =   "bold"
               Style           =   1
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "italic"
               Object.ToolTipText     =   "Italic"
               ImageKey        =   "italic"
               Style           =   1
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "underline"
               Object.ToolTipText     =   "Underline"
               ImageKey        =   "underline"
               Style           =   1
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "strike"
               Object.ToolTipText     =   "Strike Out"
               ImageKey        =   "strike"
               Style           =   1
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "color"
               Object.ToolTipText     =   "Color"
               ImageKey        =   "color"
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "font"
               Object.ToolTipText     =   "Font"
               ImageKey        =   "font"
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "email"
               Object.ToolTipText     =   "e-mail"
               ImageKey        =   "email"
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "calc"
               Object.ToolTipText     =   "Calculator"
               ImageKey        =   "calc"
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "paint"
               Object.ToolTipText     =   "Graphic Editor"
               ImageKey        =   "paint"
            EndProperty
            BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "encryption"
               Object.ToolTipText     =   "Encryption"
               ImageKey        =   "encryption"
            EndProperty
            BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "left"
               Object.ToolTipText     =   "Left Align"
               ImageKey        =   "left"
               Style           =   1
            EndProperty
            BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "center"
               Object.ToolTipText     =   "Center Align"
               ImageKey        =   "center"
               Style           =   1
            EndProperty
            BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "right"
               Object.ToolTipText     =   "Right Align"
               ImageKey        =   "right"
               Style           =   1
            EndProperty
            BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "bullets"
               Object.ToolTipText     =   "Bullets"
               ImageKey        =   "bullets"
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin VB.Label lblt 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbld 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuFILENEW 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnufileopen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnufilesave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnufilesaveas 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnufiled 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileprint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnufiles 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFILEEXIT 
         Caption         =   "&Exit"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuundo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnueditx 
         Caption         =   "-"
      End
      Begin VB.Menu mnueditcut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnueditcopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnueditpaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnueditspecial 
         Caption         =   "Paste &Special"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnueditclear 
         Caption         =   "C&lear"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnueditselect 
         Caption         =   "Select &All"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnueditdel 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu sd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuedittime 
         Caption         =   "&Time/Date"
         Shortcut        =   ^{F5}
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      WindowList      =   -1  'True
      Begin VB.Menu mnuviewtool 
         Caption         =   "&Tool Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuviewstatus 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuviewprogress 
         Caption         =   "&Progress Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnufor 
      Caption         =   "For&mat"
      Begin VB.Menu mnueditfont 
         Caption         =   "&Font"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnueditcase 
         Caption         =   "&Case"
         Begin VB.Menu mnueditucase 
            Caption         =   "&Upper Case"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnueditlcase 
            Caption         =   "&Lower Case"
            Shortcut        =   ^L
         End
      End
      Begin VB.Menu mnueditcolor 
         Caption         =   "Colo&r"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuforno 
         Caption         =   "&No Format"
      End
      Begin VB.Menu mnuscpt 
         Caption         =   "&Scripting"
         Begin VB.Menu mnuforscptsuper 
            Caption         =   "Supe&rScript"
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnuforscptno 
            Caption         =   "&No Scripting"
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuforscptsub 
            Caption         =   "Su&bScript"
            Shortcut        =   ^{F3}
         End
      End
      Begin VB.Menu mnufortxt 
         Caption         =   "&Text Attributes..."
         Begin VB.Menu mnufortxtbold 
            Caption         =   "&Bold"
         End
         Begin VB.Menu mnufortxtitalic 
            Caption         =   "&Italic"
         End
         Begin VB.Menu mnufortxtunder 
            Caption         =   "&Underline"
         End
         Begin VB.Menu mnufortxtstrike 
            Caption         =   "&Strike Out"
         End
      End
   End
   Begin VB.Menu mnufind 
      Caption         =   "&Search"
      Begin VB.Menu mnufindfind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnufindnext 
         Caption         =   "Find &Next"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnutoolsemail 
         Caption         =   "&e-Mail"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnutoolsgraph 
         Caption         =   "&Graphic Editor"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnutoolsencrypt 
         Caption         =   "E&ncrypt"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnutoolscalc 
         Caption         =   "&Calculator"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnutoolscount 
         Caption         =   "&C&ount"
         Begin VB.Menu mnueditlines 
            Caption         =   "&Line Count"
            Shortcut        =   {F9}
         End
         Begin VB.Menu mnueditcount 
            Caption         =   "&Word Count"
            Shortcut        =   {F11}
         End
      End
   End
   Begin VB.Menu mnure 
      Caption         =   "&Relax"
      Begin VB.Menu mnutoolsgragame 
         Caption         =   "&Draw Game"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelptech 
         Caption         =   "&Technical Support"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuhelpabout 
         Caption         =   "&About Xtreme Scriptpad"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmscript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QUERY As Variant
Dim OpenFile As Variant

Private Sub Form_Load()
txttext.TextRTF = ""
QUERY = txttext.TextRTF
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If txttext.TextRTF = QUERY Then
End
Unload Me
ElseIf txttext.TextRTF <> QUERY Then
reply = MsgBox("Do You Want To Save The Changes Made To Current Script " & OpenFile & "?", vbYesNoCancel + vbExclamation)
If reply = vbCancel Then
Cancel = True
ElseIf reply = vbYes Then mnufilesave_Click
Else
Exit Sub
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
Unload Me
End Sub

Private Sub mnueditclear_Click()
txttext.Text = ""
End Sub

Private Sub mnueditcolor_Click()
dlgcolor.ShowColor
txttext.SelColor = dlgcolor.Color
End Sub

Private Sub mnueditcopy_Click()
Clipboard.Clear
Clipboard.SetText txttext.SelText
End Sub

Private Sub mnueditcount_Click()
Dim Position As Long
Dim words As Long
Dim myText As String
    Position = 1
    myText = txttext.Text
    myText = Replace(myText, Chr(13) & Chr(10), " ")
    myText = Replace(myText, Chr(9), " ")
    myText = Trim(myText)
    If Len(myText) > 0 Then words = 1
    Do While Position > 0
        Position = InStr(Position, myText, " ")
        If Position > 0 Then
            words = words + 1
            While Mid(myText, Position, 1) = " "
                Position = Position + 1
            Wend
        End If
    Loop
    MsgBox "You Have Entered " & words & " words"
End Sub

Private Sub mnueditcut_Click()
Clipboard.Clear
Clipboard.SetText txttext.SelText
txttext.SelText = ""
End Sub

Private Sub mnueditdel_Click()
txttext.SelText = ""
End Sub

Private Sub mnueditfont_Click()
On Error GoTo fonterror
dlgfont.Flags = cdlCFApply Or cdlCFBoth Or cdlCFForceFontExist Or cdlCFEffects Or cdlcf
dlgfont.ShowFont
txttext.SelFontName = dlgfont.FontName
txttext.SelBold = dlgfont.FontBold
txttext.SelColor = dlgfont.Color
txttext.SelFontSize = dlgfont.FontSize
txttext.SelItalic = dlgfont.FontItalic
txttext.SelStrikeThru = dlgfont.FontStrikethru
txttext.SelUnderline = dlgfont.FontUnderline
Exit Sub
fonterror:
errornumber = Err.Number
Beep
Select Case errornumber
Case errornumber
MsgBox Err.Description
End Select
End Sub

Private Sub mnueditlcase_Click()
txttext.SelText = LCase(txttext.SelText)
End Sub

Private Sub mnueditlines_Click()
Dim tmpText As String, tmpLine As String
Dim firstChar As Integer, lastChar As Integer
Dim currentLine As Integer
On Error GoTo line_error
firstChar = 1
currentLine = 1
lastChar = InStr(txttext.Text, Chr$(10))
While lastChar > 0
    tmpLine = Format$(currentLine, "000")
    currentLine = currentLine + 1
    firstChar = lastChar + 1
    lastChar = InStr(firstChar, txttext.Text, Chr$(10))
    tmpText = tmpText + tmpLine
Wend
If tmpLine = "" Then
MsgBox "0 Lines Entered"
Else
MsgBox tmpLine & " Lines Entered"
End If
Exit Sub
line_error:
errornumber = Err.Number
Select Case errornumber
Case errornumber
MsgBox Err.Description
End Select
End Sub

Private Sub mnueditpaste_Click()
txttext.SelText = Clipboard.GetText()
End Sub

Private Sub mnueditselect_Click()
txttext.SelStart = 0
txttext.SelLength = Len(txttext.Text)
End Sub

Private Sub mnueditspecial_Click()
SendKeys ("+{insert}")
End Sub

Private Sub mnuedittime_Click()
frmtime.Show
End Sub

Private Sub mnueditucase_Click()
txttext.SelText = UCase(txttext.SelText)
End Sub

Private Sub mnufileexit_Click()
Unload Me
End
End Sub

Private Sub mnuFileNew_Click()
If txttext.TextRTF = QUERY Then
txttext.Text = ""
txttext.SelFontName = txttext.Font
txttext.SelColor = vbBlack
txttext.SelAlignment = 0
txttext.SelBold = False
txttext.SelBullet = False
txttext.SelCharOffset = 0
txttext.SelFontSize = 10
txttext.SelItalic = False
txttext.SelStrikeThru = False
txttext.SelUnderline = False
QUERY = txttext.TextRTF
dlgopen.FileName = ""
OpenFile = ""
frmscript.Caption = "XTREME SCRIPTPAD"
status.Panels(1).Text = "READY"
status.Panels(2).Text = ""
txttext.SetFocus
txttext.MousePointer = 11
tmrproga.Interval = 10
ElseIf txttext.TextRTF <> QUERY Then
reply = MsgBox("Do You Want To Save The Changes Made To Current Script " & OpenFile & "?", vbYesNoCancel + vbExclamation)
If reply = vbCancel Then
Cancel = True
ElseIf reply = vbYes Then
mnufilesave_Click
txttext.Text = ""
txttext.SelFontName = txttext.Font
txttext.SelColor = vbBlack
txttext.SelAlignment = 0
txttext.SelBold = False
txttext.SelBullet = False
txttext.SelCharOffset = 0
txttext.SelFontSize = 10
txttext.SelItalic = False
txttext.SelStrikeThru = False
txttext.SelUnderline = False
QUERY = txttext.TextRTF
dlgopen.FileName = ""
OpenFile = ""
frmscript.Caption = "XTREME SCRIPTPAD"
status.Panels(1).Text = "READY"
status.Panels(2).Text = ""
txttext.SetFocus
txttext.MousePointer = 11
tmrproga.Interval = 10
ElseIf reply = vbNo Then
txttext.Text = ""
txttext.SelFontName = txttext.Font
txttext.SelColor = vbBlack
txttext.SelAlignment = 0
txttext.SelBold = False
txttext.SelBullet = False
txttext.SelCharOffset = 0
txttext.SelFontSize = 10
txttext.SelItalic = False
txttext.SelStrikeThru = False
txttext.SelUnderline = False
QUERY = txttext.TextRTF
dlgopen.FileName = ""
OpenFile = ""
frmscript.Caption = "XTREME SCRIPTPAD"
status.Panels(1).Text = "READY"
status.Panels(2).Text = ""
txttext.SetFocus
txttext.MousePointer = 11
tmrproga.Interval = 10
Else
QUERY = ""
txttext.SelColor = vbBlack
dlgopen.FileName = ""
OpenFile = ""
txttext.Text = ""
frmscript.Caption = "XTREME SCRIPTPAD"
status.Panels(1).Text = "READY"
status.Panels(2).Text = ""
txttext.SetFocus
txttext.MousePointer = 11
tmrproga.Interval = 10
End If
End If
End Sub

Private Sub mnufileopen_Click()
On Error GoTo openerror:
Dim txt As String
Dim FNum As Integer
dlgopen.CancelError = True
On Error GoTo openerror:
dlgopen.Flags = cdlOFNFileMustExist
    If txttext.TextRTF = QUERY Then
    dlgopen.ShowOpen
    If UCase(Right(dlgopen.FileName, 3)) = "RTF" Then
        tmode = rtfRTF
    Else
tmode = rtfText
mnueditselect_Click
mnuforno_Click
SendKeys ("{home}")
End If
    txttext.LoadFile dlgopen.FileName, tmode
    OpenFile = dlgopen.FileName
    QUERY = txttext.TextRTF
    End If
If Not (txttext.TextRTF = QUERY) Then
reply = MsgBox("Do You Want To Save The Changes Made To Current Script " & OpenFile & "?", vbYesNoCancel + vbExclamation)
If reply = vbCancel Then
Cancel = True
ElseIf reply = vbYes Then
mnufilesave_Click
dlgopen.ShowOpen
If UCase(Right(dlgopen.FileName, 3)) = "RTF" Then
    tmode = rtfRTF
Else
tmode = rtfText
mnueditselect_Click
mnuforno_Click
SendKeys ("{home}")
End If
txttext.LoadFile dlgopen.FileName, tmode
QUERY = txttext.TextRTF
OpenFile = dlgopen.FileName
ElseIf reply = vbNo Then
dlgopen.ShowOpen
If UCase(Right(dlgopen.FileName, 3)) = "RTF" Then
    tmode = rtfRTF
Else
mnueditselect_Click
tmode = rtfText
mnuforno_Click
SendKeys ("{home}")
End If
txttext.LoadFile dlgopen.FileName, tmode
QUERY = txttext.TextRTF
OpenFile = dlgopen.FileName
End If
End If
frmscript.Caption = dlgopen.FileTitle + " - XTREME SCRIPTPAD"
status.Panels(1).Text = "OPEN:"
status.Panels(2).Text = dlgopen.FileTitle
status.Panels(3).Text = Format(Now, "long time")
tmrproga.Interval = 10
Exit Sub
openerror:
errornumber = Err.Number
Beep
Select Case errornumber
Case errornumber
End Select
End Sub

Private Sub mnufileprint_Click()
On Error GoTo printerror
dlgprint.Flags = cdlPDNoPageNums
dlgprint.Flags = cdlPDAllPages
dlgprint.Flags = cdlPDPageNums
Printer.Print ""
txttext.SelPrint (Printer.hDC)
printerror:
errornumber = Err.Number
Beep
Select Case errornumber
Case errornumber
MsgBox Err.Description
End Select
End Sub

Private Sub mnufilesave_Click()
On Error GoTo saveerror
Dim FNum As Integer
Dim txt As String

    If OpenFile = "" Then
        mnufilesaveas_Click
        Exit Sub
    End If
On Error GoTo saveerror
        dlgsave.Flags = cdlOFNOverwritePrompt
    If UCase(Right(OpenFile, 3)) = "RTF" Then
        tmode = rtfRTF
    Else
        tmode = rtfText
    End If
   txttext.SaveFile OpenFile, tmode
   QUERY = txttext.TextRTF
   frmscript.Caption = dlgopen.FileTitle + " - XTREME SCRIPTPAD"
status.Panels(1).Text = "SAVING"
status.Panels(2).Text = dlgopen.FileTitle
status.Panels(3).Text = Format(Now, "long time")
tmrproga.Interval = 10
   Exit Sub
saveerror:
errnumber = Err.Number
Select Case errornumber
Case errornumber
End Select
status.Panels(1).Text = "SAVED:"
status.Panels(2).Text = OpenFile
txttext.MousePointer = 11
tmrproga.Interval = 10
End Sub

Private Sub mnufilesaveas_Click()
On Error GoTo saveaserror
Dim txt As String
Dim FNum As Integer

On Error GoTo saveaserror:
    dlgsave.Flags = cdlOFNOverwritePrompt
    dlgsave.ShowSave
    If UCase(Right(dlgsave.FileName, 3)) = "RTF" Then
        tmode = rtfRTF
    Else
        tmode = rtfText
    End If
      txttext.SaveFile dlgsave.FileName, tmode
   OpenFile = dlgsave.FileName
   QUERY = txttext.TextRTF
Caption = dlgopen.FileTitle + " - XTREME SCRIPTPAD"
status.Panels(1).Text = "SAVING"
status.Panels(2).Text = dlgopen.FileTitle
status.Panels(3).Text = Format(Now, "long time")
tmrproga.Interval = 10
    Exit Sub
saveaserror:
Select Case errornumber
Case errornumber
End Select
status.Panels(1).Text = "SAVED:"
status.Panels(2).Text = OpenFile
txttext.MousePointer = 11
tmrproga.Interval = 10
End Sub

Private Sub mnufindfind_Click()
frmfind.Show
End Sub

Private Sub mnufindnext_Click()
mnufindfind_Click
End Sub

Private Sub mnuforno_Click()
txttext.SelFontName = txttext.Font
txttext.SelColor = vbBlack
txttext.SelAlignment = 0
txttext.SelBold = False
txttext.SelBullet = False
txttext.SelCharOffset = 0
txttext.SelFontSize = 10
txttext.SelItalic = False
txttext.SelStrikeThru = False
txttext.SelUnderline = False
End Sub

Private Sub mnuforscptno_Click()
txttext.SelCharOffset = 0
End Sub

Private Sub mnuforscptsub_Click()
txttext.SelCharOffset = -55
End Sub

Private Sub mnuforscptsuper_Click()
txttext.SelCharOffset = 55
End Sub

Private Sub mnufortxtbold_Click()
txttext.SelBold = Not (txttext.SelBold)
mnufortxtbold.Checked = Not (mnufortxtbold.Checked)
txttext.SetFocus
End Sub

Private Sub mnufortxtitalic_Click()
txttext.SelItalic = Not (txttext.SelItalic)
mnufortxtitalic.Checked = Not (mnufortxtitalic.Checked)
End Sub

Private Sub mnufortxtstrike_Click()
txttext.SelStrikeThru = Not (txttext.SelStrikeThru)
mnufortxtstrike.Checked = Not (mnufortxtstrike.Checked)
End Sub

Private Sub mnufortxtunder_Click()
txttext.SelUnderline = Not (txttext.SelUnderline)
mnufortxtunder.Checked = Not (mnufortxtunder.Checked)
End Sub

Private Sub mnuhelpabout_Click()
frmabout.Show
End Sub

Private Sub mnuhelptech_Click()
frmtech.Show
End Sub

Private Sub mnutoolscalc_Click()
frmcalc.Show
End Sub

Private Sub mnutoolsemail_Click()
Form2.Show
End Sub

Private Sub mnutoolsencrypt_Click()
frmcrypt.Show
End Sub

Private Sub mnutoolsgraph_Click()
Form1.Show
End Sub

Private Sub mnutoolsgragame_Click()
frmpaint.Show
End Sub

Private Sub mnuviewprogress_Click()
prgproga.Visible = Not (prgproga.Visible)
mnuviewprogress.Checked = Not (mnuviewprogress.Checked)
End Sub

Private Sub mnuviewstatus_Click()
If status.Visible = True Then
txttext.Height = 5535
txttext.Top = 480
Else
txttext.Top = 180
txttext.Height = txttext.Height + 330
End If
End Sub

Private Sub mnuviewtool_Click()
TB.Visible = Not (TB.Visible)
mnuviewtool.Checked = Not (mnuviewtool.Checked)
If TB.Visible = True Then
txttext.Height = 5535
txttext.Top = 480
Else
txttext.Top = 180
txttext.Height = txttext.Height + 330
End If
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "new"
mnuFileNew_Click
Case "open"
mnufileopen_Click
Case "save"
mnufilesave_Click
Case "print"
mnufileprint_Click
Case "find"
mnufindfind_Click
Case "cut"
mnueditcut_Click
Case "copy"
mnueditcopy_Click
Case "paste"
mnueditpaste_Click
Case "undo"
txttext.SetFocus
SendKeys ("^z")
Case "time"
frmtime.Show
Case "bold"
txttext.SelBold = Not (txttext.SelBold)
txttext.SetFocus
Case "italic"
txttext.SelItalic = Not (txttext.SelItalic)
txttext.SetFocus
Case "underline"
txttext.SelUnderline = Not (txttext.SelUnderline)
txttext.SetFocus
Case "strike"
txttext.SelStrikeThru = Not (txttext.SelStrikeThru)
txttext.SetFocus
Case "color"
dlgcolor.ShowColor
txttext.SelColor = dlgcolor.Color
txttext.SetFocus
Case "font"
mnueditfont_Click
Case "email"
mnutoolsemail_Click
Case "calc"
mnutoolscalc_Click
Case "paint"
mnutoolsgraph_Click
Case "encryption"
mnutoolsencrypt_Click
Case "left"
txttext.SelAlignment = 0
txttext.SetFocus
Case "center"
txttext.SelAlignment = 2
txttext.SetFocus
Case "right"
txttext.SelAlignment = 1
txttext.SetFocus
Case "bullets"
txttext.SelBullet = Not (txttext.SelBullet)
txttext.SetFocus
End Select
End Sub

Private Sub Timer1_Timer()
status.Panels(3).Alignment = sbrRight
status.Panels(3).Text = Format(Now, "long time")
End Sub

Private Sub tmrproga_Timer()
Static progress As Integer
progress = (progress + 1)
prgproga.Value = progress
If (progress = 10) Then
progress = 0
prgproga.Value = 0
tmrproga.Interval = 0
txttext.MousePointer = 0
End If
End Sub

Private Sub txttext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
prgproga.Top = status.Top
End Sub

Private Sub txttext_SelChange()
If txttext.SelAlignment = 0 Then
TB.Buttons.Item(29).Value = tbrPressed
TB.Buttons.Item(30).Value = tbrUnpressed
TB.Buttons.Item(31).Value = tbrUnpressed
Else
TB.Buttons.Item(29).Value = tbrUnpressed
End If
If txttext.SelAlignment = 2 Then
TB.Buttons.Item(30).Value = tbrPressed
TB.Buttons.Item(29).Value = tbrUnpressed
TB.Buttons.Item(31).Value = tbrUnpressed
Else
TB.Buttons.Item(30).Value = tbrUnpressed
End If
If txttext.SelAlignment = 1 Then
TB.Buttons.Item(31).Value = tbrPressed
TB.Buttons.Item(29).Value = tbrUnpressed
TB.Buttons.Item(30).Value = tbrUnpressed
Else
TB.Buttons.Item(31).Value = tbrUnpressed
End If
If txttext.SelBold = True Then
mnufortxtbold.Checked = True
TB.Buttons.Item(16).Value = tbrPressed
Else
mnufortxtbold.Checked = False
TB.Buttons.Item(16).Value = tbrUnpressed
End If
If txttext.SelItalic = True Then
mnufortxtitalic.Checked = True
TB.Buttons.Item(17).Value = tbrPressed
Else
mnufortxtitalic.Checked = False
TB.Buttons.Item(17).Value = tbrUnpressed
End If
If txttext.SelUnderline = True Then
mnufortxtunder.Checked = True
TB.Buttons.Item(18).Value = tbrPressed
Else
mnufortxtunder.Checked = False
TB.Buttons.Item(18).Value = tbrUnpressed
End If
If txttext.SelStrikeThru = True Then
mnufortxtstrike.Checked = True
TB.Buttons.Item(19).Value = tbrPressed
Else
mnufortxtstrike.Checked = False
TB.Buttons.Item(19).Value = tbrUnpressed
End If
If txttext.SelBullet = True Then
TB.Buttons.Item(33).Value = tbrPressed
Else
TB.Buttons.Item(33).Value = tbrUnpressed
End If
End Sub
