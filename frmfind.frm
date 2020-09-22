VERSION 5.00
Begin VB.Form frmfind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FIND..."
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "frmfind.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check2 
      Caption         =   "Match Whole Word"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton ReplaceAllButton 
      Caption         =   "Replace &All"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton ReplaceButton 
      Caption         =   "&Replace"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton FindNextButton 
      Caption         =   "Find &Next"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton FindButton 
      Caption         =   "&Find"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Match Case"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmfind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Position As Integer

Private Sub Command5_Click()
frmfind.Hide
End Sub

Private Sub FindButton_Click()
Dim FindFlags As Integer
    On Error GoTo error_2
    Position = 0
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    Position = frmscript.txttext.Find(Text1.Text, Position + 1, , FindFlags)
    If Position >= 0 Then
        ReplaceButton.Enabled = True
        ReplaceAllButton.Enabled = True
    Else
        MsgBox "SCRIPTPAD Could Not Find " & Text1.Text
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
    End If
    Exit Sub
error_2:
End Sub

Private Sub FindNextButton_Click()
Dim FindFlags
On Error GoTo error
FindFlags = Check1.Value * 4 + Check2.Value * 2
Position = frmscript.txttext.Find(Text1.Text, Position + 1, , FindFlags)
If Position > 0 Then
Else
    MsgBox "SCRIPTPAD Could Not Find " & Text1.Text
    ReplaceButton.Enabled = False
    ReplaceAllButton.Enabled = False
End If
Exit Sub
error:
End Sub

Private Sub replaceallbutton_Click()
Dim FindFlags As Integer

    FindFlags = Check1.Value * 4 + Check2.Value * 2
   frmscript.txttext.SelText = Text2.Text
    Position = frmscript.txttext.Find(Text1.Text, Position + 1, , FindFlags)
    While Position > 0
        frmscript.txttext.SelText = Text2.Text
        Position = frmscript.txttext.Find(Text1.Text, Position + 1, , FindFlags)
    Wend
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
        MsgBox "Replacing Compleated successfully "
End Sub

Private Sub replacebutton_Click()
Dim FindFlags As Integer

    frmscript.txttext.SelText = Text2.Text
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    Position = frmscript.txttext.Find(Text1.Text, Position + 1, , FindFlags)
    If Position > 0 Then
        frmscript.txttext.SetFocus
    Else
        MsgBox "String not found"
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
    End If
End Sub
