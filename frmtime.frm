VERSION 5.00
Begin VB.Form frmtime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATE/TIME"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   Icon            =   "frmtime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&CANCEL"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox Lsttime 
      Height          =   1620
      ItemData        =   "frmtime.frx":030A
      Left            =   120
      List            =   "frmtime.frx":030C
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Formats"
      Height          =   2055
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmscript.txttext.SelText = Lsttime.Text
frmtime.Hide
End Sub

Private Sub Command2_Click()
frmtime.Hide
End Sub

Private Sub Form_Load()
Lsttime.AddItem Format(Now, "long time")
Lsttime.AddItem Format(Now, "short time")
Lsttime.AddItem Format(Now, "medium time")
Lsttime.AddItem Format(Now, "general date")
Lsttime.AddItem Format(Now, "long date")
Lsttime.AddItem Format(Now, "medium date")
Lsttime.AddItem Format(Now, "short date")
Lsttime.AddItem (Date)
Lsttime.AddItem Format(Date, "dd - mm - yyyy")
Lsttime.AddItem Format(Date, "dd-mm-yy")
Lsttime.AddItem Format(Date, "dd/mm/yy")
Lsttime.AddItem Format(Date, "dd/mm/yyyy")
Lsttime.AddItem Format(Date, "dd/mm")
Lsttime.AddItem Format(Date, "dd")
Lsttime.AddItem Format(Time, "hh-mm-ss")
Lsttime.AddItem Format(Time, "hh.mm.ss")
Lsttime.AddItem Format(Time, "hh-mm")
End Sub
