VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   Caption         =   "XTREMe-MAIL"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8700
   Icon            =   "smtpmail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   2400
      TabIndex        =   4
      Text            =   "202.144.13.82"
      Top             =   1500
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Text            =   "Aditya Chaturvedi"
      Top             =   180
      Width           =   6135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4680
      TabIndex        =   8
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   3195
      Index           =   6
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2160
      Width           =   9555
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   2400
      TabIndex        =   5
      Top             =   1830
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   3
      Top             =   1170
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Text            =   "aditya_chat_2000@yahoo.com"
      Top             =   510
      Width           =   6135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send "
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   7
      Top             =   5400
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6240
      Top             =   4500
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   5880
      Width           =   1425
   End
   Begin VB.Label StatusLabel 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "E-Mail Server Name"
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   14
      Top             =   1530
      Width           =   2145
   End
   Begin VB.Label Label1 
      Caption         =   "Recipient's Name"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   13
      Top             =   870
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   12
      Top             =   210
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Subject"
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   11
      Top             =   1830
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "To "
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   10
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   9
      Top             =   540
      Width           =   2055
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufilenew 
         Caption         =   "&New"
         Begin VB.Menu mnufilemail 
            Caption         =   "&Mail"
            Shortcut        =   ^M
         End
         Begin VB.Menu mnufileaccount 
            Caption         =   "&Account"
            Shortcut        =   ^A
         End
      End
      Begin VB.Menu mnufiled 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileopen 
         Caption         =   "&Open Account"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnufiler 
         Caption         =   "-"
      End
      Begin VB.Menu mnufilesave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnufilein 
         Caption         =   "Save &In"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnufiles 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "E&xit"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'General form declarations
Dim Response As String

Sub SendEmail(MailServerName As String, SenderName As String, SenderEmailAddress As String, RecipientName As String, RecipientEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String)
    
    Dim Data1 As String, Data2 As String
    Dim Data3 As String, Data4 As String
    Dim Data5 As String, Data6 As String
    Dim Data7 As String, Data8 As String
    Dim CurrentDate As String
    Dim TimeDifference As String
    
    'Set the Winsock control's local port to 0, because otherwise
    'you may not be able to send more than one e-mail message
    'every time the program runs
    Winsock1.LocalPort = 0
    
    'Start composing the required data strings, but first check
    'if the Winsock socket is closed
    If Winsock1.State = sckClosed Then
        'Compose the current date and time string
        TimeDifference = " -200"
        CurrentDate = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & TimeDifference
        'Set the program name used to send this e-mail message (you can
        'put your program name here)
        AppName = "X-Mailer: " + "My Mail Program V1.0" + Chr(13) + Chr(10)
        'Set the e-mail address of the sender
        Data1 = "mail from:" + Chr(32) + SenderEmailAddress + Chr(13) + Chr(10)
        'Set the e-mail address of the recipient
        Data2 = "rcpt to:" + Chr(32) + RecipientEmailAddress + Chr(13) + Chr(10)
        'Set the date string
        Data3 = "Date:" + Chr(32) + CurrentDate + Chr(13) + Chr(10)
        'Set the name of the sender
        Data4 = "From:" + Chr(32) + SenderName + Chr(13) + Chr(10)
        'Set the name of the recipient
        Data5 = "To:" + Chr(32) + Text1(2) + Chr(13) + Chr(10)
        'Set the subject of the E-Mail message
        Data6 = "Subject:" + Chr(32) + EmailSubject + Chr(13) + Chr(10)
        'Set the E-mail message body string
        Data7 = EmailBodyOfMessage + Chr(13) + Chr(10)
        'Combine the whole string for proper SMTP syntax
        Data8 = Data4 + Data3 + AppName + Data5 + Data6
    
        'Set the Winsock protocol
        Winsock1.Protocol = sckTCPProtocol
        'Set the remote host name (of SMTP server)
        Winsock1.RemoteHost = MailServerName
        'Set the SMTP Port to the default port 25
        Winsock1.RemotePort = 25
        
        'Start the connection
        Winsock1.Connect
        'Wait for response from the remote host
        WaitForResponse ("220")
        
        'Report status
        StatusLabel.Caption = "Connecting...."
        StatusLabel.Refresh
        
        'Send your computer name or company name
        Winsock1.SendData ("HELO mycomputername" + Chr(13) + Chr(10))
        'Wait for response from the remote host
        WaitForResponse ("250")
    
        'Update status
        StatusLabel.Caption = "Connected"
        StatusLabel.Refresh
    
        'Send the first string
        Winsock1.SendData (Data1)
    
        'Update status
        StatusLabel.Caption = "Sending Message"
        StatusLabel.Refresh
    
        'Wait for response from the remote host
        WaitForResponse ("250")
    
        'Send the second string
        Winsock1.SendData (Data2)
    
        'Wait for response from the remote host
        WaitForResponse ("250")
    
        'Tell the SMTP server that you want to send data now
        Winsock1.SendData ("data" + Chr(13) + Chr(10))
        
        'Wait for response from the remote host
        WaitForResponse ("354")
    
        'Send the data
        Winsock1.SendData (Data8 + Chr(13) + Chr(10))
        Winsock1.SendData (Data7 + Chr(13) + Chr(10))
        Winsock1.SendData ("." + Chr(13) + Chr(10))
    
        'Wait for response from the remote host
        WaitForResponse ("250")
    
        'Send quitting acknowledgment
        Winsock1.SendData ("quit" + Chr(13) + Chr(10))
        
        'Update status
        StatusLabel.Caption = "Disconnecting"
        StatusLabel.Refresh
    
        'Wait for response from the remote host
        WaitForResponse ("221")
    
        'Close the connection
        Winsock1.Close
    Else
        'Report error
        MsgBox (Str(Winsock1.State))
    End If
   
End Sub

Sub WaitForResponse(ResponseCode As String)
    
    Dim Start As Single
    Dim TimeToWait As Single

    Start = Timer
    'Start a loop checking for response from SMTP host
    While Len(Response) = 0
        TimeToWait = Start - Timer
        DoEvents
        'If TimeToWait expires, report timeout error
        If TimeToWait > 50 Then
            MsgBox "SMTP timeout error, no response received", 64, App.Title
            Exit Sub
        End If
    Wend
    While Left(Response, 3) <> ResponseCode
        DoEvents
        If TimeToWait > 50 Then
            'Report error if incorrect code is received
            MsgBox "SMTP error, improper response code received!" + Chr(10) + "Correct code is: " + ResponseCode + ", Code received: " + Response, 64, App.Title
            Exit Sub
        End If
    Wend

    'Set response to nothing
    Response = ""

End Sub

Private Sub Command1_Click()
    
    'Call the SendEmail procedure and pass the arguments: MailServerName, SenderName, SenderEmailAddress, RecipientName, RecipientEmailAddress, EmailSubject, EmailBodyOfMessage)
    SendEmail Text1(4), Text1(0), Text1(1), Text1(2), Text1(3), Text1(5), Text1(6)
    
    'Update status
    StatusLabel.Caption = "Mail Sent"
    StatusLabel.Refresh
    
End Sub

Private Sub Command2_Click()
    
   Form2.Hide
       
End Sub

Private Sub mnufileaccount_Click()
frmnewa.Show
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    'Check for response from the remote host
    Winsock1.GetData Response

End Sub
