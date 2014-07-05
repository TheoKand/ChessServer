VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Σκάκι - CLIENT"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   7110
   Begin VB.Timer tmrEnableSend 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   2100
      Top             =   7920
   End
   Begin VB.Timer tmrIncoming 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1560
      Top             =   7920
   End
   Begin VB.Timer tmrOutgoing 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   990
      Top             =   7920
   End
   Begin MSWinsockLib.Winsock SOCKET 
      Left            =   270
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraLogin 
      Caption         =   "Ρυθμίσεις"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   60
      TabIndex        =   14
      Top             =   30
      Width           =   6975
      Begin VB.ComboBox cboNickname 
         Height          =   315
         Left            =   1988
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   630
         Width           =   3135
      End
      Begin VB.TextBox txtGamePort 
         Height          =   285
         Left            =   11460
         TabIndex        =   26
         Top             =   660
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Έξοδος"
         Height          =   405
         Index           =   1
         Left            =   3600
         TabIndex        =   13
         Top             =   1500
         Width           =   1485
      End
      Begin VB.CommandButton cmdLogin 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Σύνδεση"
         Height          =   405
         Left            =   2010
         TabIndex        =   12
         Top             =   1500
         Width           =   1485
      End
      Begin VB.TextBox txtNickname 
         Height          =   285
         Left            =   720
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1350
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   8040
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1620
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.TextBox txtLogin 
         Height          =   285
         Left            =   8040
         TabIndex        =   9
         Top             =   1260
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.TextBox txtServerPort 
         Height          =   285
         Left            =   8040
         TabIndex        =   8
         Top             =   660
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtServerIP 
         Height          =   285
         Left            =   8040
         ScrollBars      =   1  'Horizontal
         TabIndex        =   7
         Top             =   330
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Game Port:"
         Height          =   195
         Index           =   5
         Left            =   10320
         TabIndex        =   25
         Top             =   690
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nickname:"
         Height          =   195
         Index           =   4
         Left            =   3098
         TabIndex        =   19
         Top             =   390
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Index           =   3
         Left            =   6840
         TabIndex        =   18
         Top             =   1650
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Login:"
         Height          =   195
         Index           =   2
         Left            =   6840
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Server Port:"
         Height          =   195
         Index           =   1
         Left            =   6840
         TabIndex        =   16
         Top             =   690
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Server IP:"
         Height          =   195
         Index           =   0
         Left            =   6840
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin VB.Frame fraMain 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   60
      TabIndex        =   20
      Top             =   30
      Visible         =   0   'False
      Width           =   6975
      Begin VB.ListBox lstUsers 
         Height          =   2790
         Left            =   150
         TabIndex        =   24
         Top             =   480
         Width           =   2565
      End
      Begin VB.CommandButton cmdConnectToGame 
         Caption         =   "Σύνδεση"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5070
         TabIndex        =   5
         Top             =   2940
         Width           =   1755
      End
      Begin VB.CommandButton cmdHostGame 
         Caption         =   "Νέα Παρτίδα"
         Height          =   345
         Left            =   2790
         TabIndex        =   4
         Top             =   2940
         Width           =   1755
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Έξοδος"
         Height          =   375
         Index           =   0
         Left            =   5340
         TabIndex        =   2
         Top             =   7920
         Width           =   1485
      End
      Begin VB.CommandButton cmdDisconnect 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Αποσύνδεση"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   3810
         TabIndex        =   1
         Top             =   7920
         Width           =   1485
      End
      Begin VB.ComboBox cboChat 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   7470
         Width           =   6105
      End
      Begin VB.PictureBox rtbChat 
         Height          =   3975
         Left            =   120
         ScaleHeight     =   3915
         ScaleWidth      =   6645
         TabIndex        =   6
         Top             =   3390
         Width           =   6705
      End
      Begin VB.ListBox lstGames 
         Height          =   2400
         Left            =   2790
         TabIndex        =   3
         Top             =   480
         Width           =   4035
      End
      Begin VB.TextBox txtDebug 
         Height          =   3705
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   3540
         Visible         =   0   'False
         Width           =   6405
      End
      Begin VB.Label Label2 
         Caption         =   "Chat: "
         Height          =   285
         Left            =   180
         TabIndex        =   23
         Top             =   7500
         Width           =   1335
      End
      Begin VB.Label lblActiveGames 
         Alignment       =   2  'Center
         Caption         =   "Ενεργές Παρτίδες (0)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2790
         TabIndex        =   22
         Top             =   270
         Width           =   4005
      End
      Begin VB.Label lblUsersOnline 
         Alignment       =   2  'Center
         Caption         =   "Χρήστες Online (0)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   270
         Width           =   2580
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFormUnloaded As Boolean

Private Sub cboChat_Click()
    If cboChat.ListIndex = -1 Then Exit Sub
    'Client is sending a chat message to the server lobby
    SendCommand FormatCommand(mtLobbySay, cboChat.List(cboChat.ListIndex))
    cboChat.ListIndex = -1
        
    cboChat.Enabled = False
    tmrEnableSend.Enabled = True
End Sub

Private Sub cmdClose_Click(Index As Integer)
    Dim f As Form
    For Each f In Forms
        Unload f
    Next
    End
End Sub

'Attempt to connect to an open game (game that is waiting for opponent)
Private Sub cmdConnectToGame_Click()
    Dim ServerIP As String, GameDesc As String, Nickname As String
    If lstGames.ListIndex = -1 Then Exit Sub
    
    'The games listbox contains for each game the following info:
    'Nickname of hosting client, IP of hosting client
    GameDesc = lstGames.List(lstGames.ListIndex)
    Nickname = Split(GameDesc, " vs ", -1, vbTextCompare)(0)
        
    If FormIsLoaded("frmGame") Then
        If MyMsgbox("Έχετε ήδη ξεκινήσει μια παρτίδα." & vbCrLf & vbCrLf & "Θέλετε να την ακυρώσετε ωστε να συνδεθείτε στην παρτίδα " & vbCrLf & vbCrLf & Split(GameDesc, ARGUMENT_SEPERATOR)(0), vbQuestion + vbYesNo) = vbNo Then
            frmGame.WindowState = vbNormal
            Exit Sub
        Else
            gUserState = usIdle
            Unload frmGame: Set frmGame = Nothing
        End If

    End If
        
    frmGame.Init False, Nickname
    
End Sub

'Disconnect from server
Private Sub cmdDisconnect_Click()
    gConnectionState = csDisconnected
    SOCKET.Close
    Caption = "Σκάκι - CLIENT"
    fraLogin.Visible = True
    fraMain.Visible = False
    
    If FormIsLoaded("frmGame") Then Unload frmGame
    
End Sub

'Start an open game where one other client can connect to play
Private Sub cmdHostGame_Click()
    If FormIsLoaded("frmGame") Then
        MyMsgbox "Έχετε ήδη ξεκινήσει μια παρτίδα.", vbInformation
        frmGame.WindowState = vbNormal
        SetFormFocus frmGame
        Exit Sub
    End If
    
    'Check if there are already some open games that the client can play at
    Dim Result As String
    Result = OpenGames
    If Result <> "" Then
        If MyMsgbox("Υπάρχουν ήδη παρτίδες που περιμένουν αντίπαλο, και στις οποίες μπορείτε" & vbCrLf & "να συνδεθείτε ωστε να παίξετε σκάκι:" & vbCrLf & vbCrLf & Result & vbCrLf & vbCrLf & "Σίγουρα θέλετε να ξεκινήσετε νέα παρτίδα;", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    frmGame.Init True
End Sub

'Returns the open games
Private Function OpenGames() As String
    Dim i As Integer, GameDesc As String
    For i = 0 To lstGames.ListCount - 1
        GameDesc = Split(lstGames.List(i), ARGUMENT_SEPERATOR)(0)
        If InStr(1, GameDesc, " (κενή θέση)", vbTextCompare) Then
            If OpenGames <> "" Then OpenGames = OpenGames & vbCrLf
            OpenGames = OpenGames & GameDesc
        End If
    Next
End Function


Private Sub Form_Resize()
    If FormIsLoaded("frmGame") Then
        If WindowState = vbMinimized Then frmGame.WindowState = vbMinimized
        If WindowState = vbNormal Then
            frmGame.WindowState = vbNormal
            PositionGameForm
        End If
    End If
End Sub

Private Sub lstGames_Click()
    If lstGames.ListIndex = -1 Then
        cmdConnectToGame.Enabled = False
        Exit Sub
    End If
    
    Dim TheUsername As String, GameDesc As String
    TheUsername = GetUserNickname & " vs"
    GameDesc = lstGames.List(lstGames.ListIndex)
    If (Left(GameDesc, Len(TheUsername)) = TheUsername) Or (InStr(1, GameDesc, "(κενή θέση)", 1) = 0) Then
        cmdConnectToGame.Enabled = False
        Exit Sub
    End If
    
    cmdConnectToGame.Enabled = True
    
End Sub

'This event is fired when the server closes the connection
Private Sub SOCKET_Close()
    
    If FormIsLoaded("frmGame") Then Unload frmGame
    
    SOCKET.Close
    
    gConnectionState = csDisconnected
    Caption = "Σκάκι - CLIENT"
    
    MyMsgbox "Έγινε αποσύνδεση απο τον Server", vbInformation
    
    tmrOutgoing.Enabled = False
    tmrIncoming.Enabled = False
    
    fraMain.Visible = False
    fraLogin.Visible = True
    
End Sub

'Incoming tcp/ip message
Private Sub SOCKET_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    'Read the message from the socket
    SOCKET.GetData sData, vbString
    LogThis "RAW INCOMING:" & sData
    
    'Put it in the "dirty" queue
    gQueue = gQueue & sData
    
    'Process the queue, so that if complete messages are inside,
    'they must be processed and removed from the queue
    gQueue = TrimIncomingCommands(gQueue)
    
End Sub

'Remove complete messages from the queue and process them
Private Function TrimIncomingCommands(ByVal UserQueue As String) As String
    Dim CommandArray As Variant, i As Integer, Command As String, NewQueue As String
    Dim CommandToReceive As String
    
    CommandArray = Split(UserQueue, COMMAND_SEPERATOR)
    For i = 0 To UBound(CommandArray)
        Command = ""
        If CommandArray(i) <> "" Then
            If Left(CommandArray(i), 1) = "<" And Right(CommandArray(i), 1) = ">" Then
                Command = Mid(CommandArray(i), 2)
                Command = Left(Command, Len(Command) - 1)
            Else
                NewQueue = NewQueue & CommandArray(i)
                If i < UBound(CommandArray) Then NewQueue = NewQueue & COMMAND_SEPERATOR
            End If
            If Command <> "" Then
                'Process this message
                CommandToReceive = Command
            End If
            
        End If
    Next
    TrimIncomingCommands = NewQueue
    
    If CommandToReceive <> "" Then ReceiveCommand CommandToReceive
    
End Function

Private Sub tmrEnableSend_Timer()
    tmrEnableSend.Enabled = False
    cboChat.Enabled = True
End Sub

'This timer processes incoming messagess
Private Sub tmrIncoming_Timer()
    Dim TheCommand As String
    If gIncoming.Count > 0 Then
        If SOCKET.State = sckConnected Then
            TheCommand = gIncoming(1)
            gIncoming.Remove 1
            ProcessCommand TheCommand
        End If
    End If
End Sub

'This timer is sneding outgoing tcp/ip messages
'(not used finally, messages are sent out immediately without an outgoing queue)
Private Sub tmrOutgoing_Timer()
'    Dim TheCommand As String
'    If gOutgoing.Count > 0 Then
'        TheCommand = gOutgoing(1)
'        gOutgoing.Remove 1
'        SendToSocket TheCommand
'    End If
End Sub

'Attemp to connect to the chess server
Private Sub cmdLogin_Click()
    Dim OK As Boolean
    OK = True
    'Validation
    If Trim(txtServerIP) = "" Then OK = False
    If Trim(txtServerPort) = "" Then OK = False
    If Trim(txtLogin) = "" Then OK = False
    If Trim(txtPassword) = "" Then OK = False
    If Trim(cboNickname.Text) = "" Then OK = False
    If Trim(txtGamePort) = "" Then OK = False
    
    If OK = False Then
        MyMsgbox "Συμπληρώστε όλα τα πεδία και πατήστε ΣΥΝΔΕΣΗ", vbInformation
        Exit Sub
    End If
    
    Connect
    
    cboChat.Enabled = True
    cmdHostGame.Enabled = True
    cmdConnectToGame.Enabled = False
    
    rtbChat.Text = ""
    
    Me.Enabled = False
    Load frmWait
    frmWait.Move Me.Left + Me.Width / 2 - frmWait.Width / 2, Me.Top + Me.Height / 2 - frmWait.Height / 2
    frmWait.Show
    
    frmWait.lblStatus = "Γίνεται επικοινωνία με τον server..."
    
    Dim StartTimer As Single
    StartTimer = Timer
    'Wait until a connection, a failure or a timeout
    Do Until (gConnectionState = csDisconnected) Or (gConnectionState = csLoging)
        DoEvents
        If Timer > StartTimer + 45 Then ' Timeout is set to 45 sec
            gConnectionState = csDisconnected
            MyMsgbox "O server δεν είναι διαθέσιμος.", vbInformation
            Exit Do
        End If
        If gShutdown Then Exit Do
        If mFormUnloaded Then Exit Do
    Loop
    
    
    If gConnectionState = csLoging Then
        'The server replied. The login begins. Send login,password,nick and client version
        SendCommand FormatCommand(mtID, txtLogin & ARGUMENT_SEPERATOR & txtPassword & ARGUMENT_SEPERATOR & GetUserNickname & ARGUMENT_SEPERATOR & App.Major & "." & App.Minor & "." & App.Revision)
        
        frmWait.lblStatus = "Γίνεται πιστοποίηση του χρήστη..."
        
        'Wait until the login is complete
        Do Until gConnectionState <> csLoging
            DoEvents
            If gShutdown Then Exit Do
            If mFormUnloaded Then Exit Do
        Loop
        If gConnectionState = csConnected Then
            Caption = txtServerIP & ":" & txtServerPort & " - Συνδεδεμένος ως " & GetUserNickname
            fraLogin.Visible = False
            fraMain.Visible = True
            gUserState = usIdle
        End If
    
    End If
    
    If gConnectionState = csDisconnected Then SOCKET.Close
    Me.Enabled = True
    Unload frmWait: Set frmWait = Nothing

    

End Sub

'Read settings from registry.
Private Sub GetSettings()
'    txtServerIP = GetSetting(App.Title, "Settings", "txtServerIP", "194.30.227.16")
'    txtServerPort = GetSetting(App.Title, "Settings", "txtServerPort", "30200")
'    txtLogin = GetSetting(App.Title, "Settings", "txtLogin", "theok")
'    txtPassword = GetSetting(App.Title, "Settings", "txtPassword", "theok")
'    txtGamePort = GetSetting(App.Title, "Settings", "txtGamePort", "30201")

    txtServerIP = "127.0.0.1"
    'txtServerIP = "194.30.227.16"
    
    txtServerPort = "30200"
    txtLogin = "theo"
    txtPassword = "theo"
    txtGamePort = "30201"

End Sub

''SAve settings to registry
'Private Sub SaveSettings()
'    SaveSetting App.Title, "Settings", "txtServerIP", txtServerIP
'    SaveSetting App.Title, "Settings", "txtServerPort", txtServerPort
'    SaveSetting App.Title, "Settings", "txtLogin", txtLogin
'    SaveSetting App.Title, "Settings", "txtPassword", txtPassword
'    SaveSetting App.Title, "Settings", "txtNickname", cboNickname.ListIndex
'    SaveSetting App.Title, "Settings", "txtGamePort", txtGamePort
'End Sub
'
'Private Sub EraseSettings()
'    SaveSetting App.Title, "Settings", "txtServerIP", ""
'    SaveSetting App.Title, "Settings", "txtServerPort", ""
'    SaveSetting App.Title, "Settings", "txtNickname", ""
'    SaveSetting App.Title, "Settings", "txtGamePort", ""
'End Sub

Private Sub Form_Load()
        
    'If DateDiff("d", DateSerial(2004, 7, 25), Date) > 15 Then End
        
    FixNicks
    
    mFormUnloaded = False
    GetSettings
    InitLobbyChat cboChat
    gConnectionState = csDisconnected
    

End Sub

Private Sub FixNicks()
    Dim ar, i As Integer
    ar = GetNicknamesArray
    For i = 0 To UBound(ar)
        If Trim(ar(i)) <> "" Then cboNickname.AddItem Trim(ar(i))
    Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'SaveSettings
    mFormUnloaded = True
    gShutdown = True
    If FormIsLoaded("frmGame") Then Unload frmGame: Set frmGame = Nothing
    If FormIsLoaded("frmWait") Then Unload frmWait: Set frmWait = Nothing
    End
End Sub

'This event runs when a successful connection is established with the server
Private Sub SOCKET_Connect()
    gConnectionState = csLoging
    frmWait.lblStatus = ""
    
    'Set gOutgoing = Nothing
    Set gIncoming = Nothing
    
    tmrOutgoing.Enabled = True
    tmrIncoming.Enabled = True
    
End Sub



'Check if the nickname is allowed
Private Sub txtNickname_Validate(Cancel As Boolean)
    
    Dim IsValidNickname As Boolean
    
    IsValidNickname = True
    If Trim(txtNickname) = "" Then Exit Sub
    
    If InStr(1, txtNickname, COMMAND_SEPERATOR) Then IsValidNickname = False
    If InStr(1, txtNickname, MSG_SEPERATOR) Then IsValidNickname = False
    If InStr(1, txtNickname, ARGUMENT_SEPERATOR) Then IsValidNickname = False
    If InStr(1, txtNickname, ARGUMENT_SEPERATOR2) Then IsValidNickname = False
    If InStr(1, txtNickname, ARGUMENT_SEPERATOR3) Then IsValidNickname = False
    If InStr(1, txtNickname, SOCKET_SEPERATOR) Then IsValidNickname = False
    
    If Not IsValidNickname Then
        Cancel = True
        MyMsgbox "Το Nickname είναι ακατάλληλο. Δώστε ενα Nickname που περιέχει" & vbCrLf & "μόνο γράμματα ή αριθμούς.", vbInformation + vbOKOnly
    End If
    
End Sub

'Check if the server ip is allowed
Private Sub txtServerIP_Validate(Cancel As Boolean)
    Dim IsValidIp As Boolean
    IsValidIp = True
    If Not txtServerIP Like "*.*.*.*" Then IsValidIp = False
    If Not IsNumeric(Replace(txtServerIP, ".", "")) Then IsValidIp = False
    
    If Not IsValidIp Then
        MyMsgbox "Συμπληρώστε την διεύθυνση IP όπου τρέχει ο server.", vbInformation
        Cancel = True
    End If
End Sub

'Check if the server port is allowed
Private Sub txtServerPort_Validate(Cancel As Boolean)
    If Not IsNumeric(txtServerPort) Then
        MyMsgbox "Συμπληρώστε τον αριθμό της Port όπου τρέχει ο server.", vbInformation
        Cancel = True
    End If
End Sub

'Connect to the server socket
Private Sub Connect()
    SOCKET.Close
    SOCKET.Connect txtServerIP, txtServerPort
    gConnectionState = csConnecting
End Sub
