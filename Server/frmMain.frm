VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Σκάκι - SERVER"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
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
   ScaleHeight     =   4080
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox NTService 
      Height          =   480
      Left            =   1110
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   870
      Width           =   1200
   End
   Begin VB.Timer tmrServerState 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3480
      Top             =   2820
   End
   Begin VB.Timer tmrIncoming 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2010
      Top             =   2760
   End
   Begin MSWinsockLib.Winsock SOCKET 
      Index           =   0
      Left            =   390
      Top             =   2820
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

'Called when the server's state is changed, starting or ending.
Private Function UpdateServerState(ByVal NewValue As Boolean) As Boolean
    
    Dim Result As Boolean
    
    gServerIsRunning = NewValue
    
    ReDim gUsers(0) ' Clear the users array
    ReDim gGames(0) ' Clear the games array
    
    Set gOutgoing = Nothing ' Clear the outgoing messages FIFO queue
    Set gIncoming = Nothing ' Clear the incoming messages FIFO queue
    
    tmrIncoming.Enabled = gServerIsRunning
    tmrServerState.Enabled = gServerIsRunning
    
    If gServerIsRunning Then
        
        Result = SocketListen
        UpdateServerState = Result
        
        If Result = False Then
            UpdateServerState False
            Exit Function
        End If
        
    Else
        CloseSockets
    End If
    
End Function

'Close open sockets
Private Sub CloseSockets()
    Dim i As Integer
    On Error Resume Next
    For i = 0 To SOCKET.UBound
        SOCKET(i).Close
    Next
    For i = 1 To SOCKET.UBound
        Unload SOCKET(i)
    Next
End Sub

'Start "listening" on the designated port.
Private Function SocketListen() As Boolean
    On Error GoTo ErrHandler
    SOCKET(0).LocalPort = SERVER_PORT
    SOCKET(0).Listen
    
    SocketListen = True
    Exit Function

ErrHandler:
    
    LogThis "Παρουσιάστηκε πρόβλημα: " & vbCrLf & Err.Number & " " & Err.Description & " (" & Err.Source & ")"
    
End Function

'Start the server
Private Function StartServer() As Boolean
    StartServer = UpdateServerState(True)
End Function

'Stop the server
Private Sub StopServer()
        
    UpdateServerState False
    Unload Me
    
End Sub





Private Sub Form_Load()

    'StartServer
    'Exit Sub
    

    Select Case UCase(Command)
    Case "-I", "/I"
        If NTService.Install Then
            MsgBox NTService.DisplayName & " installed successfully."
        Else
            MsgBox NTService.DisplayName & " did not install successfully."
        End If
        End
    Case "-U", "/U"
        If NTService.Uninstall Then
            MsgBox NTService.DisplayName & " uninstalled successfully."
        Else
            MsgBox NTService.DisplayName & " did not uninstall successfully."
        End If
        End
    Case ""
        '-- This code should only run when the
        ' application is started without parameters
        NTService.StartService
        'StartServer
    Case Else
        '-- This code should only run when the
        ' application is started with invalid
        ' Parameters
        MsgBox "The parameter: " & Command & _
        " was is not understood. Try –I " & _
        " (Install) Or –U(Uninstall)."
        End
    End Select

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseSockets
End Sub

'Return an empty socket for a new client connection. If there is no free socket then load a new winsock control into
' control arry SOCKET
Private Function GetEmptySocket() As Integer
    Dim i As Integer, NewSocket As Integer
    For i = 1 To SOCKET.UBound
        If SOCKET(i).State = sckClosed Then
            NewSocket = i
            GetEmptySocket = NewSocket
            Exit Function
        End If
    Next
    NewSocket = SOCKET.UBound + 1
    Load SOCKET(NewSocket)
    GetEmptySocket = NewSocket
    ReDim Preserve gUsers(NewSocket)
    
End Function

Private Sub NTService_Start(Success As Boolean)
Success = StartServer
End Sub

Private Sub NTService_Stop()
    StopServer
End Sub

'This event runs when a previously connected client closes the connection to this socket.
Private Sub SOCKET_Close(Index As Integer)
    If SOCKET(Index).State <> sckClosing Then Exit Sub
    ClientDisconnected Index
End Sub

'A client has connected. Save this client in the users array with state ="connecting"
Private Sub SOCKET_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    
    Dim EmptySocket As Integer
    EmptySocket = GetEmptySocket
    
    SOCKET(EmptySocket).Accept requestID
    gUsers(EmptySocket) = Array("", "", "", "", "", "", "", "")
    gUsers(EmptySocket)(udIP) = SOCKET(EmptySocket).RemoteHostIP
    gUsers(EmptySocket)(udUserState) = usConnecting
    gUsers(EmptySocket)(udPingState) = "0"
    gUsers(EmptySocket)(udFailedPingCount) = 0
    
End Sub

'Incoming tcp/ip message
Private Sub SOCKET_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim sData As String
    'Read the message from the socket
    SOCKET(Index).GetData sData, vbString
    
    'Place message in client's queue
    gUsers(Index)(udUserQueue) = gUsers(Index)(udUserQueue) & sData
    
    'Process queue so that if complete messages exist they are processed and removed
    gUsers(Index)(udUserQueue) = TrimIncomingCommands(Index, gUsers(Index)(udUserQueue))
    
End Sub

'Process and remove complete messages from the queue
Private Function TrimIncomingCommands(ByVal socketindex As Integer, ByVal UserQueue As String) As String
    Dim CommandArray As Variant, i As Integer, Command As String, NewUserQueue As String
    CommandArray = Split(UserQueue, COMMAND_SEPERATOR)
    For i = 0 To UBound(CommandArray)
        Command = ""
        If CommandArray(i) <> "" Then
            If Left(CommandArray(i), 1) = "<" And Right(CommandArray(i), 1) = ">" Then
                Command = Mid(CommandArray(i), 2)
                Command = Left(Command, Len(Command) - 1)
            Else
                NewUserQueue = NewUserQueue & CommandArray(i)
                If i < UBound(CommandArray) Then NewUserQueue = NewUserQueue & COMMAND_SEPERATOR
            End If
            If Command <> "" Then
                Command = socketindex & SOCKET_SEPERATOR & Command
                
                'Process message (some urgent messages are not placed in the queue)
                Dim IsUrgent As Boolean
                'If this message is an answer to a ping then don't put it in queue
                Dim dummy1 As String, dummy2 As String
                dummy1 = Split(Command, SOCKET_SEPERATOR)(1)
                dummy2 = Split(dummy1, MSG_SEPERATOR)(0) ' type of message

                Select Case dummy2
                Case mtPong, mtID, mtConnectionRequest
                    IsUrgent = True
                Case Else
                    IsUrgent = False
                End Select

                If IsUrgent Then
                    ReceiveCommand Command, True
                Else
                    ReceiveCommand Command
                End If

            End If
            
        End If
    Next
    
    TrimIncomingCommands = NewUserQueue
End Function


'THis timer processes incoming messages
Private Sub tmrIncoming_Timer()
    Dim socketindex, Command As String, TheCommand As String
    If gIncoming.Count > 0 Then
        TheCommand = gIncoming(1)
        gIncoming.Remove 1
        socketindex = Split(TheCommand, SOCKET_SEPERATOR)(0)
        Command = Split(TheCommand, SOCKET_SEPERATOR)(1)
        
        If SOCKET(socketindex).State = sckConnected Then
            ProcessCommand socketindex, Command
        End If
    End If
End Sub

Private Sub PingClients()
    
    Static NotFirstTime As Boolean
    
    Dim i As Integer
    
    If NotFirstTime = True Then
        For i = 1 To UBound(gUsers)
            If gUsers(i)(udPingState) = "1" And gUsers(i)(udUserState) = usConnected Then
'                If gUsers(i)(udFailedPingCount) > 0 Then
'                    gUsers(i)(udFailedPingCount) = 0
                    ClientDisconnected i
'                Else
'                    gUsers(i)(udFailedPingCount) = gUsers(i)(udFailedPingCount) + 1
'                End If
            End If
        Next
    End If
    
    For i = 1 To UBound(gUsers)
        gUsers(i)(udPingState) = "1"
    Next
    
    BroadcastMessage FormatCommand(mtPing, ""), True
    NotFirstTime = True
    
End Sub

'THis timer sends the server state to the clients
Private Sub tmrServerState_Timer()
    Static counter As Integer
    SendServerState
    
    counter = counter + 1
    
    'Ping all clients every 30 seconds
    If counter = 36 Then
        counter = 0
        PingClients
    End If
    
End Sub

