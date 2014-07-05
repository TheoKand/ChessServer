Attribute VB_Name = "modGlobals"
'****************************************************************
' This module contains all global variables for the client,
' configuration methods, utility methods and everything else
' that is not part of the input/presentation logic.
'
' Copyright Theo Kandiliotis 2004
'****************************************************************

Option Explicit

' Here you can change the allowed client version. Clients with different version that try to connect
' will reveive an error message. In the client application you can change the version on the menu
' Visual Basic Project > Project Properties > Make and compile
Public Const CLIENT_VERSION = "1.0.0"
Public Const SERVER_PORT = 30200

'Fields of information that the server stores for each client
Public Enum enumUserData
    udUsername = 0 'username
    udPassword = 1 'password
    udIP = 2 ' IP
    udUserState = 3 'client state
    udNickname = 4 ' Nickname
    udUserQueue = 5 ' queue of client messages
    udPingState = 6
    udFailedPingCount = 7
End Enum

Public Enum enugUserstate  ' possible client connection states
    usConnecting = 0 ' The client has connected to the socket but has not logged on yet
    usConnected = 1 ' The client is logged on and other clients can see him
    usDisconnected = 2 ' The client is disconnected
End Enum

'The fields of information that the server stores for each game
Public Enum enumGameData
    gdServerPlayer = 0 ' SocketIndex of hosting player (always white)
    gdClientPlayer = 1 ' SocketIndex of connecting player (always black)
    gdGameStatE = 2 ' game state
End Enum

Public Enum enugGameState  ' Possible values for the game state
    gsOpen = 1 ' The hosting client is waiting for an opponent (connecting player)
    gsPlaying = 2 ' The game is in progress
    gsClosed = 3 ' The game is closed, either finished or canceled by the hosting client.
End Enum

Public gOutgoing As New Collection
Public gIncoming As New Collection ' Incoming socket messages are placed in this FIFO queue and processed every 500 ms

Public gUsers() As Variant ' Clients information is saved here
Public gGames()  As Variant ' Games information is saved here

Public gServerIsRunning As Boolean ' TRUE if the server is running
Public gServerStartTime As Date ' Time the server started


'This function is called when a client connects.Returns 0 if the login was unsuccessfull, 1 if the username exists
'2 if everything is ok and 3 if the client has an outdated version.
Public Function ValidateUser(ByVal Login As String, ByVal Password As String, ByVal Nick As String, ByVal ClientVersion As String) As Byte
    Dim i As Integer, LoginOK As Boolean
    
    '----------------------------------------------------
    'TODO : connect to planetinteractive database
    '----------------------------------------------------
    LoginOK = (Login = "theo" And Password = "theo")
    
    
    'The constant CLIENT_VERSION has the version that clients must have in order to play
    If CLIENT_VERSION <> ClientVersion Then
        ValidateUser = 3
        Exit Function
    End If
    
    If Not LoginOK Then
        ValidateUser = 0
        Exit Function
    End If
    
    'Check if nickname already exists amonst online players
    For i = 1 To UBound(gUsers)
        If (gUsers(i)(udUserState) = usConnected) And (LCase(gUsers(i)(udNickname)) = LCase(Nick)) Then
            ValidateUser = 1
            Exit Function
        End If
    Next
    ValidateUser = 2
End Function

'Returns a new unique key for the collection object
Public Function GetNewKey(col As Collection) As String
    Dim i As Integer, dummy As String
    On Error Resume Next
    i = 1
    Do Until GetNewKey <> ""
        dummy = col.Item("Key" & i)
        If Err <> 0 Then
            GetNewKey = "Key" & i
            Exit Function
        End If
        i = i + 1
    Loop
End Function

'Sends a socket message to the client. Either directly or through a FIFO queue
Public Sub SendCommand(ByVal CommandText As String, Optional ByVal NoQueue As Boolean = False)
    NoQueue = True
    If NoQueue = False Then
        gOutgoing.Add CommandText, GetNewKey(gOutgoing)
    Else
        Dim socketindex As Integer, Command As String
        
        'The message also contains the socketIndex of the player that must receive it
        socketindex = Split(CommandText, SOCKET_SEPERATOR)(0)
        Command = Split(CommandText, SOCKET_SEPERATOR)(1)
        SendToSocket socketindex, Command
    End If
End Sub

'Receives a socket message from a client. It's either processed immediately or placed in the FIFO queue
Public Sub ReceiveCommand(ByVal CommandText As String, Optional ByVal NoQueue As Boolean = False)
    If NoQueue = False Then
        gIncoming.Add CommandText, GetNewKey(gIncoming)
    Else
        Dim socketindex As Integer, Command As String
        
        'The message also contains the socketIndex of the player that must receive it
        socketindex = Split(CommandText, SOCKET_SEPERATOR)(0)
        Command = Split(CommandText, SOCKET_SEPERATOR)(1)
        If frmMain.SOCKET(socketindex).State = sckConnected Then
            ProcessCommand socketindex, Command
        End If
    End If
End Sub

'This function is called to process all incoming messages sent to the server
Public Sub ProcessCommand(ByVal socketindex As Integer, ByVal Command As String)
    Dim MsgType As Byte, Result As Byte
    Dim Arg1 As String, Arg2 As String, Arg3 As String, Arg4 As String, Arg5 As String, arg6 As String
    
    'The message also contains the type
    MsgType = Split(Command, MSG_SEPERATOR)(0) 'type of message
    Command = Split(Command, MSG_SEPERATOR)(1) 'content
    
    Debug.Print "INCOMING FROM " & gUsers(socketindex)(udNickname) & ":" & MsgType & MSG_SEPERATOR & Command
    
    'Process the message depenging on the type
    Select Case MsgType
    
    Case mtPong
        If gUsers(socketindex)(udUserState) = usConnected Then gUsers(socketindex)(udPingState) = "2"
    
    'Sent by the client after a connection. Contains the username,password, nickname and client EXE version
    Case mtID
        Arg1 = Split(Command, ARGUMENT_SEPERATOR)(0) ' login
        Arg2 = Split(Command, ARGUMENT_SEPERATOR)(1) ' password
        Arg3 = Split(Command, ARGUMENT_SEPERATOR)(2) ' nickname
        Arg4 = Split(Command, ARGUMENT_SEPERATOR)(3) ' client version
        Result = ValidateUser(Arg1, Arg2, Arg3, Arg4)
        If Result = 2 Then
            
            gUsers(socketindex)(udUsername) = Arg1
            gUsers(socketindex)(udPassword) = Arg2
            gUsers(socketindex)(udNickname) = Arg3
            
            SendCommand socketindex & SOCKET_SEPERATOR & FormatCommand(mtLoginOK, GetServerState), True
        
        ElseIf Result = 1 Then
            SendCommand socketindex & SOCKET_SEPERATOR & FormatCommand(mtNickExists, ""), True
        ElseIf Result = 0 Then
            SendCommand socketindex & SOCKET_SEPERATOR & FormatCommand(mtLoginNotOK, ""), True
        ElseIf Result = 3 Then
            SendCommand socketindex & SOCKET_SEPERATOR & FormatCommand(mtOldVersion, "Υπάρχει νέα έκδοση του client." & vbCrLf & vbCrLf & "Παρακαλούμε εγκαταστήστε την έκδοση " & CLIENT_VERSION & " και ξαναπροσπαθήστε."), True
        End If
    
    'Sent by the client after the login is finished
    Case mtReady
        gUsers(socketindex)(udUserState) = usConnected
        
        'BroadcastMessage FormatCommand(mtLobbyBroadcast, "0" & ARGUMENT_SEPERATOR & "Ο χρήστης " & gUsers(socketindex)(udNickname) & " συνδέθηκε στον server.")
        

    
    ' Chat message from a client to the server lobby
    Case mtLobbySay
        BroadcastMessage FormatCommand(mtLobbyBroadcast, "1" & ARGUMENT_SEPERATOR & gUsers(socketindex)(udNickname) & ">" & Command)
    
    ' Start of a new game by a hosting client
    Case mtClientStartGame
        StartNewGame socketindex
    
    ' End a game by a hosting client
    Case mtClientEndGame
        StopNewGame socketindex
    
    ' A hosting client informs the server when a connecting client has connected and the game has started
    Case mtClientGameStarted
        StartNewGame socketindex, Command
    
    ' The game ended after a client resigned and lost
    Case mtClientResigned
        BroadcastMessage FormatCommand(mtLobbyBroadcast, "0" & ARGUMENT_SEPERATOR & "Η παρτίδα " & Command & " έληξε υπέρ του " & gUsers(socketindex)(udNickname))
    
    ' The game ended with a draw
    Case mtDraw
        BroadcastMessage FormatCommand(mtLobbyBroadcast, "0" & ARGUMENT_SEPERATOR & "Η παρτίδα " & Command & " έληξε με ΙΣΟΠΑΛΙΑ")
    
    ' The game ended with a win by the player that sent this message
    Case mtPlayerWon
        BroadcastMessage FormatCommand(mtLobbyBroadcast, "0" & ARGUMENT_SEPERATOR & "Η παρτίδα " & Command & " έληξε υπέρ του " & gUsers(socketindex)(udNickname))
    Case Else
        'Forward a client message to onother client
        If IsGameCommand(MsgType) Then
            Arg1 = Split(Command, CLIENT_SEPERATOR)(0) ' nickname of the client that must receive this message
            Arg2 = Split(Command, CLIENT_SEPERATOR)(1) ' message content
            SendCommand GetUserIndex(Arg1) & SOCKET_SEPERATOR & FormatCommand(MsgType, Arg2)
        End If
    End Select
    
End Sub

'if OpponentNick is empty, a hosting client has just started a game and waiting for an opponent.
'if OpponentNick is not empty, a connecting client has just connected to the game
Public Sub StartNewGame(ByVal socketindex As Integer, Optional ByVal OpponentNick As String = "")
    Dim i As Integer, TheGameIndex As Integer
    
    'New open game waiting for opponent
    If OpponentNick = "" Then
        
        'Find an empty place in array gGames and put the new game there
        For i = 1 To UBound(gGames)
            If gGames(i)(gdGameStatE) = gsClosed Then
                TheGameIndex = i
                Exit For
            End If
        Next
        If TheGameIndex = 0 Then
            TheGameIndex = UBound(gGames) + 1
            ReDim Preserve gGames(TheGameIndex)
            gGames(TheGameIndex) = Array("", "", "")
        End If
        
        'Save the new game in memory
        gGames(TheGameIndex)(gdGameStatE) = gsOpen
        gGames(TheGameIndex)(gdServerPlayer) = socketindex
        gGames(TheGameIndex)(gdClientPlayer) = ""
    
        BroadcastMessage FormatCommand(mtLobbyBroadcast, "0" & ARGUMENT_SEPERATOR & "Ο χρήστης " & gUsers(socketindex)(udNickname) & " περιμένει αντίπαλο."), True
    
    'Begin the game
    Else
    
        For i = 1 To UBound(gGames)
            If gGames(i)(gdGameStatE) = gsOpen And gGames(i)(gdServerPlayer) = socketindex Then
                TheGameIndex = i
                Exit For
            End If
        Next
        
        'Change the game's state from Open to Playing
        gGames(TheGameIndex)(gdGameStatE) = gsPlaying
        gGames(TheGameIndex)(gdServerPlayer) = socketindex
        gGames(TheGameIndex)(gdClientPlayer) = GetUserIndex(OpponentNick)
        
        BroadcastMessage FormatCommand(mtLobbyBroadcast, "0" & ARGUMENT_SEPERATOR & "Ξεκίνησε η παρτίδα " & gUsers(socketindex)(udNickname) & " vs " & gUsers(gGames(TheGameIndex)(gdClientPlayer))(udNickname)), True
    
    End If
    
End Sub

'Return the socketIndex of the user based on this criteria
Public Function GetUserIndex(Optional ByVal Nickname As String = "", Optional ByVal IP As String = "") As Integer

    Dim Crit As Integer, i As Integer, CritValue As String
    
    If Nickname <> "" Then
        Crit = udNickname
        CritValue = Nickname
    ElseIf IP <> "" Then
        Crit = udIP
        CritValue = IP
    End If
    
    GetUserIndex = 0
    
    For i = 1 To UBound(gUsers)
        If gUsers(i)(udUserState) = usConnected And gUsers(i)(Crit) = CritValue Then
            GetUserIndex = i
            Exit Function
        End If
    Next
    

End Function

'A hosting client canceled the game
Public Sub StopNewGame(ByVal socketindex As Integer)
    Dim Gameindex As Integer, i As Integer
    For i = 1 To UBound(gGames)
        If gGames(i)(gdServerPlayer) = socketindex Then
            gGames(i)(gdGameStatE) = gsClosed
            Exit For
        End If
    Next

End Sub

'Returns the number of online users
Public Function GetOnlineUserCount() As Integer
    Dim i As Integer
    For i = 1 To UBound(gUsers)
        If gUsers(i)(udUserState) = usConnected Then GetOnlineUserCount = GetOnlineUserCount + 1
    Next
End Function

'REturns the number of currently open games
Public Function GetOnlineGamesCount() As Integer
    Dim i As Integer
    For i = 1 To UBound(gGames)
        If gGames(i)(gdGameStatE) <> gsClosed Then GetOnlineGamesCount = GetOnlineGamesCount + 1
    Next
End Function

'Returns the server state that is broadcasted to the clients every few seconds. This contains the
'list of online users and online games
Public Function GetServerState() As String
    Dim i As Integer, UsersList As String, GamesList As String
    
    'GEt online users
    For i = 1 To UBound(gUsers)
        If gUsers(i)(udUserState) = usConnected Then
            UsersList = UsersList & gUsers(i)(udNickname) & ARGUMENT_SEPERATOR2
        End If
    Next
    
    'Get online games
    For i = 1 To UBound(gGames)
        
        If gGames(i)(gdGameStatE) = gsOpen Then
            GamesList = GamesList & gUsers(gGames(i)(gdServerPlayer))(udNickname) & ARGUMENT_SEPERATOR2
            
        ElseIf gGames(i)(gdGameStatE) = gsPlaying Then
        
            GamesList = GamesList & gUsers(gGames(i)(gdServerPlayer))(udNickname) & ARGUMENT_SEPERATOR3 & gUsers(gGames(i)(gdClientPlayer))(udNickname) & _
            ARGUMENT_SEPERATOR2
        
        End If
    Next
    GetServerState = UsersList & ARGUMENT_SEPERATOR & GamesList
End Function

'Send a message to all clients
Public Sub BroadcastMessage(ByVal Command As String, Optional NoQueue As Boolean = False)
    Dim i As Long
    For i = 1 To UBound(gUsers)
        If gUsers(i)(udUserState) = usConnected Then
            SendCommand i & SOCKET_SEPERATOR & Command, NoQueue
        End If
    Next
End Sub

'Called when a client disconnects
Public Function ClientDisconnected(ByVal socketindex As Integer)
    
    frmMain.SOCKET(socketindex).Close
    
    If gUsers(socketindex)(udUserState) = usDisconnected Then Exit Function
    
    'If gUsers(socketindex)(udUserState) = usConnected Then
        'BroadcastMessage FormatCommand(mtLobbyBroadcast, "0" & ARGUMENT_SEPERATOR & "Ο χρήστης " & gUsers(socketindex)(udNickname) & " αποσυνδέθηκε.")
    'End If
    
    gUsers(socketindex)(udUserState) = usDisconnected
    
    Dim i As Long
    
    'If the client was in the middle of a game, inform the other player that the game is over
    Dim OpponentIndex As Integer
    For i = 1 To UBound(gGames)
        If gGames(i)(gdServerPlayer) = socketindex And gGames(i)(gdGameStatE) = gsPlaying Then
            OpponentIndex = gGames(i)(gdClientPlayer)
        ElseIf gGames(i)(gdClientPlayer) = socketindex And gGames(i)(gdGameStatE) = gsPlaying Then
            OpponentIndex = gGames(i)(gdServerPlayer)
        End If
        
        If OpponentIndex <> 0 Then
            SendCommand OpponentIndex & SOCKET_SEPERATOR & FormatCommand(mtGameStop, "")
            gGames(i)(gdGameStatE) = gsClosed
        End If
    Next
    
    
    'Delete any open games that the client had started
    For i = 1 To UBound(gGames)
        If gGames(i)(gdServerPlayer) = socketindex Then gGames(i)(gdGameStatE) = gsClosed
    Next

    
End Function

'Send a tcp/ip message to a specific client. Each client has a unique socketIndex which is the index of the control array
'of the winsock control that this client is connected to.
Public Function SendToSocket(ByVal socketindex As Integer, ByVal Command As String) As Boolean
    If frmMain.SOCKET(socketindex).State = sckConnected Then
        frmMain.SOCKET(socketindex).SendData Command
        
        Debug.Print "OUTGOING TO " & gUsers(socketindex)(udNickname) & ":" & Command
        
        SendToSocket = True
    Else
        SendToSocket = False
    End If
End Function

'Send the server state to all clients only if it has changed since the last sending.
Public Sub SendServerState()
    
    Static ServerState As String
    Dim NewServerState As String, i As Integer
    If Not gServerIsRunning Then Exit Sub
    
    NewServerState = GetServerState
    
    If NewServerState <> ServerState Then
        BroadcastMessage FormatCommand(mtServerState, NewServerState), True
        ServerState = NewServerState
    End If
    
End Sub

Public Sub LogThis(ByVal Msg As String)
    frmMain.NTService.LogEvent svcEventInformation, svcMessageInfo, Msg
End Sub


Public Sub Main()
    Load frmMain
End Sub
