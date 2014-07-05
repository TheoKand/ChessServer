Attribute VB_Name = "modNetwork"
'****************************************************************
' This class describes the protocol that was built for the chess
' game on top of the winsock control. There are a number of
' available commands that are sent between the peers. This class
' is common between the client and server projects
'
' Copyright Theo Kandiliotis 2004
'****************************************************************

Option Explicit

'The messages that are sent between peers are made up from several parts. The parts are
'seperated by these delimiters (constants)
Public Const COMMAND_SEPERATOR = "#"
Public Const MSG_SEPERATOR = "^"
Public Const ARGUMENT_SEPERATOR = "&"
Public Const ARGUMENT_SEPERATOR2 = "$"
Public Const ARGUMENT_SEPERATOR3 = "@"
Public Const SOCKET_SEPERATOR = "%"
Public Const CLIENT_SEPERATOR = "{"
    
'The various types of messages exchanged between the peers
Public Enum enumMsgType
    
    'Client --> Server
    mtID = 0 'Send by the client to the server after connection, contains username,password,nickname and client EXE version
    mtReady = 1 'Sent by the client to the server after the connection process was successful
    mtLobbySay = 2 ' Chat message from one client to the server lobby
    mtClientStartGame = 3 ' A hosting client starts a game
    mtClientEndGame = 4 ' A hosting client ends a game
    mtClientGameStarted = 5 ' A hosting client sends this message to the server when onother client has connected to the game and the game begins
    mtClientResigned = 6 ' A game is over after a client has resigned
    mtDraw = 7 ' The game is over with a draw
    mtPlayerWon = 8 ' The game is over, the client that sent this message has won
    mtPong = 9 ' Reply to ping
    
    'Server --> Client
    mtServerState = 10 'Server state information, contains other clients and open games
    mtLoginOK = 11 ' Server reply to client login. Authentication OK, contains other clients and games info
    mtLoginNotOK = 12 ' Server reply to client login. Authentication failed, user not found.
    mtOldVersion = 13 ' Server reply to client login. The client's version is outdated
    mtNickExists = 14 ' Server reply to client login. Nickname already exists.
    mtPing = 15
    
    'Server --> All Clients (broadcast)
    mtLobbyBroadcast = 16 ' Server broadcasts lobby message to all clients.
    
    'Client --> Client (actually routed through server)
    mtGameSay = 17 ' Chat message sent during a game
    mtGameMove = 18 ' Game move
    mtGameOfferDraw = 19 ' A player offers the other a draw
    mtGameAcceptDraw = 20 ' A player accepts a draw
    mtGameDenyDraw = 21 ' A player denies the draw
    mtGameResign = 22 ' A player resigns.
    mtNick = 23 ' A connecting player sends this message to the host-player after the connection begins
    mtConnectionAccepted = 24 ' Hosting player accepts a connection by connecting player
    mtConnectionDenied = 25 ' Hosting player denies to play with the connecting player
    mtConnectionRequest = 26 ' Connecting player asks a player that is hosting a game if he can play
    mtGameStop = 27 ' Message sent by either player when the game must stop
End Enum

'The messages sent between peers have this format
'
'< MsgType ^ Command > #
'
' Made up from two parts, type of message and the command it self
'
Public Function FormatCommand(ByVal MsgType As enumMsgType, ByVal Command As String) As String
    FormatCommand = "<" & MsgType & MSG_SEPERATOR & Command & ">" & COMMAND_SEPERATOR
End Function

'Returns TRUE if this message is one of those exchanged between clients during a game
Public Function IsGameCommand(ByVal MsgType As enumMsgType)
    
    If MsgType = mtGameSay Or MsgType = mtGameMove Or MsgType = mtGameOfferDraw Or MsgType = mtGameAcceptDraw Or _
    MsgType = mtGameDenyDraw Or MsgType = mtGameResign Or MsgType = mtNick Or MsgType = mtConnectionAccepted Or _
    MsgType = mtConnectionDenied Or MsgType = mtConnectionRequest Or MsgType = mtGameStop Then
        IsGameCommand = True
    Else
        IsGameCommand = False
    End If
    
End Function










