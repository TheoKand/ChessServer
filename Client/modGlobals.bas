Attribute VB_Name = "modGlobals"
'****************************************************************
' This module contains all global variables for the client,
' configuration methods, utility methods and everything else
' that is not part of the input/presentation logic.
'
' Copyright Theo Kandiliotis 2004
'****************************************************************

Option Explicit

'Possible values for the client's connection state during login
Public Enum enumConnectionState
    csDisconnected = 0 ' Disconnected
    csConnecting = 1 ' Connecting
    csLoging = 2 ' Logging on
    csConnected = 3 ' Log on complete. The client is online in the lobby.
End Enum

'Possible values for the client's state after the login was successful
Public Enum enumUserState
    usIdle = 0 ' Client is in the lobby without an active game
    usHosting = 1 ' Client is hosting a game and waiting for opponents
    usConnecting = 2 ' Client is in the process of connecting to onother client's hosted game
    usPlayingServer = 3 ' Client is currently in a game hosted by himself
    usPlayingClient = 4 ' Client is currently in a game hosted by onother client
End Enum

Public gIncoming As New Collection ' FIFO queue for incoming tcp/ip messages from other peers. Processed every 500 ms

Public gQueue As String ' Queue of raw incoming messages, must be trimmed before they can be processed
Public gConnectionState As enumConnectionState ' Connection state
Public gUserState As enumUserState ' User state

Public gShutdown As Boolean

'Wait for x sec
Public Sub Wait(ByVal sec As String)
    Dim StartTimer As Single
    StartTimer = Timer
    Do Until Timer > StartTimer + sec
        DoEvents
        If gShutdown Then Exit Do
    Loop
End Sub

'Here you can change the fixed messages in the lobby chat
Public Sub InitLobbyChat(cbo As ComboBox)
    
    
    Dim ChatMessages  As Variant, i As Long
    
    ChatMessages = Array( _
        "Καλημέρα σε όλους!", _
        "Καλησπέρα σε όλους!", _
        "Καληνύχτα σε όλους!", _
        "Θέλει κανένας να παίξει μαζί μου;", _
        "Είμαι πολύ δυνατός παίκτης.", _
        "Είμαι αρχάριος στο σκάκι.", _
        "Κανένας καλός για μια παρτίδα;", _
        "Ας παίξουμε.")
        
    For i = 0 To UBound(ChatMessages)
        cbo.AddItem ChatMessages(i)
    Next

End Sub

'Here you can change the fixed messages in the game chat
Public Sub InitGameChat(cbo As ComboBox)
    
    Dim ChatMessages  As Variant, i As Long
    
    
    ChatMessages = Array( _
    "Καλημέρα.", _
    "Καλησπέρα.", _
    "Καλώς σε βρήκα!", _
    "Σε έχω φέρει σε δύσκολη θέση;", _
    "Είσαι πολύ καλός παίκτης.", _
    "Δεν μπορείς να με αντιμετωπίσεις!", _
    "Συγκεντρώσου! Ήταν αφηρημένη κίνηση.", _
    "Σου ""έφαγα"" τη βασίλισσα!", _
    "Σου ""έφαγα"" τον πύργο!", _
    "Σου ""έφαγα"" τον αξιωματικό!", _
    "Σου ""έφαγα"" το αλογάκι!", _
    "Παραδώσου! Σε έχω κατατροπώσει!", _
    "Έχασα! Συγχαρητήρια, έπαιξες καλά!", _
    "Κέρδισα! Συγχαρητήρια, έπαιξες καλά!", _
    "Ας παίξουμε!", _
    "’ντε μην αργείς!", _
    "Λάθος κίνηση !", _
    "Καλή κίνηση !", _
    "Βλέπω είσαι καλός!", _
    "Λιγάκι.", _
    "Ετοιμάσου να χάσεις!", _
    "Δε νομίζω!")
            
    For i = 0 To UBound(ChatMessages)
        cbo.AddItem ChatMessages(i)
    Next

End Sub

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

'Send a message to the server
Public Sub SendCommand(ByVal CommandText As String)
    SendToSocket CommandText
End Sub

'REceive a message from the server and place in the FIFO queue
Public Sub ReceiveCommand(ByVal CommandText As String)
    ProcessCommand CommandText
    'gIncoming.Add CommandText, GetNewKey(gIncoming)
End Sub

'Process an incoming message sent to this client
Public Sub ProcessCommand(ByVal Command As String)
    Dim MsgType As Byte, OriginalCommand As String
    Dim Arg1 As String, Arg2 As String, Arg3 As String, Arg4 As String, Arg5 As String, Arg6 As String
    
    OriginalCommand = Command
    
    'The message contains the type of message and the content
    MsgType = Split(Command, MSG_SEPERATOR)(0) 'type of message
    Command = Split(Command, MSG_SEPERATOR)(1) 'content
    
    LogThis "INCOMING:" & OriginalCommand
    
    Select Case MsgType
    
    Case mtPing
        SendCommand FormatCommand(mtPong, "")
        
    ' Server reply to a client's login. User found. The message also contains other online users and current games
    Case mtLoginOK
        gConnectionState = csConnected
        UpdateServerSTate Command
        
        SendCommand FormatCommand(mtReady, "")
        
        PrintChatMessage frmMain.rtbChat, vbRed, 8, False, False, "Καλώς ήρθες χρήστη "
        PrintChatMessage frmMain.rtbChat, vbRed, 8, True, False, GetUserNickname & vbCrLf
    
    ' Server reply: usernot found.
    Case mtLoginNotOK
        MyMsgbox "Η σύνδεση με τον server Απέτυχε." & vbCrLf & "Λάθος Login ή Password", vbInformation
        gConnectionState = csDisconnected
    
    ' Server reply: nick name already exists
    Case mtNickExists
        MyMsgbox "Το Nickname που δώσατε χρησιμοποιείται ήδη.", vbInformation
        gConnectionState = csDisconnected
    
    ' Server reply: client must upgrade to latest version.
    Case mtOldVersion
        MyMsgbox Command, vbInformation, "Η έκδοση που έχετε είναι η " & App.Major & "." & App.Minor & "." & App.Revision
        gConnectionState = csDisconnected
    
    'Server is sending the state which contains other online clients and open games.
    Case mtServerState
        UpdateServerSTate Command
    
    ' Server is broadcasting a chate message to everyone
    Case mtLobbyBroadcast
        ShowChatMessage Command
        
    Case Else
        ' Server is forwarding a client message from one client to onother
        If IsGameCommand(MsgType) Then
            If FormIsLoaded("frmGame") Then frmGame.ProcessGameCommand OriginalCommand
        End If

    End Select
End Sub

Public Sub PrintChatMessage(ByVal RTB As RichTextBox, ByVal Color As ColorConstants, ByVal Size As Byte, ByVal Bold As Boolean, ByVal Italic As Boolean, ByVal Text As String)
    With RTB
        .SelStart = Len(.Text)
        .SelColor = Color
        .SelBold = Bold
        .SelItalic = Italic
        .SelFontSize = Size
        .SelText = Text
    End With
End Sub

'Display a server message in the chat window. It's either a server broadcast message like
' "the game started" or a chat message from onother client
Public Sub ShowChatMessage(ByVal Command As String)
    Dim MsgType As Byte, Message As String, Nickname As String, ChatMessage As String
    MsgType = Trim(Split(Command, ARGUMENT_SEPERATOR)(0))
    Message = Trim(Split(Command, ARGUMENT_SEPERATOR)(1))
    
    If MsgType = 0 Then ' server broadcast message
        PrintChatMessage frmMain.rtbChat, vbBlack, 8, True, False, "Server: "
        PrintChatMessage frmMain.rtbChat, vbBlack, 8, False, True, Message & vbCrLf
    ElseIf MsgType = 1 Then ' chat message from other client
        Nickname = Split(Message, ">")(0)
        ChatMessage = Split(Message, ">")(1)
        PrintChatMessage frmMain.rtbChat, vbBlue, 8, True, False, Nickname & "> "
        PrintChatMessage frmMain.rtbChat, vbBlack, 8, False, False, ChatMessage & vbCrLf
    End If
    
End Sub

'Update listboxes with online users and games
Public Sub UpdateServerSTate(ByVal ServerState As String)
    Dim UsersList As Variant, GamesList As Variant, i As Integer, UserCount As Integer, GameData As Variant, GameCount As Integer, GameDesc As String
    UsersList = Split(Split(ServerState, ARGUMENT_SEPERATOR)(0), ARGUMENT_SEPERATOR2)
    GamesList = Split(Split(ServerState, ARGUMENT_SEPERATOR)(1), ARGUMENT_SEPERATOR2)
    
    frmMain.lstUsers.Clear
    For i = 0 To UBound(UsersList)
        If Trim(UsersList(i)) <> "" Then
            frmMain.lstUsers.AddItem UsersList(i)
            UserCount = UserCount + 1
        End If
    Next
    frmMain.lblUsersOnline = "Χρήστες Online (" & UserCount & ")"
    
    frmMain.lstGames.Clear
    For i = 0 To UBound(GamesList)
        If Trim(GamesList(i)) <> "" Then
            GameCount = GameCount + 1
            GameData = Split(GamesList(i), ARGUMENT_SEPERATOR3)
            
            If UBound(GameData) = 0 Then ' Open game (hosting player is waiting for opponent)
                GameDesc = GameData(0) & " vs (κενή θέση)"
            Else ' Game in progress
                GameDesc = GameData(0) & " vs " & GameData(1) & " (Σε εξέλιξη)"
            End If
            
            frmMain.lstGames.AddItem GameDesc
        End If
    Next
    frmMain.lblActiveGames = "Ενεργές Παρτίδες (" & GameCount & ")"
    
End Sub

'Send a tcp/ip message to the server
Public Function SendToSocket(ByVal Command As String) As Boolean
    If frmMain.SOCKET.State = sckConnected Then
        frmMain.SOCKET.SendData Command
        
        LogThis "OUTGOING:" & Command
        
        SendToSocket = True
    Else
        SendToSocket = False
    End If
End Function

Public Function FormIsLoaded(ByVal FormName As String) As Boolean
    Dim f As Form
    For Each f In Forms
        If InStr(1, f.Name, FormName, vbTextCompare) Then
            FormIsLoaded = True
            Exit Function
        End If
    Next
End Function

Public Sub PositionGameForm()
    If FormIsLoaded("frmGame") Then frmGame.Move frmMain.Left + frmMain.Width, frmMain.Top
End Sub

'Convert the ip of onother client to a nickname by looking up the listbox
Public Function RemoteIPToNick(ByVal IP As String) As String
    Dim i As Integer, UserDesc As String
    For i = 0 To frmMain.lstUsers.ListCount - 1
        UserDesc = frmMain.lstUsers.List(i)
        If InStr(1, UserDesc, "(" & IP & ")", vbTextCompare) Then
            RemoteIPToNick = Trim(Split(UserDesc, "(")(0))
            Exit Function
        End If
    Next
    RemoteIPToNick = IP
End Function

Public Function MyMsgbox(ByVal Prompt As String, ByVal Buttons As VbMsgBoxStyle, Optional ByVal Title As String = "")
    If Title = "" Then Title = GetUserNickname
    If Title = "" Then Title = App.Title
    
    MyMsgbox = MsgBox(Prompt, Buttons, Title)
    
End Function

Public Sub SetFormFocus(ByVal f As Form)
    On Error Resume Next
    f.SetFocus
End Sub

Public Sub LogThis(ByVal s As String)
    'Debug.Print s
    'frmMain.txtDebug.Text = frmMain.txtDebug.Text & s & vbCrLf
    'frmMain.txtDebug.SelStart = Len(frmMain.txtDebug.Text)
End Sub

Public Function GetNicknamesArray()
    Dim fso As New FileSystemObject, fName As String
    fName = FixPath(App.Path, False) & "\settings.dat"
    
    If Not fso.FileExists(fName) Then
        MsgBox "Δεν βρέθηκε το αρχείο settings.dat" & vbCrLf & vbCrLf & "Παρακαλούμε εγκαταστήστε ξανα το πρόγραμμα.", vbExclamation + vbOKOnly
        End
    End If
    
    GetNicknamesArray = Split(DecodeText(fso.GetFile(fName).OpenAsTextStream.ReadAll), vbCrLf)
    
End Function

Private Function DecodeText(ByVal aText As String) As String
    DecodeText = ""
    Dim i As Long, j As Integer
    Dim dummy, dummy1, dummy2
    
    For i = 1 To Len(aText) Step 3
        dummy = Mid(aText, i, 3)
        DecodeText = DecodeText & Chr(CInt(dummy))
    Next
End Function

Public Function FixPath(ByVal aPath As String, ByVal WithBackslash As Boolean)
    If WithBackslash Then
        If Right(aPath, 1) <> "\" Then
            FixPath = aPath & "\"
        Else
            FixPath = aPath
        End If
    Else
        If Right(aPath, 1) = "\" Then
            FixPath = Left(aPath, Len(aPath) - 1)
        Else
            FixPath = aPath
        End If
        
    End If
End Function

Public Function GetUserNickname() As String
    GetUserNickname = frmMain.cboNickname.Text
End Function
