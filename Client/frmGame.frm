VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The_TROOPER vs (κενή θέση)"
   ClientHeight    =   10590
   ClientLeft      =   7275
   ClientTop       =   330
   ClientWidth     =   12705
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   706
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   847
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Board_black 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00000080&
      Height          =   5850
      Left            =   8820
      Picture         =   "frmGame.frx":030A
      ScaleHeight     =   390
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   390
      TabIndex        =   17
      Top             =   2730
      Visible         =   0   'False
      Width           =   5850
   End
   Begin VB.Timer tmrEnableSend 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   660
      Top             =   8880
   End
   Begin VB.PictureBox Ruler1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5265
      Left            =   7920
      MousePointer    =   11  'Hourglass
      Picture         =   "frmGame.frx":1B2D
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   15
      Top             =   6960
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox Ruler2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5265
      Left            =   8400
      MousePointer    =   11  'Hourglass
      Picture         =   "frmGame.frx":2122
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   14
      Top             =   6750
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox Board_white 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00000080&
      Height          =   5850
      Left            =   9150
      Picture         =   "frmGame.frx":270C
      ScaleHeight     =   390
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   390
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   5850
   End
   Begin VB.PictureBox Sprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2460
      Picture         =   "frmGame.frx":3F37
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.Frame Frame2 
      Caption         =   "Επιλογές"
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
      Left            =   6060
      TabIndex        =   9
      Top             =   6330
      Width           =   1785
      Begin VB.CommandButton cmdStats 
         Caption         =   "&Απώλειες"
         Enabled         =   0   'False
         Height          =   345
         Left            =   150
         TabIndex        =   16
         Top             =   1020
         Width           =   1485
      End
      Begin VB.CommandButton cmdResign 
         Caption         =   "&Παραίτηση"
         Enabled         =   0   'False
         Height          =   345
         Left            =   150
         TabIndex        =   2
         Top             =   660
         Width           =   1485
      End
      Begin VB.CommandButton cmdOfferDraw 
         Caption         =   "&Ισοπαλία"
         Enabled         =   0   'False
         Height          =   345
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   1485
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Έξοδος"
         Height          =   345
         Left            =   150
         TabIndex        =   0
         Top             =   1650
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ιστορικό"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   6060
      TabIndex        =   7
      Top             =   360
      Width           =   1785
      Begin VB.PictureBox rtbHistory 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5445
         Left            =   90
         ScaleHeight     =   5445
         ScaleWidth      =   1605
         TabIndex        =   8
         Top             =   300
         Width           =   1605
      End
   End
   Begin VB.ComboBox cboChat 
      Enabled         =   0   'False
      Height          =   315
      Left            =   630
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   8100
      Width           =   5325
   End
   Begin VB.PictureBox Screen 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5910
      Left            =   60
      MousePointer    =   11  'Hourglass
      ScaleHeight     =   390
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   390
      TabIndex        =   5
      Top             =   360
      Width           =   5910
   End
   Begin VB.PictureBox rtbChat 
      BackColor       =   &H00E0E0E0&
      Height          =   1635
      Left            =   60
      ScaleHeight     =   1575
      ScaleWidth      =   5835
      TabIndex        =   3
      Top             =   6390
      Width           =   5895
   End
   Begin VB.PictureBox Buffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5850
      Left            =   8100
      ScaleHeight     =   390
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   390
      TabIndex        =   11
      Top             =   720
      Width           =   5850
   End
   Begin VB.Label lblGameStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Περιμένετε αντίπαλο ..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   60
      TabIndex        =   10
      Top             =   30
      Width           =   7800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Chat: "
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   8160
      Width           =   540
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFormUnloaded As Boolean

'Το object cChessGame περιέχει όλη την λειτουργικότητα του παιγνιδιού
Private WithEvents mGame As cChessGame
Attribute mGame.VB_VarHelpID = -1

Private mIsServer As Boolean ' Είναι ο client αυτός ο Hosting client ή ο connecting client?
Private mOpponentName As String ' Το nickname του αντιπάλου
Private mRemoteIP As String ' To IP του αντιπάλου
Private mQueue As String ' Το queue των incoming tcp/ip μυνημάτων που στέλνονται απο τον
'αντίπαλο κατα την διάρκεια του game


'Αποστολή chat μυνήματος στον αντίπαλο
Private Sub cboChat_Click()
    If cboChat.ListIndex = -1 Then Exit Sub
    Dim ChatMessage As String
    ChatMessage = GetUserNickname & " > " & cboChat.List(cboChat.ListIndex)
    
    SendToGameSocket FormatCommand(mtGameSay, ChatMessage)
    ShowGameChatMessage ChatMessage
    
    cboChat.ListIndex = -1
        
    cboChat.Enabled = False
    tmrEnableSend.Enabled = True

End Sub

'Ο client κλείνει μια παρτίδα που είναι σε εξέλιξη
Private Sub cmdExit_Click()
    
    Dim Msg As String
    If gUserState = usPlayingClient Or gUserState = usPlayingServer Then
        Msg = "Η παρτίδα βρίσκεται σε εξέλιξη. Σίγουρα θέλετε να την ακυρώσετε;"
    ElseIf gUserState <> usIdle Then
        Msg = "Είστε σίγουρος οτι θέλετε να ακυρώσετε την παρτίδα;"
    End If
    
    If Msg <> "" Then
        If MyMsgbox(Msg, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        UpdateUserState usIdle
    End If
    
    
    SaveXML
    
    Unload Me
    
End Sub


Private Sub SaveXML()
'    Dim i As Integer
'    Dim aMove As udtMove
'    For Each aMove In mGame.MovesCollection
'        Debug.Print aMove.FromX & aMove.FromY & aMove.ToX & aMove.ToY & " (" & aMove.Castling & ") - (" & aMove.EnPassant & ") - (" & aMove.PawnPromotesTo & ")"
'    Next
End Sub


'Προσφορά ισοπαλίας στον αντίπαλο. Ο αντίπαλος πρέπει να δώσει την συγκατάθεση του,
'διαφορετικά η παρτίδα συνεχίζεται κανονικά
Private Sub cmdOfferDraw_Click()
    If MyMsgbox("Σίγουρα θέλετε να προτείνετε στον αντίπαλο να λήξει η παρτίδα με ισοπαλία;", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    SendToGameSocket FormatCommand(mtGameOfferDraw, "")
End Sub

'Παραίτηση. Ο αντίπαλος νικάει
Private Sub cmdResign_Click()
    
    If MyMsgbox("Σίγουρα θέλετε να παραιτηθείτε και να χάσετε την παρτίδα;", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    UpdateUserState usIdle
    
    'Αποστολή του μυνήματος παραίτησης στον αντίπαλο
    SendToGameSocket FormatCommand(mtGameResign, "")
    
    Wait 0.5
    
    'Ενημέρωση του server για την παραίτηση ενός παίκτη
    SendCommand FormatCommand(mtClientResigned, Caption)
    
    GameResult False, "Παραιτηθήκατε απο την παρτίδα."
    
    Unload Me
    
End Sub

'Ενημέρωση του server για το αποτέλεσμα της παρτίδας
Private Sub AnnounceGameResult(Optional ByVal Draw As Boolean = False)
    
    If Draw Then
        If Not mIsServer Then Exit Sub
        SendToSocket FormatCommand(mtDraw, Caption)
    Else
        SendToSocket FormatCommand(mtPlayerWon, Caption)
    End If

End Sub

'Προβολή του αποτελέσματος της παρτίδας
Private Sub GameResult(ByVal Win As Boolean, Optional ByVal Reason As String = "", Optional ByVal Draw As Boolean = False)
    
    Dim Msg As String
    If Reason <> "" Then Msg = Reason & vbCrLf & vbCrLf
    If Win And Not Draw Then
        Msg = "ΚΕΔΙΣΑΤΕ!!!" & vbCrLf & vbCrLf & vbCrLf & _
        "Η παρτίδα με αντίπαλο τον παίκτη " & mOpponentName & " έληξε με εσάς νικητή."
    ElseIf Win = False And Not Draw Then
        Msg = "ΧΑΣΑΤΕ." & vbCrLf & vbCrLf & vbCrLf & _
        "O παίκτης " & mOpponentName & " κέρδισε την παρτίδα."
    ElseIf Draw Then
        Msg = "Η παρτίδα έληξε με ΙΣΟΠΑΛΙΑ."
    End If
    
    MyMsgbox Msg, vbInformation + vbOKOnly, "Η Παρτίδα Έληξε"

End Sub

'Προβολή στατιστικών στοιχείων
Private Sub cmdStats_Click()
    Dim Stats As String
    Stats = mGame.GetStats
    If Stats <> "" Then MyMsgbox Stats, vbInformation + vbOKOnly, "Απώλειες"
End Sub

'Η παρτίδα έληξε
Private Sub mGame_GameEnded(ByVal Msg As String, ByVal GameResult As enumGameResult)
    
    MyMsgbox Msg, vbExclamation, "ΤΕΛΟΣ ΠΑΡΤΙΔΑΣ"
    
    UpdateUserState usIdle
    
    If GameResult = grYOU_WIN Then
        lblGameStatus = "Κ Ε Ρ Δ Ι Σ Α Τ Ε"
        AnnounceGameResult
    ElseIf GameResult = grYOU_LOOSE Then
        lblGameStatus = "Χ Α Σ Α Τ Ε"
    ElseIf GameResult = grSTALEMATE Then
        lblGameStatus = " Ι Σ Ο Π Α Λ Ι Α"
        AnnounceGameResult True
    End If
    
    lblGameStatus.Font.Size = lblGameStatus.Font.Size + 1
    lblGameStatus.BackColor = vbRed
    lblGameStatus.ForeColor = vbWhite
    cboChat.Enabled = True
    
End Sub

'Αποστολή κίνησης στον αντίπαλο
Private Sub mGame_MustSendMove(ByVal FromX As Byte, ByVal FromY As Byte, ByVal ToX As Byte, ByVal ToY As Byte, ByVal PromoteToWhat As enumChessPiece, ByVal Castling As Boolean, ByVal EnPassant As Byte)
    SendToGameSocket FormatCommand(mtGameMove, FromX & ARGUMENT_SEPERATOR & FromY & ARGUMENT_SEPERATOR & ToX & ARGUMENT_SEPERATOR & ToY & ARGUMENT_SEPERATOR & PromoteToWhat & ARGUMENT_SEPERATOR & IIf(Castling = True, "1", "0") & ARGUMENT_SEPERATOR & EnPassant)
End Sub

'Ένα στρατιωτάκι πρέπει να πάρει προαγωγή. Προβολή της αντίστοιχης φόρμας για να επιλέξει ο παίκτης
Private Sub mGame_PawnPromotion(PromoteToWhat As enumChessPiece)

    frmPromote.Show vbModal
    PromoteToWhat = frmPromote.PromoteTo
    Unload frmPromote: Set frmPromote = Nothing
    
End Sub

'Ο ένας παίκτης έπαιξε, σειρά του άλλου να παίξει
Private Sub mGame_TurnChanged(ByVal PlayerToPlay As enumChessOpponent, ByVal LocalPlayer As enumChessOpponent, ByVal MoveMade As String)

    If PlayerToPlay = LocalPlayer Then
        lblGameStatus = "Είναι η σειρά σας να παίξετε..."
        Screen.MousePointer = vbDefault
        
        cmdOfferDraw.Enabled = True
        cmdResign.Enabled = True
        
    Else
        lblGameStatus = "Ο αντίπαλος σκέφτεται..."
        Screen.MousePointer = vbHourglass
        
        cmdOfferDraw.Enabled = False
        cmdResign.Enabled = False
        
    End If
    
    If MoveMade <> "" Then
        Dim Moveindex As String, FromHere As String, ToHere As String
        Moveindex = Trim(Split(MoveMade, ".")(0))
        FromHere = Split(Split(MoveMade, ".")(1), "-")(0)
        ToHere = Split(Split(MoveMade, ".")(1), "-")(1)
        
        
        If PlayerToPlay = coBLACK Then
            PrintChatMessage rtbHistory, vbBlack, 7, True, False, Format(Moveindex + 1, "00")
            PrintChatMessage rtbHistory, vbBlack, 7, False, False, " "
            PrintChatMessage rtbHistory, &H80&, 10, True, False, FromHere
            PrintChatMessage rtbHistory, &H80&, 10, True, False, ToHere
            PrintChatMessage rtbHistory, vbBlack, 7, False, False, "-"
        Else
            PrintChatMessage rtbHistory, &H80&, 10, True, False, FromHere
            PrintChatMessage rtbHistory, &H80&, 10, True, False, ToHere & vbCrLf
        End If
        
        
    End If
    
End Sub

'Ο local παίκτης ξεκινά, ή ολοκληρώνει μια κίνηση πάνω στην σκακιέρα. Μια κίνηση γίνεται
'με 2 click, και ο παρακάτω κώδικας τρέχει 2 φορές.
Private Sub Screen_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    Static BeginMove As Boolean
    Static BeginMoveX As Integer, BeginMoveY As Integer
    If Screen.MousePointer = vbHourglass Then Exit Sub
    
    Dim BoardX As Integer, BoardY As Integer
    BoardX = mGame.MouseX_To_BoardX(X)
    BoardY = mGame.MouseY_To_BoardY(y)
    
    If BoardX = -1 Or BoardY = -1 Then
        If BeginMove = True Then
            BeginMove = False
            mGame.DrawBufferToScreen
        End If
        Exit Sub
    End If
    
    If BeginMove = False Then
        'Ο local παίκτης δεν βρίσκεται στην μέση μιας κίνησης. ’ρα ξεκινάει μια κίνηση
        BeginMove = mGame.BeginMove(BoardX, BoardY)
        If BeginMove Then
            BeginMoveX = BoardX
            BeginMoveY = BoardY
        Else
            mGame.DrawBufferToScreen
        End If
    Else
        BeginMove = False
        'Ο local παίκτης βρίσκεται στην μέση μιας κίνησης. ’ρα ολοκληρώνει μια κίνηση
        If mGame.CompleteMove(BeginMoveX, BeginMoveY, BoardX, BoardY) = False Then
            mGame.DrawBufferToScreen
            BeginMove = mGame.BeginMove(BoardX, BoardY)
            If BeginMove Then
                BeginMoveX = BoardX
                BeginMoveY = BoardY
            End If
        End If
    End If
End Sub

Private Sub tmrEnableSend_Timer()
    tmrEnableSend.Enabled = False
    cboChat.Enabled = True
End Sub

'Initialization της φόρμας
Public Function Init(ByVal Server As Boolean, Optional ByVal Nickname As String) As Boolean
    
    Width = 8010
    Height = 8955
    InitGameChat cboChat
    PositionGameForm
    
    Dim InitOK As Boolean, TheLocalPlayer As enumChessOpponent
    
    mIsServer = Server
    mOpponentName = Nickname
    
    If mIsServer Then
    
        Set Screen.Picture = Board_white.Picture
        
        InitOK = UpdateUserState(usHosting)
        TheLocalPlayer = coWHITE
        'Ενημέρωση του server οτι ξεκίνησε η παρτίδα
        If InitOK Then SendCommand FormatCommand(mtClientStartGame, "")
           
    Else
        Set Screen.Picture = Board_black.Picture
    
        InitOK = UpdateUserState(usConnecting)
        TheLocalPlayer = coBLACK
        
        
        
    End If
    
    If InitOK Then
        Me.Show
    Else
        Unload Me
    End If
    
    
End Function

'Απόπειρα σύνδεσης σε μια παρτίδα
Private Function SocketConnect() As Boolean
    On Error GoTo ErrHandler

    gUserState = usConnecting
    
    Dim StartTimer
    StartTimer = Timer
    'Αναμονή μέχρι να γίνει σύνδεση ή timeout (45 δευτερόλεπτα)
    
    SendToGameSocket FormatCommand(mtConnectionRequest, GetUserNickname)

    
    Do Until gUserState <> usConnecting Or Timer > StartTimer + 45
        DoEvents
        If gShutdown Then Exit Do
        If mFormUnloaded Then Exit Do
    Loop
    If gUserState = usConnecting Then MyMsgbox "Η σύνδεση δεν είναι εφικτή. Ο παίκτης δεν απάντησε έγκαιρα.", vbInformation
    
    If gUserState = usPlayingClient Then
        SocketConnect = True
    Else
        SocketConnect = False
    End If
    Exit Function
    
ErrHandler:
        
    MyMsgbox vbCrLf & Err & " " & Err.Description & " (" & Err.Source & ")", vbCritical, "Παρουσιάστηκε πρόβλημα"
    
End Function

'Ξεκίνημα Listen, δηλαδή αναμονής για να συνδεθεί ένας client στην παρτίδα
Private Function SocketListen() As Boolean
    gUserState = usHosting
    SocketListen = True
End Function

Private Sub Form_Load()
    rtbHistory.Text = ""
    rtbChat.Text = ""
    mFormUnloaded = False
    cboChat.Enabled = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mFormUnloaded = True

    'Αν η παρτίδα διακοπεί, ο hosting client ενημερώνει τον server
    If mIsServer Then SendCommand FormatCommand(mtClientEndGame, "")

    If (mOpponentName <> "") And gUserState <> usConnecting And gUserState <> usHosting Then SendToGameSocket FormatCommand(mtGameStop, "")
    
    frmMain.cmdHostGame.Enabled = True

End Sub


Private Sub SocketClose()
    If gUserState = usPlayingClient Or gUserState = usPlayingServer Then
        gUserState = usIdle
        MyMsgbox "Ο αντίπαλος έκλεισε την σύνδεση.", vbInformation
        Unload Me
    End If
End Sub

'Αποστολή tcp/ip μυνήματος στον αντίπαλο
Private Function SendToGameSocket(ByVal Command As String)
    
    If mOpponentName = "" Then
        Err.Raise vbObjectError + 1, , "The opponent name is not known yet"
        Exit Function
    End If
    
    Dim MsgType As String
    Dim MsgCommand As String
    Dim ModifiedCommand As String
    
    MsgType = Mid(Split(Command, MSG_SEPERATOR)(0), 2)
    MsgCommand = Replace(Split(Command, MSG_SEPERATOR)(1), COMMAND_SEPERATOR, "")
    MsgCommand = Left(MsgCommand, Len(MsgCommand) - 1)
    
    ModifiedCommand = FormatCommand(MsgType, mOpponentName & CLIENT_SEPERATOR & MsgCommand)
    
    SendToSocket ModifiedCommand

End Function




'Μεταβολή της κατάστασης του local client
'
'usHosting - O client είναι ο hosting client και περιμένει αντίπαλο στην παρτίδα
'usConnecting - O client προσπαθεί να συνδεθεί στην παρτίδα ενός άλλου client
'usIdle - O client είναι ανενεργός και δεν κάνει τίποτα
'usPlayingClient - Ο client παίζει σε παρτίδα την οποία κάνει host άλλος client
'usPlayingServer - Ο client παίζει σε παρτίδα την οποία κάνει host ο ίδιος
Private Function UpdateUserState(ByVal NewState As enumUserState) As Boolean

    gUserState = NewState
    
    If gUserState = usHosting Then
    
        cmdExit.Enabled = True
        cboChat.Enabled = False
        cmdOfferDraw.Enabled = False
        cmdStats.Enabled = False
        cmdResign.Enabled = False
        
        frmMain.cmdHostGame.Enabled = False
        Caption = GetUserNickname & " vs (κενή θέση)"
        lblGameStatus = "Περιμένετε αντίπαλο ..."
        
        mOpponentName = ""
        
        UpdateUserState = SocketListen
        
    ElseIf gUserState = usConnecting Then
    
        frmMain.cmdHostGame.Enabled = False
        cboChat.Enabled = False
        cmdOfferDraw.Enabled = False
        cmdStats.Enabled = False
        cmdResign.Enabled = False
        
        Caption = mOpponentName & " vs " & GetUserNickname
        lblGameStatus = "Γίνεται σύνδεση με τον παίκτη " & mOpponentName & ". Παρακαλώ περιμένετε ..."
        
        Show
        
        UpdateUserState = SocketConnect()
    
    ElseIf gUserState = usIdle Then
    
        cmdExit.Enabled = True
        cboChat.Enabled = False
        frmMain.cmdHostGame.Enabled = True
        cmdOfferDraw.Enabled = False
        cmdStats.Enabled = False
        cmdResign.Enabled = False
        
        UpdateUserState = True
        
    ElseIf gUserState = usPlayingClient Then
    
        lblGameStatus = "Έγινε επιτυχής σύνδεση"
        cmdExit.Enabled = True
        frmMain.cmdHostGame.Enabled = False
        cmdOfferDraw.Enabled = True
        cmdStats.Enabled = True
        cmdResign.Enabled = True
        cboChat.Enabled = True
        
        Set mGame = New cChessGame
        mGame.Init coBLACK, gsPlaying, Screen, Sprites, Buffer, Board_black, Ruler1, Ruler2
        
        PrintChatMessage rtbChat, vbBlack, 8, False, False, "Η παρτίδα ξεκίνησε στις "
        PrintChatMessage rtbChat, vbBlack, 8, True, False, FormatDateTime(Now, vbShortTime) & vbCrLf
        
        UpdateUserState = True
        
    ElseIf gUserState = usPlayingServer Then
    
        Caption = GetUserNickname & " vs " & mOpponentName
        lblGameStatus = "Έγινε επιτυχής σύνδεση"
        cmdExit.Enabled = True
        cmdOfferDraw.Enabled = True
        cmdStats.Enabled = True
        cmdResign.Enabled = True
        frmMain.cmdHostGame.Enabled = False
        cboChat.Enabled = True
        
        Set mGame = New cChessGame
        mGame.Init coWHITE, gsPlaying, Screen, Sprites, Buffer, Board_white, Ruler1, Ruler2
        
        PrintChatMessage rtbChat, vbBlack, 8, False, False, "Η παρτίδα ξεκίνησε στις "
        PrintChatMessage rtbChat, vbBlack, 8, True, False, FormatDateTime(Now, vbShortTime) & vbCrLf
        
        UpdateUserState = True
        
    End If
    
End Function



'Η function καλείται για να επεξεργαστεί κάθε incoming tcp/ip μύνημα που στέλνεται στον client
'απο τον αντίπαλο του, κατα την διάρκεια της παρτίδας
Public Sub ProcessGameCommand(ByVal Command As String)

    Dim MsgType As Byte
    Dim Arg1 As String, Arg2 As String, Arg3 As String, Arg4 As String, Arg5 As String, Arg6 As String, Arg7 As String
    
    'Το μύνημα περιέχει μέσα και τον τύπο του μυνήματος
    MsgType = Split(Command, MSG_SEPERATOR)(0) 'Τύπος μυνήματος
    Command = Split(Command, MSG_SEPERATOR)(1) 'Περιεχόμενο μυνήματος
    
    Select Case MsgType
    
    Case mtConnectionAccepted ' Μύνημα που στέλνει ο hosting παίκτης στον client παίκτη, αφού ο 1ος δεχθεί την σύνδεση απο τον 2ο
        UpdateUserState usPlayingClient
    
    Case mtConnectionDenied ' Μύνημα που στέλνει o hosting παίκτης στον client παίκτη, όταν δεν θέλει να παίξει μαζί του
        MyMsgbox "Ο χρήστης " & mOpponentName & " αρνήθηκε να παίξει μαζί σας.", vbInformation + vbOKOnly
        gUserState = usIdle
        Unload Me
        
    Case mtGameSay ' Μύνημα chat που στέλνεται απο έναν παίκτη στον άλλον κατα την διάρκεια μιας παρτίδας
        ShowGameChatMessage Command
        
    Case mtGameMove ' Κίνηση που στέλνεται απο έναν παίκτη στον άλλον κατα την διάρκεια μιας παρτίδας
        Arg1 = Split(Command, ARGUMENT_SEPERATOR)(0) ' from x
        Arg2 = Split(Command, ARGUMENT_SEPERATOR)(1) ' from y
        Arg3 = Split(Command, ARGUMENT_SEPERATOR)(2) ' to x
        Arg4 = Split(Command, ARGUMENT_SEPERATOR)(3) ' to y
        Arg5 = Split(Command, ARGUMENT_SEPERATOR)(4) ' promote pawn to what
        Arg6 = Split(Command, ARGUMENT_SEPERATOR)(5) ' Castling move ?
        Arg7 = Split(Command, ARGUMENT_SEPERATOR)(6) ' en passant move?
        
        Dim StartTimer
        StartTimer = Timer
        mGame.BeginMove Arg1, Arg2
        lblGameStatus = "Ο αντίπαλος κάνει την κίνηση του ... "
        Do Until Timer > StartTimer + 0.7
            DoEvents
            If gShutdown Then Exit Do
            If mFormUnloaded Then Exit Do
        Loop
        mGame.CompleteMove Arg1, Arg2, Arg3, Arg4, Arg5, IIf(Arg6 = "1", True, False), Arg7
        
    Case mtGameResign ' Παραίτηση ενός παίκτη. Στέλνεται στον άλλον παίκτη κατα την διάρκεια μιας παρτίδας
        UpdateUserState usIdle
        GameResult True, "Ο αντίπαλος παραιτήθηκε."
        Unload Me
        
    Case mtGameOfferDraw ' Πρόταση ισοπαλίας απο έναν παίκτη στον άλλον
        If MyMsgbox("Ο αντίπαλος σας προτείνει να λήξει η παρτίδα με ΙΣΟΠΑΛΙΑ." & vbCrLf & vbCrLf & "Δέχεστε;", vbQuestion + vbYesNo) = vbNo Then
            SendToGameSocket FormatCommand(mtGameDenyDraw, "")
        Else
            SendCommand FormatCommand(mtDraw, Caption)
            SendToGameSocket FormatCommand(mtGameAcceptDraw, "")
            UpdateUserState usIdle
            GameResult True, , True
            Unload Me
        End If
        
    Case mtGameAcceptDraw ' Ο παίκτης δέχεται την προσφορά ισοπαλίας απτον άλλο παίκτη
        UpdateUserState usIdle
        GameResult True, , True
        Unload Me
        
    Case mtGameDenyDraw ' Ο παίκτης αρνείται την προσφορά ισοπαλίας απτον άλλο παίκτη
        MyMsgbox "Ο αντίπαλος δέν δέχεται να λήξει η παρτίδα με ΙΣΟΠΑΛΙΑ.", vbInformation + vbOKOnly
    
    Case mtConnectionRequest ' Ένας client θέλει να παίξει στην παρτίδα αυτού του client
        Arg1 = Command ' nickname του client που θέλει να συνδεθεί στην παρτίδα
        lblGameStatus = "Απόπειρα σύνδεσης απο τον χρήστη " & Arg1 & " ..."
        
        StartTimer = Timer

        If MyMsgbox("Ο χρήστης " & Arg1 & " θέλει να παίξει μαζί σας." & vbCrLf & vbCrLf & "Θέλετε να το επιτρέψετε;", vbQuestion + vbYesNo) = vbNo Then
            mOpponentName = Arg1
            SendToGameSocket FormatCommand(mtConnectionDenied, "")
            
            UpdateUserState usHosting
        Else
        
            If Timer < StartTimer + 44 Then
        
                'Αποδοχή σύνδεσης
                mOpponentName = Arg1
                SendToGameSocket FormatCommand(mtConnectionAccepted, "")
                lblGameStatus = "Γίνεται σύνδεση με τον χρήστη " & Arg1 & " ..."
                
                UpdateUserState usPlayingServer
                SendCommand FormatCommand(mtClientGameStarted, mOpponentName)
                
            Else
                MyMsgbox "Ο αντίπαλος δεν είναι πλέον διαθέσιμος. Παρακαλώ ξεκινήστε νέα παρτίδα.", vbOKOnly + vbInformation
                UpdateUserState usIdle
                Unload Me
            End If
            
        End If
    
    Case mtGameStop ' Μύνημα που στέλνει ένας client στον άλλον, όταν ο 1ος διακόπει την παρτίδα
        SocketClose
    End Select
    
End Sub


Private Sub ShowGameChatMessage(ByVal Command As String)
    Dim Arg1 As String, Arg2 As String
    Arg1 = Trim(Split(Command, ">")(0))
    Arg2 = Trim(Split(Command, ">")(1))
    
    PrintChatMessage rtbChat, vbBlue, 8, True, False, Arg1 & "> "
    PrintChatMessage rtbChat, vbBlack, 8, False, False, Arg2 & vbCrLf
End Sub
