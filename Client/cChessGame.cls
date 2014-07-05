VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cChessGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************
' This class contains the chess engine. The rules of the game
' are implemented here, like allowed moves, conditions
' for the game's end etc.'
'
' Copyright Theo Kandiliotis 2004
'*******************************************************

Option Explicit

Private Const SPRITE_SIZE = 35
Private Const SQUARE_SIZE = 44

Private Const BOARD_X = 20
Private Const BOARD_BOTTOM_Y = 326
Private Const BOARD_TOP_Y = 20

Public Enum enumChessPiece
    cpEmpty = 0
    cpPawn = 1
    cpKnight = 2
    cpBishop = 3
    cpRook = 4
    cpQueen = 5
    cpKing = 6
End Enum

Public Enum enumGameResult
    grYOU_WIN
    grYOU_LOOSE
    grSTALEMATE
End Enum
    
Public Enum enumPlayerPiece
    ppEmpty = 0
    
    ppWHITE_Pawn = 1
    ppWHITE_Knight = 2
    ppWHITE_Bishop = 3
    ppWHITE_Rook = 4
    ppWHITE_Queen = 5
    ppWHITE_King = 6
    
    ppBLACK_Pawn = -1
    ppBLACK_Knight = -2
    ppBLACK_Bishop = -3
    ppBLACK_Rook = -4
    ppBLACK_Queen = -5
    ppBLACK_King = -6
End Enum

Public Enum enumChessOpponent
    coWHITE = 1
    coBLACK = -1
End Enum

Public Enum enumGameState
    gsWaiting = 1
    gsPlaying = 2
    gsClosed = 3
End Enum

Private Type udtMove
    FromX As Byte
    FromY As Byte
    ToX As Byte
    ToY As Byte
    PawnPromotesTo As enumChessPiece
    Castling As Boolean
    EnPassant As Byte '0=no enpassant,1=en passant to left,2=en passant to right
End Type

Private mGameBoard() As enumPlayerPiece
Private mLocalPlayer As enumChessOpponent
Private mPlayerToPlay As enumChessOpponent
Private mMoveCount As Integer
Private mCapturedByWhite() As enumChessPiece
Private mCapturedByBlack() As enumChessPiece
Private mMoveHistory As String
Private mLastMove As udtMove


Private mWhiteKingMoved As Boolean
Private mWhiteRook1Moved As Boolean
Private mWhiteRook2Moved As Boolean
Private mBlackKingMoved As Boolean
Private mBlackRook1Moved As Boolean
Private mBlackRook2Moved As Boolean

Public GameState As enumGameState
Public MovesList As String

Private mScreen As PictureBox
Private mSprites As PictureBox
Private mBuffer As PictureBox
Private mBoard As PictureBox
Private mRuler1 As PictureBox
Private mRuler2 As PictureBox

Private mPlayerToPlayCanCastleRight As Boolean
Private mPlayerToPlayCanCastleLeft As Boolean

Event MustSendMove(ByVal FromX As Byte, ByVal FromY As Byte, ByVal ToX As Byte, ByVal ToY As Byte, ByVal PromoteToWhat As enumChessPiece, ByVal Castling As Boolean, ByVal EnPassant As Byte)
Event TurnChanged(ByVal PlayerToPlay As enumChessOpponent, ByVal LocalPlayer As enumChessOpponent, ByVal MoveMade As String)
Event GameEnded(ByVal Msg As String, ByVal GameResult As enumGameResult)
Event PawnPromotion(ByRef PromoteToWhat As enumChessPiece)

Private Function GetSpriteXMargin(ByVal aPiece As enumPlayerPiece) As Byte
    aPiece = Abs(aPiece)
    If aPiece = cpPawn Then
        GetSpriteXMargin = 0
    ElseIf aPiece = cpKnight Then
        GetSpriteXMargin = 1
    ElseIf aPiece = cpBishop Then
        GetSpriteXMargin = 2
    ElseIf aPiece = cpRook Then
        GetSpriteXMargin = 3
    ElseIf aPiece = cpQueen Then
        GetSpriteXMargin = 4
    ElseIf aPiece = cpKing Then
        GetSpriteXMargin = 5
    End If
    GetSpriteXMargin = GetSpriteXMargin * SPRITE_SIZE
    
End Function

Private Function getSpriteYMargin(ByVal aPlayer As enumChessOpponent) As Byte
    If aPlayer = coWHITE Then
        getSpriteYMargin = 0
    Else
        getSpriteYMargin = SPRITE_SIZE
    End If
End Function

Public Sub DrawBufferToScreen()
    mScreen.PaintPicture mBuffer.Image, 0, 0
End Sub

Private Sub DrawPiecesToBuffer()
    Dim X As Byte, y As Byte
    Dim DestX As Integer, DestY As Integer
    
    
    mBuffer.PaintPicture mBoard.Picture, 0, 0
    
    If mLocalPlayer = coWHITE Then
        mBuffer.PaintPicture mRuler1.Picture, BOARD_X - mRuler1.ScaleWidth - 5, BOARD_TOP_Y
    Else
        mBuffer.PaintPicture mRuler2.Picture, BOARD_X - mRuler2.ScaleWidth - 5, BOARD_TOP_Y
    End If
    
    For X = 1 To 8
        For y = 1 To 8
        
            If mGameBoard(X, y) <> ppEmpty Then
                DestX = BOARD_X + (X - 1) * SQUARE_SIZE + ((SQUARE_SIZE - SPRITE_SIZE) / 2)
                If mLocalPlayer = coWHITE Then
                    DestY = BOARD_BOTTOM_Y - (y - 1) * SQUARE_SIZE + ((SQUARE_SIZE - SPRITE_SIZE) / 2)
                Else
                    DestY = BOARD_TOP_Y + (y - 1) * SQUARE_SIZE + ((SQUARE_SIZE - SPRITE_SIZE) / 2)
                End If
                
                DrawSprite mGameBoard(X, y), DestX, DestY
            End If
        Next
    Next
        
End Sub

Private Sub DrawSprite(ByVal aPiece As enumPlayerPiece, ByVal DestX As Integer, ByVal DestY As Integer, Optional ByVal DestWidth As Integer = SPRITE_SIZE, Optional ByVal DestHeight As Integer = SPRITE_SIZE)
        
    Dim SpriteSrcX As Integer, SpriteSrcY As Integer
    Dim MaskSrcX As Integer, MaskSrcY As Integer
    Dim XMargin As Integer, YMargin As Integer
    
    XMargin = GetSpriteXMargin(aPiece)
    YMargin = getSpriteYMargin(Sgn(aPiece))
    
    MaskSrcX = XMargin
    MaskSrcY = 2 * SPRITE_SIZE
    
    SpriteSrcX = XMargin
    SpriteSrcY = YMargin
    
    mBuffer.PaintPicture mSprites.Picture, DestX, DestY, DestWidth, DestHeight, MaskSrcX, MaskSrcY, SPRITE_SIZE, SPRITE_SIZE, vbMergePaint
    mBuffer.PaintPicture mSprites.Picture, DestX, DestY, DestWidth, DestHeight, SpriteSrcX, SpriteSrcY, SPRITE_SIZE, SPRITE_SIZE, vbSrcAnd


End Sub
    

    
Public Sub Init(ByVal LocalPlayer As enumChessOpponent, ByVal GameState As enumGameState, ByVal Screen As PictureBox, ByVal Sprites As PictureBox, ByVal Buffer As PictureBox, ByVal Board As PictureBox, ByVal Ruler1 As PictureBox, ByVal Ruler2 As PictureBox)
    
    mLocalPlayer = LocalPlayer
    mPlayerToPlay = coWHITE
    
    Set mScreen = Screen
    Set mSprites = Sprites
    Set mBuffer = Buffer
    Set mBoard = Board
    Set mRuler1 = Ruler1
    Set mRuler2 = Ruler2
    
    Me.GameState = GameState
    
    ReDim mGameBoard(9, 9)
    ReDim mCapturedByWhite(0)
    ReDim mCapturedByBlack(0)
    
    Dim X As Byte, y As Byte
    For X = 1 To 8
        mGameBoard(X, 2) = ppWHITE_Pawn
        mGameBoard(X, 7) = ppBLACK_Pawn
    Next

    mGameBoard(1, 1) = ppWHITE_Rook
    mGameBoard(2, 1) = ppWHITE_Knight
    mGameBoard(3, 1) = ppWHITE_Bishop
    mGameBoard(4, 1) = ppWHITE_Queen
    mGameBoard(5, 1) = ppWHITE_King
    mGameBoard(6, 1) = ppWHITE_Bishop
    mGameBoard(7, 1) = ppWHITE_Knight
    mGameBoard(8, 1) = ppWHITE_Rook

    mGameBoard(1, 8) = ppBLACK_Rook
    mGameBoard(2, 8) = ppBLACK_Knight
    mGameBoard(3, 8) = ppBLACK_Bishop
    mGameBoard(4, 8) = ppBLACK_Queen
    mGameBoard(5, 8) = ppBLACK_King
    mGameBoard(6, 8) = ppBLACK_Bishop
    mGameBoard(7, 8) = ppBLACK_Knight
    mGameBoard(8, 8) = ppBLACK_Rook
    
    
    DrawPiecesToBuffer
    DrawBufferToScreen
    
    If GameState = gsPlaying Then RaiseTurnChanged mPlayerToPlay, mLocalPlayer, ""
    
    
End Sub

Private Sub RaiseTurnChanged(ByVal PlayerToPlay As enumChessOpponent, ByVal LocalPlayer As enumChessOpponent, ByVal MoveMade As String)
    RaiseEvent TurnChanged(PlayerToPlay, LocalPlayer, MoveMade)
    
    'Check if king, room or inbetween squares are safe
    
    'Get all possible moves by opponent
    Dim OpponentMovesList() As udtMove, KingY As Byte
    GetPlayerPossibleMoves mGameBoard, -PlayerToPlay, OpponentMovesList
    
    If PlayerToPlay = coWHITE Then
        KingY = 1
    Else
        KingY = 8
    End If
    
RightSide:

    mPlayerToPlayCanCastleRight = True
    If PlayerCanAttackHere(OpponentMovesList, mGameBoard, -PlayerToPlay, 5, KingY) Then mPlayerToPlayCanCastleRight = False: GoTo LeftSide
    If PlayerCanAttackHere(OpponentMovesList, mGameBoard, -PlayerToPlay, 6, KingY) Then mPlayerToPlayCanCastleRight = False: GoTo LeftSide
    If PlayerCanAttackHere(OpponentMovesList, mGameBoard, -PlayerToPlay, 7, KingY) Then mPlayerToPlayCanCastleRight = False: GoTo LeftSide
    If PlayerCanAttackHere(OpponentMovesList, mGameBoard, -PlayerToPlay, 8, KingY) Then mPlayerToPlayCanCastleRight = False: GoTo LeftSide
    
LeftSide:
    
    mPlayerToPlayCanCastleLeft = True
    If PlayerCanAttackHere(OpponentMovesList, mGameBoard, -PlayerToPlay, 5, KingY) Then mPlayerToPlayCanCastleLeft = False
    If PlayerCanAttackHere(OpponentMovesList, mGameBoard, -PlayerToPlay, 1, KingY) Then mPlayerToPlayCanCastleLeft = False
    If PlayerCanAttackHere(OpponentMovesList, mGameBoard, -PlayerToPlay, 2, KingY) Then mPlayerToPlayCanCastleLeft = False
    If PlayerCanAttackHere(OpponentMovesList, mGameBoard, -PlayerToPlay, 3, KingY) Then mPlayerToPlayCanCastleLeft = False
    If PlayerCanAttackHere(OpponentMovesList, mGameBoard, -PlayerToPlay, 4, KingY) Then mPlayerToPlayCanCastleLeft = False
    
    
    
End Sub

Private Function GetMoveDesc(aMove As udtMove)
    GetMoveDesc = mMoveCount & "." & Chr(Asc("a") + aMove.FromX - 1) & aMove.FromY & "-" & Chr(Asc("a") + aMove.ToX - 1) & aMove.ToY
    GetMoveDesc = UCase(GetMoveDesc)
End Function

Public Function MouseX_To_BoardX(ByVal MouseX As Integer) As Integer
    If MouseX < BOARD_X Or MouseX > (BOARD_X + 8 * SQUARE_SIZE) Then
        MouseX_To_BoardX = -1
        Exit Function
    End If
    MouseX_To_BoardX = (MouseX - BOARD_X) \ SQUARE_SIZE + 1
End Function
Public Function MouseY_To_BoardY(ByVal MouseY As Integer) As Integer
    If MouseY < BOARD_TOP_Y Or MouseY > BOARD_BOTTOM_Y + SQUARE_SIZE Then
        MouseY_To_BoardY = -1
        Exit Function
    End If
    MouseY_To_BoardY = 8 - (MouseY - BOARD_TOP_Y) \ SQUARE_SIZE
    If mLocalPlayer = coBLACK Then MouseY_To_BoardY = 9 - MouseY_To_BoardY
End Function
Public Function BoardX_To_MouseX(ByVal BoardX As Byte) As Integer
    If BoardX < 1 Or BoardX > 8 Then
        BoardX_To_MouseX = -1
        Exit Function
    End If
    BoardX_To_MouseX = BOARD_X + (BoardX - 1) * SQUARE_SIZE
End Function
Public Function BoardY_To_MouseY(ByVal BoardY As Byte) As Integer
    If BoardY < 1 Or BoardY > 8 Then
        BoardY_To_MouseY = -1
        Exit Function
    End If
    
    BoardY = 9 - BoardY
    
    If mLocalPlayer = coWHITE Then
        BoardY_To_MouseY = BOARD_TOP_Y + (BoardY - 1) * SQUARE_SIZE
    Else
        BoardY_To_MouseY = BOARD_BOTTOM_Y - (BoardY - 1) * SQUARE_SIZE + 1
    End If
End Function

Private Function NewMove(ByVal FromX As Integer, ByVal FromY As Integer, ByVal ToX As Integer, ByVal ToY As Integer, Optional ByVal PawnPromotesTo As enumChessPiece = cpEmpty, Optional ByVal Castling As Boolean = False, Optional ByVal EnPassant As Byte = 0) As udtMove
    
    If Not ((FromX < 1) Or (FromX > 8) Or (FromY < 1) Or (FromY > 8) Or (ToX < 1) Or (ToX > 8) Or (ToY < 1) Or (ToY > 8)) Then
        With NewMove
        .FromX = FromX
        .FromY = FromY
        .ToX = ToX
        .ToY = ToY
        .PawnPromotesTo = PawnPromotesTo
        .Castling = Castling
        .EnPassant = EnPassant
        End With
        
    End If
     
End Function

Private Function IsValidMove(GameBoard() As enumPlayerPiece, aMove As udtMove) As Boolean
    With aMove

    If .FromX < 1 Or .FromX > 8 Or .FromY < 1 Or .FromY > 8 Or .ToX < 1 Or .ToX > 8 Or .ToY < 1 Or .ToY > 8 Then
        IsValidMove = False
        Exit Function
    End If
    
    If Sgn(GameBoard(.ToX, .ToY)) = Sgn(GameBoard(.FromX, .FromY)) Then
        IsValidMove = False
        Exit Function
    End If
    
    End With
    
    IsValidMove = True
    
End Function

Private Sub AddMove(GameBoard() As enumPlayerPiece, ByRef MovesList() As udtMove, aMove As udtMove)
    If IsValidMove(GameBoard, aMove) Then
        ReDim Preserve MovesList(UBound(MovesList) + 1)
        MovesList(UBound(MovesList)) = aMove
    End If
End Sub

Private Sub GetPlayerPossibleMoves(GameBoard() As enumPlayerPiece, ByVal Player As enumChessOpponent, ByRef MovesList() As udtMove)
    
    ReDim MovesList(0)
    Dim X As Byte, y As Byte
    For X = 1 To 8
        For y = 1 To 8
            If Sgn(GameBoard(X, y)) = Player Then
                GetPossibleMoves GameBoard, X, y, MovesList
            End If
        Next
    Next

End Sub

Private Function MoveAftermath() As Boolean

    Dim PlayerThatMoved As enumChessOpponent, MovesList() As udtMove, i As Integer
    Dim IsCheck As Boolean, IsCheckMate As Boolean, IsDraw As Boolean
    Dim GameBoard(9, 9) As enumPlayerPiece

    PlayerThatMoved = -mPlayerToPlay

    IsCheck = PlayerIsInCheck(mGameBoard, -PlayerThatMoved)
    
    'Check if this moves causes a CHECK
    If IsCheck = True Then
        'Check if this move causes a CHECK MATE
        IsCheckMate = PlayerIsInCheckMate(mGameBoard, -PlayerThatMoved)
    Else
        'Check if move causes a DRAW
        IsDraw = PlayerHasNoMove(mGameBoard, -PlayerThatMoved)
    End If

    If IsCheckMate Then
        MoveAftermath = True
        If PlayerThatMoved = mLocalPlayer Then
            RaiseEvent GameEnded("ΜΠΡΑΒΟ! " & vbCrLf & vbCrLf & "Κάνατε ΡΟΥΑ ΜΑΤ." & vbCrLf & vbCrLf & "Συγχαρητήρια, ΚΕΡΔΙΣΑΤΕ !", grYOU_WIN)
        Else
            RaiseEvent GameEnded("ΧΑΣΑΤΕ! " & vbCrLf & vbCrLf & "Ο αντίπαλος έκανε ΡΟΥΑ ΜΑΤ.", grYOU_LOOSE)
        End If
    ElseIf IsCheck Then
        If PlayerThatMoved = mLocalPlayer Then
            MyMsgbox "Κάνατε ΡΟΥΑ στον αντίπαλο.", vbInformation
        Else
            MyMsgbox "ΡΟΥΑ! Ο βασιλιάς σας είναι σε κίνδυνο.", vbInformation
        End If
    ElseIf IsDraw Then
        MoveAftermath = True
        If PlayerThatMoved = mLocalPlayer Then
            RaiseEvent GameEnded("ΙΣΟΠΑΛΙΑ!" & vbCrLf & vbCrLf & "Ο αντίπαλος δεν έχει καμία κίνηση να κάνει." & vbCrLf & vbCrLf & "Η παρτίδα λήγει σε ισοπαλια.", grSTALEMATE)
        Else
            RaiseEvent GameEnded("ΙΣΟΠΑΛΙΑ!" & vbCrLf & vbCrLf & "Δέν έχετε κίνηση να κάνετε." & vbCrLf & vbCrLf & "Η παρτίδα λήγει σε ισοπαλια.", grSTALEMATE)
        End If
    End If

End Function

'Returns TRUE if the only available moves put the player in a CHECK.
'In this case, there is a DRAW
Private Function PlayerHasNoMove(GameBoard() As enumPlayerPiece, ByVal Player As enumChessOpponent) As Boolean
    Dim MovesList() As udtMove, i As Integer
    Dim MemoryBoard() As enumPlayerPiece
    
    PlayerHasNoMove = True
    
    GetPlayerPossibleMoves GameBoard, Player, MovesList
    For i = 1 To UBound(MovesList)
        MemoryBoard = GameBoard
        ApplyMoveToBoard MemoryBoard, MovesList(i)
        If Not PlayerIsInCheck(MemoryBoard, Player) Then
            PlayerHasNoMove = False
            Exit Function
        End If
    Next
    
End Function
Private Function PlayerIsInCheckMate(GameBoard() As enumPlayerPiece, ByVal Player As enumChessOpponent) As Boolean
    
    Dim MovesList() As udtMove, i As Integer
    Dim MemoryBoard() As enumPlayerPiece
    
    PlayerIsInCheckMate = True
    
    'get all player's possible moves
    GetPlayerPossibleMoves GameBoard, Player, MovesList
    For i = 1 To UBound(MovesList)
        MemoryBoard = GameBoard
        ApplyMoveToBoard MemoryBoard, MovesList(i)
        If Not PlayerIsInCheck(MemoryBoard, Player) Then
            PlayerIsInCheckMate = False
            Exit Function
        End If
    Next
    
End Function

Private Function PlayerIsInCheck(GameBoard() As enumPlayerPiece, ByVal Player As enumChessOpponent) As Boolean
    
    Dim i As Integer, MovesList() As udtMove
    
    'Get all the opponent's possible moves
    GetPlayerPossibleMoves GameBoard, -Player, MovesList
    
    'Check if the player is in check, if his king is threatened
    For i = 1 To UBound(MovesList)
        If GameBoard(MovesList(i).ToX, MovesList(i).ToY) = Player * cpKing Then
            PlayerIsInCheck = True
            Exit Function
        End If
    Next
    
End Function

'Returns TRUE if the move (aMove) of player Player) on the chessboard (GameBoard), puts the king in danger
'In this case the move is not allowed.
Private Function PlayerMovesIntoCheck(GameBoard() As enumPlayerPiece, ByVal Player As enumChessOpponent, TheNewMove As udtMove) As Boolean
    Dim MemoryGameBoard() As enumPlayerPiece
    MemoryGameBoard = GameBoard
    ApplyMoveToBoard MemoryGameBoard, TheNewMove
    PlayerMovesIntoCheck = PlayerIsInCheck(MemoryGameBoard, Player)
    
End Function

'All the rules of permitted moves in a chess game are implemented in this method.
'Populate array MovesList with all possible moves of the piece at X,Y
Private Sub GetPossibleMoves(GameBoard() As enumPlayerPiece, ByVal X As Byte, ByVal y As Byte, ByRef MovesList() As udtMove)
    
    Dim Player As enumChessOpponent, Piece As enumChessPiece
    Dim aMove As udtMove, xi As Integer, yi As Integer
    
    If GameBoard(X, y) = ppEmpty Then Exit Sub
    Player = Sgn(GameBoard(X, y))
    Piece = Abs(GameBoard(X, y))
    
    If Piece = cpPawn Then
        
        xi = X: yi = y + Player * 1
        If GameBoard(xi, yi) = ppEmpty Then AddMove GameBoard, MovesList, NewMove(X, y, xi, yi)
        
        xi = X - 1: yi = y + Player * 1
        If Sgn(GameBoard(xi, yi)) = -Player Then AddMove GameBoard, MovesList, NewMove(X, y, xi, yi)
        
        xi = X + 1: yi = y + Player * 1
        If Sgn(GameBoard(xi, yi)) = -Player Then AddMove GameBoard, MovesList, NewMove(X, y, xi, yi)
        
        
        If (Player = coWHITE And y = 2) Or (Player = coBLACK And y = 7) Then
            xi = X: yi = y + Player * 2
            If GameBoard(xi, yi) = ppEmpty And GameBoard(xi, yi - Player) = ppEmpty Then AddMove GameBoard, MovesList, NewMove(X, y, xi, yi)
        End If
        
        'en passant to the left?
        xi = X - 1: yi = y
        If (GameBoard(xi, yi) = -Player * cpPawn) And _
            (mLastMove.FromX = xi And mLastMove.FromY = yi + Player * 2) And _
            (mLastMove.ToX = xi And mLastMove.ToY = yi) And _
            GameBoard(xi, yi + Player * 1) = ppEmpty Then
                        
            AddMove GameBoard, MovesList, NewMove(X, y, xi, yi + Player * 1, , , 1)
                        
        End If
        
        
        'en passant to the right?
        xi = X + 1: yi = y
        If (GameBoard(xi, yi) = -Player * cpPawn) And _
            (mLastMove.FromX = xi And mLastMove.FromY = yi + Player * 2) And _
            (mLastMove.ToX = xi And mLastMove.ToY = yi) And _
            GameBoard(xi, yi + Player * 1) = ppEmpty Then
                        
            AddMove GameBoard, MovesList, NewMove(X, y, xi, yi + Player * 1, , , 2)
                        
        End If
        
        
        
    End If
    
    If Piece = cpBishop Or Piece = cpQueen Then
        
        yi = y - 1
        For xi = X - 1 To 1 Step -1
            AddMove GameBoard, MovesList, NewMove(X, y, xi, yi)
            If Sgn(GameBoard(xi, yi)) <> 0 Then Exit For
            yi = yi - 1
            If yi <= 0 Then Exit For
         Next
            
        yi = y + 1
        For xi = X - 1 To 1 Step -1
            AddMove GameBoard, MovesList, NewMove(X, y, xi, yi)
            If Sgn(GameBoard(xi, yi)) <> 0 Then Exit For
            yi = yi + 1
            If yi >= 9 Then Exit For
        Next
        
        yi = y - 1
        For xi = X + 1 To 8
            AddMove GameBoard, MovesList, NewMove(X, y, xi, yi)
            If Sgn(GameBoard(xi, yi)) <> 0 Then Exit For
            yi = yi - 1
            If yi <= 0 Then Exit For
        Next
            
        yi = y + 1
        For xi = X + 1 To 8
            AddMove GameBoard, MovesList, NewMove(X, y, xi, yi)
            If Sgn(GameBoard(xi, yi)) <> 0 Then Exit For
            yi = yi + 1
            If yi >= 9 Then Exit For
        Next

    End If
    
    If Piece = cpRook Or Piece = cpQueen Then
        yi = y
        For xi = X - 1 To 1 Step -1
            AddMove GameBoard, MovesList, NewMove(X, y, xi, yi)
            If Sgn(GameBoard(xi, yi)) <> 0 Then Exit For
        Next
        For xi = X + 1 To 8
            AddMove GameBoard, MovesList, NewMove(X, y, xi, yi)
            If Sgn(GameBoard(xi, yi)) <> 0 Then Exit For
        Next
        
        xi = X
        For yi = y - 1 To 1 Step -1
            AddMove GameBoard, MovesList, NewMove(X, y, xi, yi)
            If Sgn(GameBoard(xi, yi)) <> 0 Then Exit For
        Next
        For yi = y + 1 To 8
            AddMove GameBoard, MovesList, NewMove(X, y, xi, yi)
            If Sgn(GameBoard(xi, yi)) <> 0 Then Exit For
        Next
    End If
    
    If Piece = cpKnight Then
        AddMove GameBoard, MovesList, NewMove(X, y, X + 1, y + 2)
        AddMove GameBoard, MovesList, NewMove(X, y, X + 1, y - 2)
        AddMove GameBoard, MovesList, NewMove(X, y, X - 1, y + 2)
        AddMove GameBoard, MovesList, NewMove(X, y, X - 1, y - 2)
        AddMove GameBoard, MovesList, NewMove(X, y, X + 2, y + 1)
        AddMove GameBoard, MovesList, NewMove(X, y, X + 2, y - 1)
        AddMove GameBoard, MovesList, NewMove(X, y, X - 2, y + 1)
        AddMove GameBoard, MovesList, NewMove(X, y, X - 2, y - 1)
    End If
        
    If Piece = cpKing Then
        AddMove GameBoard, MovesList, NewMove(X, y, X - 1, y - 1)
        AddMove GameBoard, MovesList, NewMove(X, y, X, y - 1)
        AddMove GameBoard, MovesList, NewMove(X, y, X + 1, y - 1)
        
        AddMove GameBoard, MovesList, NewMove(X, y, X - 1, y + 1)
        AddMove GameBoard, MovesList, NewMove(X, y, X, y + 1)
        AddMove GameBoard, MovesList, NewMove(X, y, X + 1, y + 1)
        
        AddMove GameBoard, MovesList, NewMove(X, y, X - 1, y)
        AddMove GameBoard, MovesList, NewMove(X, y, X + 1, y)
        
        Dim CastlingMove As udtMove, CanCastle As Boolean
        
        'Can it do castling move from left?
        CanCastle = PlayerCanDoCastlingMove(GameBoard, Player, CastlingMove, True)
        If CanCastle Then
            AddMove GameBoard, MovesList, CastlingMove
        End If

        'Can it do castling move from right?
        CanCastle = PlayerCanDoCastlingMove(GameBoard, Player, CastlingMove, False)
        If CanCastle Then
            AddMove GameBoard, MovesList, CastlingMove
        End If

    End If
    
End Sub

'Returns TRUE if a player can do this castling move, and save the move in variable TheMove
Private Function PlayerCanDoCastlingMove(GameBoard() As enumChessPiece, ByVal Player As enumChessOpponent, ByRef TheMove As udtMove, ByVal ToLeftSide As Boolean) As Boolean
    
    If ToLeftSide Then
        If mPlayerToPlayCanCastleLeft = False Then Exit Function
    Else
        If mPlayerToPlayCanCastleRight = False Then Exit Function
    End If
    
    
    'Check if the king or rook have been moved before
    If Player = coWHITE Then
        If ToLeftSide = True Then
            If mWhiteKingMoved Or mWhiteRook1Moved Then Exit Function
        Else
            If mWhiteKingMoved Or mWhiteRook2Moved Then Exit Function
        End If
    Else
        If ToLeftSide = True Then
            If mBlackKingMoved Or mBlackRook1Moved Then Exit Function
        Else
            If mBlackKingMoved Or mBlackRook2Moved Then Exit Function
        End If
    End If
    
    'Get king's coordinates
    Dim KingX As Byte, KingY As Byte
    KingX = 5
    If Player = coWHITE Then
        KingY = 1
    Else
        KingY = 8
    End If
    
    'Check if inbetween squares are empty
    If ToLeftSide = True Then
        If (GameBoard(4, KingY) <> cpEmpty) Or (GameBoard(3, KingY) <> cpEmpty) Or (GameBoard(2, KingY) <> cpEmpty) Then Exit Function
    Else
        If (GameBoard(6, KingY) <> cpEmpty) Or (GameBoard(7, KingY) <> cpEmpty) Then Exit Function
    End If
    
    PlayerCanDoCastlingMove = True
    
    TheMove.Castling = True
    TheMove.FromX = KingX
    TheMove.FromY = KingY
    TheMove.ToY = KingY
    If ToLeftSide Then
        TheMove.ToX = KingX - 2
    Else
        TheMove.ToX = KingX + 2
    End If
    
End Function

Private Sub HighlightSquare(ByVal SquareX As Byte, ByVal SquareY As Byte, ByVal Color As ColorConstants)
    Dim X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer
    X1 = BoardX_To_MouseX(SquareX)
    Y1 = BoardY_To_MouseY(SquareY)
    X2 = X1 + SQUARE_SIZE
    Y2 = Y1 + SQUARE_SIZE
    
    mScreen.DrawWidth = 2
    mScreen.Line (X1 + 3, Y1 + 3)-(X2 - 2, Y2 - 3), Color, B
    
End Sub
Private Function DrawPossibleMoves(ByVal X As Byte, ByVal y As Byte) As Boolean
    Dim MovesList() As udtMove, i As Integer
    ReDim MovesList(0)
    GetPossibleMoves mGameBoard, X, y, MovesList
    
    For i = 1 To UBound(MovesList)
        DrawPossibleMoves = True
        HighlightSquare MovesList(i).ToX, MovesList(i).ToY, vbYellow
    Next
End Function

Public Function BeginMove(ByVal X As Byte, ByVal y As Byte) As Boolean
    
    If GameState <> gsPlaying Then Exit Function
    
    If Sgn(mGameBoard(X, y)) <> mPlayerToPlay Then Exit Function
    
    DrawBufferToScreen
    BeginMove = DrawPossibleMoves(X, y)
    
End Function

Public Function CompleteMove(ByVal FromX As Byte, ByVal FromY As Byte, ByVal ToX As Byte, ByVal ToY As Byte, Optional ByVal PawnPromotesTo As enumChessPiece = cpEmpty, Optional ByVal Castling As Boolean = False, Optional ByVal EnPassant As Byte = 0) As Boolean
    
    Dim MovesList() As udtMove, i As Integer, TheNewMove As udtMove
    
    ReDim MovesList(0)
    GetPossibleMoves mGameBoard, FromX, FromY, MovesList
    
    If Abs(mGameBoard(FromX, FromY)) = cpKing And Abs(CLng(ToX) - CLng(FromX)) > 1 Then Castling = True
    If Abs(mGameBoard(FromX, FromY)) = cpPawn And mGameBoard(ToX, ToY) = ppEmpty And ToX <> FromX Then
        If ToX > FromX Then
            EnPassant = 2
        ElseIf ToX < FromX Then
            EnPassant = 1
        End If
    End If
    
        
        
    TheNewMove = NewMove(FromX, FromY, ToX, ToY, PawnPromotesTo, Castling, EnPassant)
    If MovesListContainsMove(MovesList, TheNewMove) Then
    
        If (PawnMustPromote(mGameBoard, TheNewMove)) And (TheNewMove.PawnPromotesTo = cpEmpty) Then
            RaiseEvent PawnPromotion(PawnPromotesTo)
            TheNewMove.PawnPromotesTo = PawnPromotesTo
        End If
    
        If PlayerMovesIntoCheck(mGameBoard, mPlayerToPlay, TheNewMove) Then
            MyMsgbox "Η κίνηση αυτή δεν επιτρέπεται γιατι βάζει τον βασιλιά σε κίνδυνο!", vbExclamation
            CompleteMove = False
            Exit Function
        Else
            MakeMove TheNewMove
            CompleteMove = True
        End If
        
    End If
    
End Function

'Returns TRUE if this move is a pawn move towars the last row of the chess board
'In this case the pawn can be promoted to either Bishop, Rook, knight or Queen
Private Function PawnMustPromote(GameBoard() As enumPlayerPiece, aMove As udtMove) As Boolean
    If Abs(GameBoard(aMove.FromX, aMove.FromY)) <> cpPawn Then Exit Function
    If (Not (Sgn(GameBoard(aMove.FromX, aMove.FromY)) = coWHITE And aMove.ToY = 8)) And (Not (Sgn(GameBoard(aMove.FromX, aMove.FromY)) = coBLACK And aMove.ToY = 1)) Then Exit Function
    
    PawnMustPromote = True
    
End Function
Private Function MovesListContainsMove(MovesList() As udtMove, aMove As udtMove) As Boolean
    Dim i As Integer
    For i = 1 To UBound(MovesList)
        If CompareMoves(MovesList(i), aMove) Then
            MovesListContainsMove = True
            Exit Function
        End If
    Next
End Function

Private Function CompareMoves(move1 As udtMove, move2 As udtMove) As Boolean
    If move1.FromX <> move2.FromX Then CompareMoves = False: Exit Function
    If move1.ToX <> move2.ToX Then CompareMoves = False: Exit Function
    If move1.FromY <> move2.FromY Then CompareMoves = False: Exit Function
    If move1.ToY <> move2.ToY Then CompareMoves = False: Exit Function
    
    CompareMoves = True
End Function
Private Sub MakeMove(aMove As udtMove)
    
    Dim dummy As enumPlayerPiece
    dummy = ApplyMoveToBoard(mGameBoard, aMove)
    
    If dummy <> ppEmpty Then
        If mPlayerToPlay = coWHITE Then
            ReDim Preserve mCapturedByWhite(UBound(mCapturedByWhite) + 1)
            mCapturedByWhite(UBound(mCapturedByWhite)) = Abs(dummy)
        Else
            ReDim Preserve mCapturedByBlack(UBound(mCapturedByBlack) + 1)
            mCapturedByBlack(UBound(mCapturedByBlack)) = Abs(dummy)
        End If
    End If
    
    If mPlayerToPlay = mLocalPlayer Then RaiseEvent MustSendMove(aMove.FromX, aMove.FromY, aMove.ToX, aMove.ToY, aMove.PawnPromotesTo, aMove.Castling, aMove.EnPassant)
    
    If mGameBoard(aMove.ToX, aMove.ToY) = ppWHITE_King Then
        mWhiteKingMoved = True
    ElseIf mGameBoard(aMove.ToX, aMove.ToY) = ppBLACK_King Then
        mBlackKingMoved = True
    ElseIf mGameBoard(aMove.ToX, aMove.ToY) = ppBLACK_Rook And aMove.FromX = 1 Then
        mBlackRook1Moved = True
    ElseIf mGameBoard(aMove.ToX, aMove.ToY) = ppBLACK_Rook And aMove.FromX = 8 Then
        mBlackRook2Moved = True
    ElseIf mGameBoard(aMove.ToX, aMove.ToY) = ppWHITE_Rook And aMove.FromX = 1 Then
        mWhiteRook1Moved = True
    ElseIf mGameBoard(aMove.ToX, aMove.ToY) = ppWHITE_Rook And aMove.FromX = 8 Then
        mWhiteRook2Moved = True
    End If
        
    mPlayerToPlay = -mPlayerToPlay
    
    DrawPiecesToBuffer
    DrawBufferToScreen
    
    If mPlayerToPlay = coWHITE Then mMoveCount = mMoveCount + 1
    
    If MoveAftermath = True Then
        GameState = gsClosed
        Exit Sub
    End If
    
    mLastMove = aMove
    
    'MovesCollection.Add aMove
    
    Dim TheMoveDesc As String
    TheMoveDesc = GetMoveDesc(aMove)
    RaiseTurnChanged mPlayerToPlay, mLocalPlayer, TheMoveDesc
    mMoveHistory = mMoveHistory & TheMoveDesc & vbCrLf
    
End Sub

'Make this move to the specified board
Private Function ApplyMoveToBoard(ByRef GameBoard() As enumPlayerPiece, aMove As udtMove) As enumPlayerPiece
    ApplyMoveToBoard = GameBoard(aMove.ToX, aMove.ToY)
    GameBoard(aMove.ToX, aMove.ToY) = GameBoard(aMove.FromX, aMove.FromY)
    GameBoard(aMove.FromX, aMove.FromY) = ppEmpty
    
    If aMove.PawnPromotesTo <> cpEmpty Then GameBoard(aMove.ToX, aMove.ToY) = Sgn(GameBoard(aMove.ToX, aMove.ToY)) * aMove.PawnPromotesTo
    
    'Move rook also if this is a castling move
    If aMove.Castling = True Then
        Dim RookDestX As Byte, RookDestY As Byte, RookX As Byte, RookY As Byte
        If Sgn(GameBoard(aMove.ToX, aMove.ToY)) = coWHITE Then
            RookDestY = 1
            RookY = 1
        Else
            RookDestY = 8
            RookY = 8
        End If
        If aMove.ToX = 7 Then
            RookDestX = 6
            RookX = 8
        Else
            RookDestX = 4
            RookX = 1
        End If
        GameBoard(RookX, RookY) = ppEmpty
        GameBoard(RookDestX, RookDestY) = Sgn(GameBoard(aMove.ToX, aMove.ToY)) * cpRook
        
    End If
    
    'Remove opponent pawn if the move is en passant
    Dim DeadPawnX As Byte, DeadPawnY As Byte
    If aMove.EnPassant <> 0 Then
        DeadPawnY = aMove.FromY
        If aMove.EnPassant = 1 Then 'en passant αριστερά
            DeadPawnX = aMove.FromX - 1
        ElseIf aMove.EnPassant = 2 Then 'en passant δεξιά
            DeadPawnX = aMove.FromX + 1
        End If
        
        GameBoard(DeadPawnX, DeadPawnY) = ppEmpty
'        If mPlayerToPlay = coWHITE Then
'            ReDim Preserve mCapturedByWhite(UBound(mCapturedByWhite) + 1)
'            mCapturedByWhite(UBound(mCapturedByWhite)) = cpPawn
'        Else
'            ReDim Preserve mCapturedByBlack(UBound(mCapturedByBlack) + 1)
'            mCapturedByBlack(UBound(mCapturedByBlack)) = cpPawn
'        End If
        If mPlayerToPlay = coWHITE Then
            ApplyMoveToBoard = ppBLACK_Pawn
        Else
            ApplyMoveToBoard = ppWHITE_Pawn
        End If
        
    End If
    
End Function

Public Function GetStats() As String
    Dim Pieces() As enumChessPiece, i As enumChessPiece
    Dim PiecesCount(1 To 6) As Integer, CapturedList As String, LostList As String
    
    If mLocalPlayer = coWHITE Then
        Pieces = mCapturedByWhite
    Else
        Pieces = mCapturedByBlack
    End If
    For i = 1 To UBound(PiecesCount)
        PiecesCount(i) = 0
    Next
    For i = 1 To UBound(Pieces)
        PiecesCount(Pieces(i)) = PiecesCount(Pieces(i)) + 1
    Next
    
    For i = 1 To 6
        If i = cpPawn Then
            If PiecesCount(i) > 1 Then
                CapturedList = CapturedList & PiecesCount(i) & " πιόνια" & vbCrLf
            ElseIf PiecesCount(i) = 1 Then
                CapturedList = CapturedList & PiecesCount(i) & " πιόνι" & vbCrLf
            End If
        End If
        If i = cpKnight Then
            If PiecesCount(i) > 1 Then
                CapturedList = CapturedList & PiecesCount(i) & " άλογα" & vbCrLf
            ElseIf PiecesCount(i) = 1 Then
                CapturedList = CapturedList & PiecesCount(i) & " άλογο" & vbCrLf
            End If
        End If
        If i = cpBishop Then
            If PiecesCount(i) > 1 Then
                CapturedList = CapturedList & PiecesCount(i) & " αξιωματικούς" & vbCrLf
            ElseIf PiecesCount(i) = 1 Then
                CapturedList = CapturedList & PiecesCount(i) & " αξιωματικό" & vbCrLf
            End If
        End If
        If i = cpRook Then
            If PiecesCount(i) > 1 Then
                CapturedList = CapturedList & PiecesCount(i) & " πύργους" & vbCrLf
            ElseIf PiecesCount(i) = 1 Then
                CapturedList = CapturedList & PiecesCount(i) & " πύργο" & vbCrLf
            End If
        End If
        If i = cpQueen Then
            If PiecesCount(i) = 1 Then
                CapturedList = CapturedList & PiecesCount(i) & " βασίλισσα" & vbCrLf
            End If
        End If
    Next
    
    GetStats = "Έχεις ΚΑΤΑΚΤΗΣΕΙ τα παρακάτω:" & vbCrLf & vbCrLf & CapturedList
    
    If mLocalPlayer = coWHITE Then
        Pieces = mCapturedByBlack
    Else
        Pieces = mCapturedByWhite
    End If
    For i = 1 To UBound(PiecesCount)
        PiecesCount(i) = 0
    Next
    For i = 1 To UBound(Pieces)
        PiecesCount(Pieces(i)) = PiecesCount(Pieces(i)) + 1
    Next
    
    For i = 1 To 6
        If i = cpPawn Then
            If PiecesCount(i) > 1 Then
                LostList = LostList & PiecesCount(i) & " πιόνια" & vbCrLf
            ElseIf PiecesCount(i) = 1 Then
                LostList = LostList & PiecesCount(i) & " πιόνι" & vbCrLf
            End If
        End If
        If i = cpKnight Then
            If PiecesCount(i) > 1 Then
                LostList = LostList & PiecesCount(i) & " άλογα" & vbCrLf
            ElseIf PiecesCount(i) = 1 Then
                LostList = LostList & PiecesCount(i) & " άλογο" & vbCrLf
            End If
        End If
        If i = cpBishop Then
            If PiecesCount(i) > 1 Then
                LostList = LostList & PiecesCount(i) & " αξιωματικούς" & vbCrLf
            ElseIf PiecesCount(i) = 1 Then
                LostList = LostList & PiecesCount(i) & " αξιωματικό" & vbCrLf
            End If
        End If
        If i = cpRook Then
            If PiecesCount(i) > 1 Then
                LostList = LostList & PiecesCount(i) & " πύργους" & vbCrLf
            ElseIf PiecesCount(i) = 1 Then
                LostList = LostList & PiecesCount(i) & " πύργο" & vbCrLf
            End If
        End If
        If i = cpQueen Then
            If PiecesCount(i) = 1 Then
                LostList = LostList & PiecesCount(i) & " βασίλισσα" & vbCrLf
            End If
        End If
    Next
    
    GetStats = GetStats & vbCrLf & vbCrLf & "Έχεις ΧΑΣΕΙ τα παρακάτω:" & vbCrLf & vbCrLf & LostList
    
End Function



Private Function PlayerCanAttackHere(PlayerMoves() As udtMove, GameBoard() As enumPlayerPiece, ByVal AttackingPlayer As enumChessOpponent, ByVal ToX As Byte, ByVal ToY As Byte) As Boolean

    PlayerCanAttackHere = False
    
    Dim i As Integer

    For i = 1 To UBound(PlayerMoves)
        If PlayerMoves(i).ToX = ToX And PlayerMoves(i).ToY = ToY Then
            PlayerCanAttackHere = True
            Exit Function
        End If
    Next

End Function


Private Sub Class_Initialize()
    'Set MovesCollection = Nothing
End Sub
