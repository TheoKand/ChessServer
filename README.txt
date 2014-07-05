ChessServer
===========

Visual Basic 6 Chess server with Winsock

This is a fully functional chess server developed by me in 2004 using Visual Basic 6 and the Winsock control (MSWINSCK.OCX)

It consists of two seperate projects:

- The server is running on a computer with a static ip and listening on a specified port for incoming connections. 
- The client can be launched on any machine and it connects to the server in order to join the server lobby. 

In the lobby clients can exchange chat messages and arrange games.

A client can "host" a game so that other clients can connect to it to play a game of chess.

The complete and official gameplay of chess is implemented inside the client project (class cChessGame.cls) and even rules like en-passant, castling moves and pawn promotions are supported.

Client Screenshot :
https://github.com/TheoKand/ChessServer/blob/master/screenshots/Client.png

Server screenshot :
https://github.com/TheoKand/ChessServer/blob/master/screenshots/Server.png
