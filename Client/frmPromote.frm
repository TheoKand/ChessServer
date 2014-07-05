VERSION 5.00
Begin VB.Form frmPromote 
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   3870
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Ο στρατιώτης θα προβιβαστεί σε ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3675
      Begin VB.OptionButton optPiece 
         Caption         =   "Option1"
         Height          =   285
         Index           =   3
         Left            =   2460
         TabIndex        =   5
         Tag             =   "5"
         Top             =   900
         Width           =   255
      End
      Begin VB.OptionButton optPiece 
         Caption         =   "Option1"
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   4
         Tag             =   "4"
         Top             =   900
         Width           =   255
      End
      Begin VB.OptionButton optPiece 
         Caption         =   "Option1"
         Height          =   285
         Index           =   1
         Left            =   1410
         TabIndex        =   3
         Tag             =   "3"
         Top             =   900
         Width           =   255
      End
      Begin VB.OptionButton optPiece 
         Caption         =   "Option1"
         Height          =   285
         Index           =   0
         Left            =   870
         TabIndex        =   2
         Tag             =   "2"
         Top             =   900
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   600
         Left            =   690
         Picture         =   "frmPromote.frx":0000
         ScaleHeight     =   540
         ScaleWidth      =   2115
         TabIndex        =   1
         Top             =   300
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmPromote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PromoteTo As enumChessPiece

Private Sub Form_Load()
    optPiece(0).Value = False
    optPiece(1).Value = False
    optPiece(2).Value = False
    optPiece(3).Value = False
End Sub

Private Sub optPiece_Click(Index As Integer)
    PromoteTo = optPiece(Index).Tag
    Me.Hide
End Sub
