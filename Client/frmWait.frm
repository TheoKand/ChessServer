VERSION 5.00
Begin VB.Form frmWait 
   ClientHeight    =   2190
   ClientLeft      =   1695
   ClientTop       =   3630
   ClientWidth     =   4500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWait.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   4500
   Begin VB.Frame Frame1 
      Caption         =   "Παρακαλώ περιμένετε..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4245
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Ακύρωση"
         Height          =   465
         Left            =   2640
         TabIndex        =   2
         Top             =   1410
         Width           =   1485
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   150
         Picture         =   "frmWait.frx":0442
         Top             =   330
         Width           =   480
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Height          =   960
         Left            =   810
         TabIndex        =   1
         Top             =   330
         Width           =   3270
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    gConnectionState = csDisconnected
End Sub

