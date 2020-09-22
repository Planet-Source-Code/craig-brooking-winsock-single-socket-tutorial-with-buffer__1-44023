VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winsock Example (Single Socket) Server"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock WinSock 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox TxtConsole 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

' Add date to Console
TxtConsole.Text = "Server Started at " & Time & " On " & Date

' Number of Users
NumUsers = 1

' Port to listen On
WinSock(0).LocalPort = 27478

' Start Sniffing for connections
WinSock(0).Listen

End Sub

Private Sub WinSock_Close(Index As Integer)

' Close Socket
WinSock(Index).Close

' Unload Winsock Control
Unload WinSock(Index)

' -1 user online
NumUsers = NumUsers - 1

' Reset Stats
UserList(Index).Index = 0
UserList(Index).Connected = 0

' Add to console
TxtConsole.Text = TxtConsole.Text & vbNewLine & UserList(Index).Name & " Has Dis-Connected!"

' Reset Name
UserList(Index).Name = ""

'Make socket "free"
FreeSocket(Index).Used = 0
FreeSocket(Index).UsedBy = ""

End Sub

Private Sub WinSock_ConnectionRequest(Index As Integer, ByVal RequestID As Long)

Dim LoopC As Integer

For LoopC = 1 To MAX_USERS ' Loop through allm possible users
If FreeSocket(LoopC).Used = 0 Then ' if not used
FreeSocket(LoopC).Used = 1 ' used now
FreeSocket(LoopC).UsedBy = LoopC ' used by Index
Call Connect(LoopC, RequestID) ' Call connect (index)
Exit Sub ' exit after Connect sub been and gone
End If ' if not, nothing (No Space)
Next ' loop

End Sub

Private Sub WinSock_DataArrival(Index As Integer, ByVal bytesTotal As Long)

Dim RData As String

WinSock(Index).GetData RData ' Assign data to variable RData (String)

Call Buffer(RData, Index) ' Call Buffer (Splits up merged stuff)

End Sub
