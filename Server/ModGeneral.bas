Attribute VB_Name = "ModGeneral"
Option Explicit

' ***** General Variables *****
Public NumUsers As Integer ' Holds how many people are connected

' ***** Public Constants *****
Public Const MAX_USERS As Integer = 999

' ***** Public Types *****
Public Type FreeSock
Used As Byte
UsedBy As String
End Type

Public Type Users
Name As String
Index As Integer
Connected As Byte
End Type

' ***** Public Type Arrays *****
Public FreeSocket(1 To MAX_USERS) As FreeSock
Public UserList(1 To MAX_USERS) As Users

Sub Connect(Index As Integer, RequestID As Long)


' ***** Add to UserList *****
UserList(Index).Index = Index
UserList(Index).Connected = 1
UserList(Index).Name = "Client" & UserList(Index).Index

' Load Winsock
Load FrmMain.WinSock(Index)

' Incrememnt users by 1
NumUsers = NumUsers + 1

' Accept User
FrmMain.WinSock(Index).Accept RequestID

' Let the person know there connected so as there TxtSend is unlocked
Call SendData("TOINDEX", Index, "CON")

'Add to Console
FrmMain.TxtConsole.Text = FrmMain.TxtConsole.Text & vbNewLine & UserList(Index).Name & " Has connected!"


End Sub

Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String
  
Seperator = Chr(SepASCII) ' Character to look for
LastPos = 0 ' reset LastPos
FieldNum = 0 ' Reset FieldNum

For i = 1 To Len(Text) ' Loop through string
    CurChar = Mid(Text, i, 1) ' Find ASCII of Character
    If CurChar = Seperator Then ' if Character = Seperator then...
        FieldNum = FieldNum + 1 ' Field = +1
        If FieldNum = Pos Then ' If FIeldNum is same as position then
            ReadField = Mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos)) ' Call Sub Again
            Exit Function
        End If
        LastPos = i ' Last position = value of loop
    End If
        
Next i

FieldNum = FieldNum + 1 ' Field Number = +1
If FieldNum = Pos Then
    ReadField = Mid(Text, LastPos + 1)
End If

End Function

Sub SendData(ByVal sndRoute As String, ByVal Index As Integer, ByVal StrData As String)

'*****************************************************************
'Sends data to sendRoute
'*****************************************************************
Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer
Dim Clan As String

'send NONE
If UCase(sndRoute) = "TONONE" Then ' If SndRoute = TONONE do nothing
    Exit Sub
End If
  
  
'Send to All
If UCase(sndRoute) = "TOALL" Then ' If SndRoute = TOALL
    For LoopC = 1 To NumUsers ' Loop through numusers
        If UserList(LoopC).Connected = 1 Then ' If user is connected
            FrmMain.WinSock(LoopC).SendData ((Len(StrData) + 1) & "~" & StrData) ' Send length of string, then string
        End If
    Next LoopC
    Exit Sub
End If

'Send to everyone but the sndindex
If UCase(sndRoute) = "TOALLBUTINDEX" Then
    For LoopC = 1 To NumUsers
              
      If UserList(LoopC).Connected = 1 And LoopC <> Index Then ' If user = connected and  index then
            FrmMain.WinSock(LoopC).SendData ((Len(StrData) + 1) & "~" & StrData) ' send length of string, then string
      End If
      
    Next LoopC
    Exit Sub
End If


'Send to the UserIndex
If UCase(sndRoute) = "TOINDEX" Then
    FrmMain.WinSock(Index).SendData ((Len(StrData) + 1) & "~" & StrData) ' send length of string, then string
    Exit Sub
End If

End Sub
