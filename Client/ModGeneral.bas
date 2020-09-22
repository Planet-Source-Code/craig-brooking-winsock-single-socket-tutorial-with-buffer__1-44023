Attribute VB_Name = "ModGeneral"
Option Explicit

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

Sub SendData(SndData)

FrmMain.WinSock.SendData ((Len(SndData)) & "~" & SndData) 'Send length of string, then ~ then string

End Sub
