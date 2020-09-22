Attribute VB_Name = "ModRecieved"
Option Explicit

Sub Buffer(TempData As String, Index As Integer)

Dim StrData As String
Dim LngLenString As Long

On Error Resume Next

LngLenString = Val(ReadField(1, TempData, 126)) ' Find out how long string is by reading first field

TempData = Right$(TempData, Len(TempData) - Len(Str(LngLenString))) ' Take the first field (Length of next string) off

StrData = Left$(TempData, LngLenString) ' StrDatas length = LngLenString, take this off front of string
TempData = Right$(TempData, Len(TempData) - Len(StrData)) ' Cut string off text, ready for next one to be removed.

Call HandleData(StrData, Index) ' Call handle data(StrData, index)


If Len(TempData) = 0 Then ' If theres nothing left, stop.
Exit Sub
End If

Call Buffer(TempData, Index) ' If there is, call buffer again.

End Sub

Sub HandleData(RData As String, Index As Integer)

' ***** INCOMING CHAT *****
If Left$(RData, 5) = "/CHAT" Then
RData = Right$(RData, Len(RData) - 5)

'Rdata = "/CHAT" & UserList(Index).Name & ": " & Rdata
Call SendData("ToAll", Index, "/CHAT" & UserList(Index).Name & ": " & RData)

' Add text to console
FrmMain.TxtConsole.Text = FrmMain.TxtConsole.Text & vbNewLine & UserList(Index).Name & ": " & RData
'Fix Console height
FrmMain.TxtConsole.SelLength = FrmMain.TxtConsole.Height

Exit Sub
End If

End Sub
