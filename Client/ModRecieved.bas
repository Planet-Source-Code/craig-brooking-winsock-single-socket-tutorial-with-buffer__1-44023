Attribute VB_Name = "ModRecieved"

Option Explicit

Sub Buffer(TempData As String)

Dim StrData As String
Dim LngLenString As Long

On Error Resume Next

LngLenString = Val(ReadField(1, TempData, 126)) ' Find out how long string is by reading first field

TempData = Right$(TempData, Len(TempData) - Len(Str(LngLenString))) ' Take the first field (Length of next string) off

StrData = Left$(TempData, LngLenString - 1) ' StrDatas length = LngLenString, take this off front of string

TempData = Right$(TempData, Len(TempData) - Len(StrData)) ' Cut string off text, ready for next one to be removed.
Call HandleData(StrData) ' Call handle data(StrData)

If Len(TempData) = 0 Then ' If theres nothing left, stop.
Exit Sub
End If

Call Buffer(TempData) ' If there is, call buffer again.

End Sub

Sub HandleData(Rdata As String)

' ***** MAKE SURE IT ISNT A CONNECT STRING *****
If Rdata = "CON" Then ' if Rdata = CON
FrmMain.TxtSend.Locked = False ' unload txtsend (Users can type now)
FrmMain.TxtRecieved.Text = "Connected..." ' Add to txtrecieved
Exit Sub ' (Exit sub)
End If ' If Not, Continue

' ***** INCOMING CHAT *****
If Left$(Rdata, 5) = "/CHAT" Then
Rdata = Right$(Rdata, Len(Rdata) - 5)

FrmMain.TxtRecieved = FrmMain.TxtRecieved & Rdata

Exit Sub
End If

End Sub
