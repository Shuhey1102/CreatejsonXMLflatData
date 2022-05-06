VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sBuf As String
Private iBuf As Long

Private Sub Class_Initialize()
    sBuf = String$(32768, vbNullChar)
    iBuf = 0
End Sub

Public Sub Append(ByRef sValue As String)
    If iBuf + Len(sValue) > Len(sBuf) Then
        sBuf = sBuf & String$(CLng(Len(sBuf) * 2) + Len(sValue), vbNullChar)
    End If
    Mid(sBuf, iBuf + 1) = sValue
    iBuf = iBuf + Len(sValue)
End Sub

Public Function ToString() As String
    ToString = Left$(sBuf, iBuf)
End Function

