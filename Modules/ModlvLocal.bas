Attribute VB_Name = "modlvLocal"
' CurPath, "\", True
Public Function GetRightWord(ByRef Sentense As String, ByRef StopChar As String, RemoveIt As Boolean)
    
    Dim pos As Integer
    pos = InStrRev(Sentense, StopChar, -1, vbTextCompare)
    
    If pos = 0 Then
        GetRightWord = Sentense
    Else
        GetRightWord = Right(Sentense, Len(Sentense) - pos)
    End If

    If RemoveIt = True Then Sentense = Left$(Sentense, pos)

End Function

Public Function GetLeftWord(ByRef Sentense As String, ByRef StopChar As String, RemoveIt As Boolean)
    
    Dim pos As Integer
    pos = InStr(1, Sentense, StopChar, vbTextCompare)
    
    If pos = 0 Then
        GetLeftWord = Sentense
    Else
        GetLeftWord = Left$(Sentense, pos - 1)
    End If

    If RemoveIt = True Then Sentense = Right(Sentense, Len(Sentense) - pos)

End Function
