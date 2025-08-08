Attribute VB_Name = "modGlobal"
Option Explicit

' Convert Long bytes to String KB
Public Function KBfromB(ByVal Bytesize As Long) As String
  Dim lKB As Long
  lKB = Bytesize / 1024
  If lKB < 1 Then lKB = 1
  KBfromB = lKB & "KB"
End Function

' Return count of selected files in specified list
Public Function CountSelectedFiles(poList As ListView) As Integer
  Dim iCnt As Integer
  Dim oLI As ListItem
  With poList
    For Each oLI In .ListItems
      If oLI.Selected Then iCnt = iCnt + 1
    Next
  End With
  CountSelectedFiles = iCnt
End Function

' Add backslash to end of path if it needs it
Public Function AddSlash(ByRef PathSpec As String, Optional SepChar As String = "\") As String
  If Right$(" " + PathSpec, 1) <> SepChar Then PathSpec = PathSpec + SepChar
  AddSlash = PathSpec
End Function

' Return a tab delimited list of filenames
Public Function GetSelectedItems(poList As ListView) As String
  Dim oLI As ListItem
  Dim svRet As String
  With poList
    For Each oLI In .ListItems
      If oLI.Selected Then svRet = svRet & oLI.Text & vbTab
    Next
  End With
  GetSelectedItems = svRet
End Function

' Remove eading slash if present
Public Function RemoveLeadSlash(ByRef Data, Optional SepChar As String = "/")
  If Left$(Data + " ", 1) = SepChar Then
    Data = Trim$(Mid$(Data + " ", 2))
  End If
  RemoveLeadSlash = Data
End Function

' Shows size in KB, MB, GB or TB depending on file size
Public Function ShowSize(Data As Double, Optional JustKB As Boolean = True) As String
  
  Dim ndData As Double
  ' ---------------------
  ndData = Data / 1024
  If JustKB Then
    If ndData = 0 Then
      ShowSize = " "
    ElseIf ndData > 0 And ndData < 1 Then
      ndData = "1"
    End If
    If ndData <> 0 Then ShowSize = Format$(CStr(ndData), "###,###") & "KB"
  Else
    If ndData = 0 Then
      ShowSize = " "
    Else
      If ndData < 1024 Then
        ShowSize = Format$(CStr(ndData), "###,###") & "KB"
      Else
        ndData = ndData / 1024
        If ndData < 1024 Then
          ShowSize = Format$(CStr(ndData), "###,###.0") & "MB"
        Else
          ndData = ndData / 1024
          If ndData < 1024 Then
            ShowSize = Format$(CStr(ndData), "###,###.00") & "GB"
          Else
            ndData = ndData / 1024
            If ndData < 1024 Then
              ShowSize = Format$(CStr(ndData), "###,###.000") & "TB"
            End If
          End If
        End If
      End If
    End If
  End If
End Function
