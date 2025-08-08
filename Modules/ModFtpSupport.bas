Attribute VB_Name = "ModFtpSupport"
Public p_intCounter As Integer
Public p_strTimeOutedIDEvents As String

' UTILE POUR LE DIMENTIONNEMENT :
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

 
Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    p_strTimeOutedIDEvents = p_strTimeOutedIDEvents & CStr(idEvent) & ";"
End Sub

'Redimentionne les cellules pour qu'elles soient de taille équivalente à la taille du texte.
Public Sub lvwAutofitColumnWidth(ByVal lvw As ListView)
    Dim iCounter As Long
    On Error Resume Next
    
    If lvw.View <> lvwReport Then Exit Sub
    
    lvw.Visible = False
    
    For iCounter = 0 To lvw.ColumnHeaders.Count - 1
       Call SendMessage(lvw.hwnd, LVM_SETCOLUMNWIDTH, iCounter, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next
    
    lvw.Visible = True
 End Sub
