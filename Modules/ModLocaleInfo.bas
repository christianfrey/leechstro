Attribute VB_Name = "ModLocaleInfo"
Private Declare Function GetThreadLocale Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" _
Alias "GetLocaleInfoA" (ByVal Locale As Long, _
                        ByVal LCType As Long, _
                        ByVal lpLCData As String, _
                        ByVal cchData As Long) As Long


Private Const LOCALE_IDEFAULTANSICODEPAGE = &H1004&
Private Const LOCALE_SENGCOUNTRY = &H1002
Private Const LOCALE_SNATIVECTRYNAME = &H8
Private Const LOCALE_SCOUNTRY = &H6

Public Function GetCharSet() As Integer
    '
    Dim lngLCID         As Long
    Dim strLcid         As String
    Dim strCodePage     As String
    Dim lngRetVal       As Long
    '
    strCodePage = String$(16, " ")
    '
    'Get Current locale
    lngLCID = GetThreadLocale()
    'Convert to Hex
    strLcid = Hex$(Trim$(CStr(lngLCID)))
    '
    'Get code page
    lngRetVal = GetLocaleInfo(LCID, LOCALE_IDEFAULTANSICODEPAGE, strCodePage, Len(strCodePage))
    strCodePage = Left$(strCodePage, InStr(1, strCodePage, Chr(0)) - 1)
    '
    'Get char set from code page
    Select Case strCodePage
        Case "932" ' Japanese
            GetCharSet = 128
        Case "936" ' Simplified Chinese
            GetCharSet = 134
        Case "949" ' Korean
            GetCharSet = 129
        Case "950" ' Traditional Chinese
            GetCharSet = 136
        Case "1250" ' Eastern Europe
            GetCharSet = 238
        Case "1251" ' Russian
            GetCharSet = 204
        Case "1252" ' Western European Languages
            GetCharSet = 0
        Case "1253" ' Greek
            GetCharSet = 161
        Case "1254" ' Turkish
            GetCharSet = 162
        Case "1255" ' Hebrew
            GetCharSet = 177
        Case "1256" ' Arabic
            GetCharSet = 178
        Case "1257" ' Baltic
            GetCharSet = 186
        Case Else
            GetCharSet = 0
    End Select
    '
End Function

