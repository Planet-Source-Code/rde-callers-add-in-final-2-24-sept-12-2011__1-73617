Attribute VB_Name = "modErrors"
Option Explicit                                    ' ©Rd

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long ' ©Rd
Private Declare Function GetModuleHandleZ Lib "kernel32" Alias "GetModuleHandleA" (ByVal hNull As Long) As Long

'Randy Birch, VBnet.com
Private Declare Function StrLenW Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

Private Const GWL_HINSTANCE = &HFFFFFFFA
Private Const MAX_PATH As Long = 260

Private mFile As String
Private mInit As Boolean

Public Sub InitErr(Optional sCompName As String, Optional ByVal fClearMsgLog As Boolean)
  On Error GoTo Fail
    Dim i As Integer
    If Not mInit Then
        mFile = RTrimChr(GetParentPath, "\") & "\" & sCompName
        If fClearMsgLog Then
            i = FreeFile()
            Open mFile & "_Msg.log" For Output As #i
            Close #i
        End If
        mInit = True
    End If
Fail:
End Sub

Public Sub LogError(sProcName As String, Optional sExtraInfo As String)
    Dim Num As Long, Src As String, Desc As String
    With Err
      Num = .Number: Src = .Source: Desc = .Description
    End With
  On Error GoTo Fail
    If Erl Then Desc = Desc & vbNewLine & "Error on line " & Erl
    If LenB(sExtraInfo) Then Desc = Desc & vbNewLine & sExtraInfo
    If mInit Then Else InitErr
    Dim i As Integer: i = FreeFile()
    Open mFile & "_Error.log" For Append As #i
        Print #i, Src; " error log ";
        Print #i, Format$(Now, "h:nn:ss am/pm mmmm d, yyyy")
        Print #i, sProcName; " error!"
        Print #i, "Error #"; Num; " - "; Desc
        Print #i, " * * * * * * * * * * * * * * * * * * *"
Fail:
    Close #i
    Beep
End Sub

Public Sub LogMsg(Msg As String)
  On Error GoTo Fail
    If mInit Then Else InitErr
    Dim i As Integer: i = FreeFile()
    Open mFile & "_Msg.log" For Append As #i
        Print #i, Format$(Now, "h:nn:ss am/pm mmmm d, yyyy")
        Print #i, Msg
        Print #i, " * * * * * * * * * * * * * * * * * * *"
    Close #i
Fail:
End Sub

Private Function GetParentPath() As String
  On Error GoTo Fail
    Dim sModName As String, hInst As Long, rc As Long
    ' Get the application hInstance. By passing NULL, GetModuleHandle
    ' returns a handle to the file used to create the calling process.
    hInst = GetWindowLong(GetModuleHandleZ(0&), GWL_HINSTANCE)
    ' Get the module file name
    sModName = String$(MAX_PATH, vbNullChar)
    rc = GetModuleFileName(hInst, sModName, MAX_PATH)
    GetParentPath = TrimZ(sModName)
Fail:
    ' Return empty string on error
End Function

Public Function TrimZ(StrZ As String) As String
    ' StrZ = "strZstrZstrZstrZZ[ZZZZZZ]" >> TrimZ = "str"
    ' StrZ = "strZ[ZZZZZZ]"              >> TrimZ = "str"
    ' StrZ = "str  "                     >> TrimZ = "str  "
    Dim lLen As Long
    lLen = StrLenW(StrPtr(StrZ)) 'Randy Birch
    TrimZ = LeftB$(StrZ, lLen + lLen) 'Rd
End Function

' ========================================================================
' This vb5 function removes from sStr the first occurrence from the right
' of the specified character(s) and everything following it, and returns just
' the start of the string up to but not including the specified character(s).
' It always searches from right to left starting at the end of sStr. If the
' character(s) does not exist in sStr then the whole of sStr is returned and
' lRetPos is set to Len(sStr) + 1. sChar defaults to a backslash if omitted.
' ========================================================================
Public Function RTrimChr(sStr As String, Optional sChar As String = "\", Optional ByRef lRetPos As Long, _
                                Optional ByVal eCompare As VbCompareMethod = vbBinaryCompare) As String
    Dim lPos As Long
    ' Default to return the passed string
    lRetPos = Len(sStr) + 1&
    If LenB(sChar) Then
        lPos = InStr(1&, sStr, sChar, eCompare)
        Do Until lPos = 0&
            lRetPos = lPos
            lPos = InStr(lRetPos + 1&, sStr, sChar, eCompare)
        Loop
    End If
    ' Return sStr w/o sChar and any following substring
    RTrimChr = LeftB$(sStr, lRetPos + lRetPos - 2&)
End Function

'     ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯    :›)
