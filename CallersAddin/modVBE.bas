Attribute VB_Name = "modVBE"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lLenB As Long)

' Application reference - do not store in the Connect designer
Public oVBE As VBIDE.VBE
Public oPopupMenu As Office.CommandBar

Private cMenuItem() As cMenuItem
Private saCallers() As String
Private iaCallers() As Long
Private fExempt As Long

Public nCallers As Long

Public Sub RedimCallers(ByVal NumCallers As Long)
    ReDim Preserve saCallers(0 To NumCallers) As String
    ReDim Preserve iaCallers(0 To NumCallers) As Long
    ReDim Preserve cMenuItem(1 To NumCallers) As cMenuItem
End Sub

Public Sub EraseCallerArrays()
    Erase saCallers()
    Erase iaCallers()
    Do While nCallers
        cMenuItem(nCallers).Remove
        Set cMenuItem(nCallers) = Nothing
        nCallers = nCallers - 1
    Loop
    Erase cMenuItem()
End Sub

Public Sub CodePaneMenuItem_Click(sMenuCaption As String, ByVal Idx As Long)
    Dim sCompName As String
    Dim lStartLine As Long
    Dim lTopLine As Long
    Dim i As Long
  On Error GoTo ErrHandler
    i = InStr(1, sMenuCaption, ".")
    sCompName = Left$(sMenuCaption, i - 1)
  On Error GoTo ErrWith
    With oVBE.ActiveVBProject.VBComponents(sCompName).CodeModule
       lStartLine = iaCallers(Idx)
       lTopLine = lStartLine - CLng(.CodePane.CountOfVisibleLines / 2.5)
       If lTopLine < 1 Then lTopLine = 1
       .CodePane.TopLine = lTopLine
       Call .CodePane.SetSelection(lStartLine, 1, lStartLine, 1)
       If Not oVBE.ActiveCodePane Is .CodePane Then
          .Parent.Activate ' Activate the component
       End If
      .CodePane.Show
ErrWith:
    End With
ErrHandler:
 If Err Then LogError "modVBE.CodePaneMenuItem_Click", sMenuCaption
End Sub

Public Sub DisplayCallee()
    CodePaneMenuItem_Click saCallers(0), 0
End Sub

Public Sub ResetContextMenu()
   Do While nCallers
      cMenuItem(nCallers).Remove
      Set cMenuItem(nCallers) = Nothing
      nCallers = nCallers - 1
   Loop
   If Not oPopupMenu Is Nothing Then
      Call oPopupMenu.Delete
      Set oPopupMenu = Nothing
   End If
End Sub

Public Sub RefreshProjectReferences()
          ' Adapted from Add Error Handling addin by Kamilche
          Dim eProcKind As vbext_ProcKind
          Dim oComp As VBComponent
          Dim oCodePane As CodePane
          Dim oThisMod As CodeModule
          Dim oNextMod As CodeModule

          Dim i As Long, j As Long, k As Long
          Dim lCurrentLine As Long
          Dim lCurrentCol As Long
          Dim vbeScope As vbext_Scope
          Dim sProcName As String
          Dim sCompName As String
          Dim sProcRef As String

130      On Error GoTo ErrHandler
140        Set oCodePane = oVBE.ActiveCodePane
          ' Exit if we're not in the code pane
150        If oCodePane Is Nothing Then Exit Sub
160        Set oThisMod = oCodePane.CodeModule

          ' Retrieve the current line the cursor is on
170        oCodePane.GetSelection lCurrentLine, lCurrentCol, j, k

          ' Retrieve the procedure name and the ProcKind
180        sProcName = oThisMod.ProcOfLine(lCurrentLine, eProcKind)

           fExempt = 0
           If LenB(sProcName) Then
185            vbeScope = oThisMod.Members(sProcName).Scope
           Else 'LenB(sProcName) = 0
190            sProcName = GetDeclareName(oThisMod, lCurrentLine, lCurrentCol, vbeScope)
           End If

200        If LenB(sProcName) Then

210           sCompName = oThisMod.Parent.Name
220           sProcRef = sCompName & "." & sProcName

230           If sProcRef = saCallers(0) Then Exit Sub
240           saCallers(0) = sProcRef
250           iaCallers(0) = lCurrentLine
260           Call ResetContextMenu

             ' Search current code module for references to current procedure
270           FindCallers oThisMod, sProcName, vbNullString, lCurrentLine

280           If vbeScope <> vbext_Private Then ' vbext_Friend | vbext_Public

                 ' Search the other components for references to current procedure
290               For Each oComp In oVBE.ActiveVBProject.VBComponents
300                  If Not oComp Is Nothing Then
310                   If Not oComp.Name = vbNullString Then
320                    If Not oComp Is oThisMod.Parent Then

                         On Error Resume Next
330                       Set oNextMod = Nothing  ' Bug fix Dec 18, 2010
                           Select Case oComp.Type
                              Case vbext_ct_RelatedDocument, vbext_ct_ResFile
                               ' Related docs, res files throw exception
                              Case Else
340                            Set oNextMod = oComp.CodeModule
                              'Case vbext_ct_DocObject
                              'Case vbext_ct_ClassModule
                              'Case vbext_ct_MSForm
                              'Case vbext_ct_PropPage
                              'Case vbext_ct_StdModule
                              'Case vbext_ct_UserControl
                              'Case vbext_ct_VBForm
                              'Case vbext_ct_VBMDIForm
                              'Case vbext_ct_ActiveXDesigner
                           End Select
                          On Error GoTo ErrHandler

350                        If Not oNextMod Is Nothing Then
360                            FindCallers oNextMod, sProcName, sCompName
370                        End If
380                     End If
                      End If
                    End If
390               Next oComp
400           End If

410           If nCallers Then
                 ' (Re-)create the popup menu using the Position argument
420               Set oPopupMenu = oVBE.CommandBars.Add(Name:="Callers", Position:=msoBarPopup)

430               For i = 1 To nCallers
440                   Set cMenuItem(i) = New cMenuItem ' Create new item in popup-menu
450                   cMenuItem(i).Add oPopupMenu, saCallers(i), i, LoadResPicture(102, vbResBitmap)
460               Next
470           End If

480       Else 'LenB(sProcName) = 0
490           saCallers(0) = vbNullString
500           Call ResetContextMenu
510       End If

ErrHandler:
      If Err Then LogError "modVBE.RefreshProjectReferences", sProcRef
End Sub

Private Function GetInstanceName(oCodeMod As CodeModule, sClassName As String, ByVal lStartLine As Long, ByVal lEndLine As Long) As String
          Dim sCodeLine As String
          Dim lLineEnd As Long
          Dim lStartCol As Long
          Dim lEndColumn As Long
          Dim i As Long, j As Long
          Dim fTryAgain As Boolean
610     On Error GoTo EndWith
620       With oCodeMod
630        lStartCol = 1
           lEndColumn = -1
           lLineEnd = lEndLine
640        Do
            ' Search the procedure for class instantiation (also frm As New form, ctl As New ctrl, etc)
650          Do While .Find(sClassName, lStartLine, lStartCol, lEndLine, lEndColumn, True, True)
660              sCodeLine = .Lines(lStartLine, 1)
670              If IsCode(sCodeLine, lStartCol, lEndColumn) Then
680                  If lStartCol > 4 Then
690                      If Mid$(sCodeLine, lStartCol - 4, 4) = " As " Then
700                         j = 4   ' cClass  As 'sClassName'
                           'Do While IsDelim(Mid$(sCodeLine, lStartCol - j - 1, 1))
710                         Do While IsDelimI(MidI(sCodeLine, lStartCol - j - 1))
720                             j = j + 1
730                         Loop
740                         i = j + 1
750                         Do While Not IsDelimI(MidI(sCodeLine, lStartCol - i - 1))
760                             i = i + 1
770                         Loop
780                         GetInstanceName = Mid$(sCodeLine, lStartCol - i, i - j)
790                         GoTo EndWith
800                     End If
810                 End If
820                 If lStartCol > 8 Then
830                     If Mid$(sCodeLine, lStartCol - 8, 8) = " As New " Then
840                         j = 8   ' cClass  As New 'sClassName'
850                         Do While IsDelimI(MidI(sCodeLine, lStartCol - j - 1))
860                             j = j + 1
870                         Loop
880                         i = j + 1
890                         Do While Not IsDelimI(MidI(sCodeLine, lStartCol - i - 1))
900                             i = i + 1
910                         Loop
920                         GetInstanceName = Mid$(sCodeLine, lStartCol - i, i - j)
930                         GoTo EndWith
940                     End If
950                 End If
960                 If lStartCol > 7 Then
970                     If Mid$(sCodeLine, lStartCol - 7, 7) = " = New " Then
980                         i = InStr(sCodeLine, "Set ") + 4  'Set cClass = New 'sClassName'
990                         If i > 4 And i < lStartCol - 7 Then ' Bug fix Dec 12, 2010
1000                             GetInstanceName = Mid$(sCodeLine, i, lStartCol - 7 - i)
1010                             GoTo EndWith
1020                         End If
1030                     End If
1040                 End If
1050                 If lStartCol > 3 Then
1060                     If Mid$(sCodeLine, lStartCol - 3, 3) = " = " Then
1070                         i = InStr(sCodeLine, "Set ") + 4  'Set cClass = 'sClassName'
1080                         If i > 4 And i < lStartCol - 3 Then ' Bug fix Dec 12, 2010
1090                             GetInstanceName = Mid$(sCodeLine, i, lStartCol - 3 - i)
1100                             GoTo EndWith
1110                         End If
1120                     End If
1130                 End If
1140             End If
1150             lStartCol = lEndColumn + 1
1160             lEndColumn = -1
                 lEndLine = lLineEnd
1170         Loop

1180         fTryAgain = (lLineEnd > .CountOfDeclarationLines + 1)

            ' Search the component for class instantiation (also frm As New form, ctl As New ctrl, etc)
             If fTryAgain Then
               lStartLine = 1
               lLineEnd = .CountOfDeclarationLines + 1
               lEndLine = lLineEnd
               lStartCol = 1
               lEndColumn = -1
             End If

1190       Loop While fTryAgain

          ' Default to Class name
           GetInstanceName = sClassName
EndWith:
        End With
    If Err Then LogError "modVBE.GetInstanceName", sClassName
End Function

Private Sub FindCallers(oCodeMod As CodeModule, sProcName As String, sCompName As String, Optional ByVal lProcLine As Long)
          ' Adapted from Project References addin by ':) Ulli
          Dim eProcKind As vbext_ProcKind
          Dim sMembName As String
          Dim sInstName As String
          Dim sCodeLine As String
          Dim sCaller As String
          Dim lStartCol As Long
          Dim lEndColumn As Long
          Dim lLineStart As Long
          Dim lLineCount As Long
          Dim lLineLen As Long
          Dim lCodeLine As Long
          Dim lEndLine As Long
          Dim lContinue As Long
          Dim fIsCode As Long
          Dim fMustQualify As Long
          Dim i As Long, j As Long

1200     On Error GoTo ErrWith
1210       With oCodeMod

            ' First search in the declarations section
1220         lCodeLine = 1
1230         lEndLine = .CountOfDeclarationLines + 1

1240         lStartCol = 1
1250         lEndColumn = -1

1260         Do While .Find(sProcName, lCodeLine, lStartCol, lEndLine, lEndColumn, True, True)

1270            sCodeLine = .Lines(lCodeLine, 1)
1280            If IsWholeWord(sCodeLine, lStartCol, lEndColumn) Then

1290               fIsCode = IsCode(sCodeLine, lStartCol, lEndColumn)
1300               If fIsCode Then

1310                  If LenB(sCompName) Then
1320                     If oVBE.ActiveVBProject.VBComponents(sCompName).Type <> vbext_ct_StdModule Then
1330                       If Not fExempt Then fMustQualify = -1
                         End If
                      End If

1340                  If IsValid(oCodeMod, lCodeLine, lStartCol, sCompName, lCodeLine, fMustQualify, sProcName, lProcLine) Then

1350                     lContinue = lCodeLine
1360                     Do While lCodeLine > 1 ' .ProcOfLine kinda thing
                           'If Right$(.Lines(lCodeLine - 1, 1), 1) = "_" Then
1370                        If RightI(.Lines(lCodeLine - 1, 1), 1) = 95 Then
1380                           lCodeLine = lCodeLine - 1 ' Line continuation
1390                           sCodeLine = .Lines(lCodeLine, 1)
1400                           lLineLen = Len(sCodeLine)                       ' Check for beginning a comment
1410                           fIsCode = IsCode(sCodeLine, lLineLen, lLineLen) ' before the end of the line
                               If Not fIsCode Then Exit Do
                            Else
                               Exit Do
                            End If
                         Loop
1420                     If fIsCode Then
1430                        If lCodeLine <> lProcLine Then ' lProcLine is zero if not current comp
      
1440                           sMembName = GetDeclareName(oCodeMod, lCodeLine, 1)
1450                           If LenB(sMembName) And (sMembName <> sProcName) Then

1460                              sCaller = oCodeMod.Parent.Name & "." & sMembName

1470                              If saCallers(nCallers) <> sCaller Then
1480                                 nCallers = nCallers + 1
1490                                 If nCallers > UBound(saCallers) Then
1500                                     RedimCallers UBound(saCallers) + 100
1510                                 End If
1520                                 saCallers(nCallers) = sCaller
1530                                 iaCallers(nCallers) = lCodeLine
1540                              End If
1550                           End If
                            End If
                         End If
1560                     lCodeLine = lContinue
                      End If
                   End If
                End If
1570            lStartCol = lEndColumn + 1
1580            lEndColumn = -1
1590            lEndLine = .CountOfDeclarationLines + 1
             Loop

            ' Locate the first line of procedure code
1600         lCodeLine = .CountOfDeclarationLines + 1
1610         lEndLine = -1

1620         lStartCol = 1
1630         lEndColumn = -1

1640         Do While .Find(sProcName, lCodeLine, lStartCol, lEndLine, lEndColumn, True, True)
1650            If lCodeLine <> lProcLine Then
1660               sCodeLine = .Lines(lCodeLine, 1)

1670               If IsCode(sCodeLine, lStartCol, lEndColumn) Then

                      ' Grab member name (and procedure kind) of procedure
1680                   sMembName = .ProcOfLine(lCodeLine, eProcKind)

1690                    If sMembName <> sProcName Then

1700                       lLineStart = .ProcBodyLine(sMembName, eProcKind)
1710                       If lLineStart <> lCodeLine Then
1720                          If Not IsWholeWord(sCodeLine, lStartCol, lEndColumn) Then GoTo DoNext
1730                       Else
1740                          i = InStr(sMembName, sProcName)
                             ' If procedure name is not within member name (naming conflict with
1750                          If i = 0 Or i > 1 Then GoTo DoNext '  param name) or "..._ProcName"
                           End If

                           fMustQualify = 0
1760                       If LenB(sCompName) Then
1770                          lLineCount = .ProcCountLines(sMembName, eProcKind)
1780                          sInstName = GetInstanceName(oCodeMod, sCompName, lLineStart, lLineStart + lLineCount)
1790                          If oVBE.ActiveVBProject.VBComponents(sCompName).Type <> vbext_ct_StdModule Then
1800                             If Not fExempt Then fMustQualify = -1
                              End If
                           End If

1810                       If IsValid(oCodeMod, lCodeLine, lStartCol, sInstName, lLineStart, fMustQualify, sProcName, lProcLine) Then

1820                          sCaller = oCodeMod.Parent.Name & "." & sMembName

1830                          If saCallers(nCallers) <> sCaller Then
1840                              nCallers = nCallers + 1
1850                              If nCallers > UBound(saCallers) Then
1860                                  RedimCallers UBound(saCallers) + 100
1870                              End If
1880                              saCallers(nCallers) = sCaller
1890                              iaCallers(nCallers) = lCodeLine
1900                          End If
1910                       End If
1920                    End If
1930                 End If
                  End If
DoNext:
1940              lStartCol = lEndColumn + 1
1950              lEndColumn = -1
1960              lEndLine = -1
1970         Loop
ErrWith:
1980       End With
       If Err Then LogError "modVBE.FindCallers", sCaller
End Sub

Private Function GetDeclareName(oCodeMod As CodeModule, ByRef lCodeLine As Long, ByVal lStartCol As Long, Optional vbeScope As vbext_Scope) As String
         Dim i As Long, j As Long, k As Long
         Dim sCodeLine As String
         Dim sTempLine As String
         Dim sMembName As String
         Dim sBuffer As String

2000     On Error GoTo ErrHandler

2010      sCodeLine = oCodeMod.Lines(lCodeLine, 1)
2020      If LenB(sCodeLine) = 0 Then Exit Function

2030      If IsCode(sCodeLine, lStartCol, lStartCol) Then

            ' Try to match the member name with the selection
2040         i = lStartCol
2050         j = lStartCol
2060         k = Len(sCodeLine)

2070         Do While i > 1  ' Step back to a delimiter
2080            If IsDelimI(MidI(sCodeLine, i - 1)) Then Exit Do
2090            i = i - 1
2100        Loop
2110        Do Until j > k ' Step forward to a delimiter
2120           If IsDelimI(MidI(sCodeLine, j)) Then Exit Do
2130           j = j + 1
2140        Loop

2150        sMembName = Trim$(Mid$(sCodeLine, i, j - i))
2160     End If

2170     sCodeLine = LTrim$(sCodeLine)
2180     If AscW(sCodeLine) = 35 Then Exit Function '# Line
2190     If AscW(sCodeLine) = 39 Then Exit Function 'Comment Line

2200     If LenB(sMembName) Then
2210       On Error Resume Next ' Is it a member name
2220        sBuffer = oCodeMod.Members(sMembName).Name

2230        If LenB(sBuffer) Then ' If so, is it a code member
2240          vbeScope = oCodeMod.Members(sBuffer).Scope

2250          If vbeScope <> 0 Then ' If so, we have it
2260            GetDeclareName = sBuffer
2270            Exit Function
2280          End If
2290        End If
2300       On Error GoTo ErrHandler
2310     End If

2320     If lCodeLine <= oCodeMod.CountOfDeclarationLines Then

          ' Check for line continuation
           Do While lCodeLine > 1 ' .ProcOfLine kinda thing
             'If Right$(oCodeMod.Lines(lCodeLine - 1, 1), 1) = "_" Then
              If RightI(oCodeMod.Lines(lCodeLine - 1, 1), 1) = 95 Then
                 lCodeLine = lCodeLine - 1 ' Line continuation
              Else
                 Exit Do
              End If
           Loop

2330       sCodeLine = LTrim$(oCodeMod.Lines(lCodeLine, 1))
           sTempLine = " " & sCodeLine

          ' Enums and Types are not included in the
          ' members collection so try them first
2340       If InStr(sTempLine, " Enum ") Then
2350          j = InStr(sTempLine, " Enum ") + 6
2360          k = InStr(j, sTempLine, " ")
2370          If k = 0 Then
2380             GetDeclareName = Mid$(sTempLine, j)
2390          Else
2400             GetDeclareName = Mid$(sTempLine, j, k - j)
2410          End If
2420          j = InStr(sCodeLine, " ")
2430          Select Case Left$(sCodeLine, j - 1)
                Case "Public"
2440              vbeScope = vbext_Public
2460              fExempt = -1
               'Case "Private"
                 'vbeScope = vbext_Private
2480            Case Else
2490              vbeScope = vbext_Private
2500          End Select
2510          Exit Function

2520       ElseIf InStr(sTempLine, " Type ") Then
2530          j = InStr(sTempLine, " Type ") + 6
2540          k = InStr(j, sTempLine, " ")
2550          If k = 0 Then
2560             GetDeclareName = Mid$(sTempLine, j)
2570          Else
2580             GetDeclareName = Mid$(sTempLine, j, k - j)
2590          End If
2600          j = InStr(sCodeLine, " ")
2610          Select Case Left$(sCodeLine, j - 1)
                Case "Private"
2630              vbeScope = vbext_Private
               'Case "Public"
                 'vbeScope = vbext_Public
2650            Case Else
2660              vbeScope = vbext_Public
2670          End Select
2680          Exit Function

          ' An Implements object is also overlooked
2690       ElseIf Left$(sTempLine, 12) = " Implements " Then
2700          j = 13
2710          k = InStr(j, sTempLine, " ")
2720          If k = 0 Then
2725             GetDeclareName = Mid$(sTempLine, j)
2730          Else
2735             GetDeclareName = Mid$(sTempLine, j, k - j)
2740          End If
2745          vbeScope = vbext_Private
2750          Exit Function

          ' Also try a raised Event
2755       ElseIf Left$(sTempLine, 7) = " Event " Then
2760          j = 8
2765          k = InStr(j, sTempLine, "(")
2770          If Not (k = 0) Then
2775             GetDeclareName = Mid$(sTempLine, j, k - j)
2780             vbeScope = vbext_Private
2785             Exit Function
2790          End If

2795       End If

2800       j = InStr(sCodeLine, " ") ' Member of a Type or Enum?
2810       If j = 0 Then j = Len(sCodeLine) + 1
2820       sMembName = Left$(sCodeLine, j - 1)

2830       sBuffer = " Public Private Declare Const Dim Global "
2840       If InStr(sBuffer, " " & sMembName & " ") = 0 Then

2850          i = -1
2855          Do While (lCodeLine > 1) And i ' .ProcOfLine kinda thing
2860             lCodeLine = lCodeLine - 1
2865             sTempLine = " " & LTrim$(oCodeMod.Lines(lCodeLine, 1))

2870             j = InStr(2, sTempLine, " ")
2875             If j = 0 Then j = Len(sTempLine) + 1
2880             sMembName = LTrim$(Left$(sTempLine, j - 1))

2885             If InStr(sBuffer, " " & sMembName & " ") <> 0 Then i = 0

2890             If InStr(sTempLine, " Type ") Then
2900                 j = InStr(sTempLine, " Type ") + 6
2910                 k = InStr(j, sTempLine, " ")
2920                 If k = 0 Then
2930                    GetDeclareName = Mid$(sTempLine, j)
2940                 Else
2950                    GetDeclareName = Mid$(sTempLine, j, k - j)
2960                 End If
2970                 sCodeLine = LTrim$(sTempLine)
2980                 j = InStr(sCodeLine, " ")
2990                 Select Case Left$(sCodeLine, j - 1)
                        Case "Public"
3010                       vbeScope = vbext_Public
                       'Case "Private"
                         'vbeScope = vbext_Private
3030                    Case Else
3040                      vbeScope = vbext_Private
3050                End Select
3060                Exit Function

3070             ElseIf InStr(sTempLine, " Enum ") Then
3080                j = InStr(sTempLine, " Enum ") + 6
3090                k = InStr(j, sTempLine, " ")
3100                If k = 0 Then
3110                   GetDeclareName = Mid$(sTempLine, j)
3120                Else
3130                   GetDeclareName = Mid$(sTempLine, j, k - j)
3140                End If
3150                sCodeLine = LTrim$(sTempLine)
3160                j = InStr(sCodeLine, " ")
3170                Select Case Left$(sCodeLine, j - 1)
                       Case "Public"
3180                      vbeScope = vbext_Public
3200                      fExempt = -1
                      'Case "Private"
                         'vbeScope = vbext_Private
3220                   Case Else
3230                      vbeScope = vbext_Private
3240                End Select
3250                Exit Function
3260             End If

3270         Loop
3280      End If 'Not a declaration keyword?

3290    End If 'If lCodeLine <= oCodeMod.CountOfDeclarationLines

       ' User did not click on top of a member so loop through
       ' the members to try to find one within this code line
3300    For i = 1 To oCodeMod.Members.Count

3310      j = InStr(sCodeLine, " " & oCodeMod.Members(i).Name)
3320      If j > 0 Then

3330         sMembName = oCodeMod.Members(i).Name
3340         k = j + 1 + Len(sMembName)
3350         Select Case oCodeMod.Members(i).Type

                Case vbext_mt_Variable
3360               If MidI(sCodeLine, k) = 32 Or MidI(sCodeLine, k) = 40 Then ' " " Or "("
3370                  If InStr(sCodeLine, " ") = j Then
3380                     GetDeclareName = sMembName
3390                     vbeScope = oCodeMod.Members(i).Scope
3400                     Exit For
3410                  ElseIf InStr(sCodeLine, " WithEvents ") = j - 11 Then
3420                     GetDeclareName = sMembName
3430                     vbeScope = oCodeMod.Members(i).Scope
3440                     Exit For
3450                  End If
3460               ElseIf k > Len(sCodeLine) Then
3470                  GetDeclareName = sMembName
3480                  vbeScope = oCodeMod.Members(i).Scope
3490                  Exit For
3500               End If

3510            Case vbext_mt_Const
3520               If InStr(sCodeLine, "Const") = j - 5 Then
3530                  If MidI(sCodeLine, k) = 32 Then ' " "
3540                     GetDeclareName = sMembName
3550                     vbeScope = oCodeMod.Members(i).Scope
3560                     Exit For
3570                  End If
3580               End If

3590            Case vbext_mt_Event  ' Raised Events
3600               If InStr(sCodeLine, "Event") = j - 5 Then
3610                  GetDeclareName = sMembName
3620                  vbeScope = oCodeMod.Members(i).Scope
3630                  Exit For
3640               End If

3650            Case vbext_mt_Method      ' Parses ALL procedures except properties,
3660               If Mid$(sCodeLine, k, 5) = " Lib " Then ' including API Declares
3670                  GetDeclareName = sMembName
3680                  vbeScope = oCodeMod.Members(i).Scope
3690                  Exit For
3700               End If

               'Case vbext_mt_Property ' Handled by ProcOfLine in RefreshProjectReferences
3710         End Select
           
3720      End If
3730    Next

ErrHandler:
   If Err Then LogError "modVBE.GetDeclareName", sMembName
End Function

Private Function IsWholeWord(sLine As String, ByVal lStartCol As Long, ByVal lEndColumn As Long) As Boolean
  On Error GoTo ErrHandler
    Dim fDelim As Long ' The VBE's CodeModule Find function accepts an underscore as a delimiter
    If (lStartCol > 1) Then
        fDelim = MidI(sLine, lStartCol - 1) <> 95 '<> "_"
    Else
        fDelim = -1
    End If
    If fDelim Then
       If lEndColumn <= Len(sLine) Then
          fDelim = MidI(sLine, lEndColumn) <> 95 '<> "_"
       Else
          fDelim = -1
       End If
    End If
    IsWholeWord = fDelim
ErrHandler:
  If Err Then LogError "modVBE.IsWholeWord", sLine
End Function

Private Function IsCode(sLine As String, ByVal lStartCol As Long, ByVal lEndColumn As Long) As Boolean
    ' Adapted from Project References addin by ':) Ulli
    Const Comment As String = "'"
    Const Quote As String = """"
    Dim i As Long, j As Long, k As Long
   On Error GoTo ErrHandler
   ' See if word is in a comment
    If Left$(LTrim$(sLine), 4) = "Rem " Then Exit Function
    If Left$(LTrim$(sLine), 1) = Comment Then Exit Function

    i = InStr(1, sLine, Comment)
    k = InStr(1, sLine, Quote)

    Do Until i = 0 Or k = 0

       j = InStr(k + 1, sLine, Quote)
       If j = 0 Then Exit Function ' Error out!

       If k < i And j > i Then ' If comment is in a string literal
            i = InStr(i + 1, sLine, Comment)
       End If
       k = InStr(j + 1, sLine, Quote)
    Loop

    If i = 0 Then i = Len(sLine)
    If lStartCol > i Then Exit Function

    k = InStr(1, sLine, Quote)
    Do Until k = 0 Or k > i

       j = InStr(k + 1, sLine, Quote)
       If j = 0 Then Exit Function ' Error out!

       If k < lStartCol And j >= lEndColumn Then ' Bug fix Dec 19, 2010
             Exit Function ' The word is in a string literal
       End If
       k = InStr(j + 1, sLine, Quote)
    Loop
    IsCode = True

ErrHandler:
 If Err Then LogError "modVBE.IsCode", sLine
End Function

Private Function IsValid(oCodeMod As CodeModule, ByVal lCodeLine As Long, ByVal lStartCol As Long, sCompName As String, ByVal lLineStart As Long, ByVal fMustQualify As Long, sMembName As String, ByVal lProcLine As Long) As Boolean
    Dim sCodeLine As String
    Dim fWithBlock As Long
    Dim i As Long, j As Long
    
   On Error GoTo ErrHandler

     If lStartCol > 1 Then
         sCodeLine = oCodeMod.Lines(lCodeLine, 1)
         If MidI(sCodeLine, lStartCol - 1) = 46 Then ' "."
             If lStartCol > 2 Then
                'If Not IsDelim(Mid$(sCodeLine, lStartCol - 2, 1)) Then
                 If Not IsDelimI(MidI(sCodeLine, lStartCol - 2)) Then
                     If lStartCol > 3 Then
                        If Mid$(sCodeLine, lStartCol - 3, 2) = "Me" Then
                           If lStartCol > 4 Then ' Bug fix Dec 11, 2010
                              If IsDelimI(MidI(sCodeLine, lStartCol - 4)) Then
                                 IsValid = Not fMustQualify
                              End If
                           Else ' If Me is current class, fMustQualify is false
                              IsValid = Not fMustQualify
                           End If
                        End If
                     End If ' Qualified, but is it our component?
                     i = Len(sCompName)
                     If lStartCol > i + 1 Then ' "comp.proc"
                        IsValid = (Mid$(sCodeLine, lStartCol - i - 1, i) = sCompName)
                     End If
                 Else ' We have a With block " .proc" | "(.proc" etc
                     fWithBlock = -1
                 End If
             Else ' We have a With block ".proc"
                 fWithBlock = -1
             End If
         Else ' " proc" | "(proc" etc
            IsValid = Not fMustQualify
            If IsValid And (lProcLine = 0) Then ' lProcLine is zero if not current comp
              ' Check for duplicate name with narrower scope
               For i = 1 To oCodeMod.Members.Count
                  If oCodeMod.Members(i).Name = sMembName Then
                     IsValid = False
                     Exit For
                  End If
               Next i
            End If
         End If

         If fWithBlock Then
            Do While (lCodeLine > lLineStart) ' .ProcOfLine kinda thing
               lCodeLine = lCodeLine - 1
               sCodeLine = " " & oCodeMod.Lines(lCodeLine, 1)
               i = InStr(sCodeLine, " With ")
               If i Then
                   i = i + 6
                   j = InStr(i, sCodeLine, " ")
                   If j = 0 Then j = Len(sCodeLine) + 1

                   IsValid = (Mid$(sCodeLine, i, j - i) = sCompName)
                   Exit Do
               End If
            Loop
         End If
     Else ' "proc"
         IsValid = Not fMustQualify
         If IsValid And (lProcLine = 0) Then
           ' Check for duplicate name with narrower scope
            For i = 1 To oCodeMod.Members.Count
               If oCodeMod.Members(i).Name = sMembName Then
                  IsValid = False
                  Exit For
               End If
            Next i
         End If
     End If

ErrHandler:
 If Err Then LogError "modVBE.IsValid", sCodeLine
End Function

' ¤¤ IsDelim ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  This function checks if the character passed is a common word
'  delimiter, and then returns True or False accordingly.
'
'  By default, any non-alphabetic character (except for an
'  underscore) is considered a word delimiter, including numbers.
'
'  By default, an underscore is treated as part of a whole word,
'  and so is not considered a word delimiter.
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Public Function IsDelimI(ByVal iAscW As Long) As Boolean ' ©Rd
    Select Case iAscW
        ' Uppercase, Underscore, Lowercase chars not delimiters
        Case 65 To 90, 95, 97 To 122: IsDelimI = False

        'Case 39, 146: IsDelimI = False  ' Apostrophes not delimiters
        Case 48 To 57: IsDelimI = False ' Numeric chars not delimiters

        Case Else: IsDelimI = True ' Any other character is delimiter
    End Select
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Visual Basic is one of the few languages where you
' can't extract a character from or insert a character
' into a string at a given position without creating
' another string.
'
' The following Property fixes that limitation.
'
' Twice as fast as AscW and Mid$ when compiled.
'        iChr = AscW(Mid$(sStr, lPos, 1))
'        iChr = MidI(sStr, lPos)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get MidI(sStr As String, ByVal lPos As Long) As Integer
    CopyMemory MidI, ByVal StrPtr(sStr) + lPos + lPos - 2&, 2&
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Mid$(sStr, lPos, 1) = Chr$(iChr)
'        MidI(sStr, lPos) = iChr
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let MidI(sStr As String, ByVal lPos As Long, ByVal iChrW As Integer)
    CopyMemory ByVal StrPtr(sStr) + lPos + lPos - 2&, iChrW, 2&
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       iChr = AscW(Right$(sStr, lPos, 1))
'       iChr = RightI(sStr, lPos)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get RightI(sStr As String, ByVal lRightPos As Long) As Integer
    CopyMemory RightI, ByVal StrPtr(sStr) + LenB(sStr) - lRightPos - lRightPos, 2&
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Right$(sStr, lPos, 1) = Chr$(iChr)
'      RightI(sStr, lPos) = iChr
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let RightI(sStr As String, ByVal lRightPos As Long, ByVal iChrW As Integer)
    CopyMemory ByVal StrPtr(sStr) + LenB(sStr) - lRightPos - lRightPos, iChrW, 2&
End Property
