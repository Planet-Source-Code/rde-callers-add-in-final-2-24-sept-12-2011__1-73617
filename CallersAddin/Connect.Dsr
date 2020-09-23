VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   11010
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   14205
   _ExtentX        =   25056
   _ExtentY        =   19420
   _Version        =   393216
   Description     =   "Callers AddIn"
   DisplayName     =   "Callers AddIn"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSafe     =   -1  'True
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private WithEvents ProjectsEvents As VBProjectsEvents
'Private WithEvents ComponentEvents As VBComponentsEvents

Private oCodePaneMenuCallers As Office.CommandBarControl
Private oCodePaneMenuCallee As Office.CommandBarControl
Private oImmediateWindMenu As Office.CommandBarControl

Private WithEvents CallersHandler As CommandBarEvents
Attribute CallersHandler.VB_VarHelpID = -1
Private WithEvents CalleeHandler As CommandBarEvents
Attribute CalleeHandler.VB_VarHelpID = -1
Private WithEvents ImmediateHandler As CommandBarEvents
Attribute ImmediateHandler.VB_VarHelpID = -1

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
 On Error GoTo ErrHandler
   Call InitErr("CallersAddin")
   Set oVBE = Application
   If ConnectMode = ext_cm_Startup Then
        'Set ProjectsEvents = oVBE.Events.VBProjectsEvents()
       If AddItemToMenu("Ca&llers...", "Code Window", oCodePaneMenuCallers, LoadResPicture(102, vbResBitmap)) Then
          Set CallersHandler = oVBE.Events.CommandBarEvents(oCodePaneMenuCallers)
       End If
       If AddItemToMenu("Call&ee", "Code Window", oCodePaneMenuCallee, LoadResPicture(101, vbResBitmap)) Then
          Set CalleeHandler = oVBE.Events.CommandBarEvents(oCodePaneMenuCallee)
       End If
       If AddItemToMenu("C&lear", "Immediate Window", oImmediateWindMenu) Then
          Set ImmediateHandler = oVBE.Events.CommandBarEvents(oImmediateWindMenu)
       End If
   End If
   Call RedimCallers(100)
ErrHandler:
 If Err Then LogError "Connect.AddinInstance_OnConnection"
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    Call ResetContextMenu
    Call EraseCallerArrays
    If Not oCodePaneMenuCallers Is Nothing Then
       Call oCodePaneMenuCallers.Delete
       Set oCodePaneMenuCallers = Nothing
       Set CallersHandler = Nothing
    End If
    If Not oCodePaneMenuCallee Is Nothing Then
       Call oCodePaneMenuCallee.Delete
       Set oCodePaneMenuCallee = Nothing
       Set CalleeHandler = Nothing
    End If
    If Not oImmediateWindMenu Is Nothing Then
       Call oImmediateWindMenu.Delete
       Set oImmediateWindMenu = Nothing
       Set ImmediateHandler = Nothing
    End If
    'Set ComponentEvents = Nothing
    'Set ProjectsEvents = Nothing
    Set oVBE = Nothing
End Sub

Private Function AddItemToMenu(sCaption As String, sMenuName As String, cbMenuCBar As Office.CommandBarControl, Optional oBitmap As Object) As Boolean
    Dim cbMenu As CommandBar
    Dim oTemp As Object, sClipText As String
  On Error GoTo AddItemError
     Set cbMenu = oVBE.CommandBars(sMenuName)
     If cbMenu Is Nothing Then Exit Function
                                    ' Create new item in the menu
     Set cbMenuCBar = cbMenu.Controls.Add(Type:=msoControlButton)
     cbMenuCBar.Caption = sCaption  ' Assign the specified caption
     AddItemToMenu = True           ' Return success to caller

     If Not oBitmap Is Nothing Then
       On Error GoTo ErrWith
         With Clipboard
            sClipText = .GetText
            Set oTemp = .GetData
            .SetData oBitmap, vbCFBitmap ' Copy the icon to the clipboard
            cbMenuCBar.PasteFace        ' Set the icon for the button
            .Clear
            If Not oTemp Is Nothing Then
               .SetData oTemp
               Set oTemp = Nothing
            End If
            .SetText sClipText
ErrWith:
        End With
    End If
AddItemError:
  If Err Then LogError "Connect.AddItemToMenu", sMenuName
End Function

Private Sub CallersHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
    Call RefreshProjectReferences
    If nCallers Then oPopupMenu.ShowPopup
End Sub

Private Sub CalleeHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
    If nCallers Then DisplayCallee
End Sub

Private Sub ImmediateHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
 On Error GoTo ErrHandler
    oVBE.Windows("Immediate").SetFocus
    SendKeys "^{Home}", True
    SendKeys "^+{End}", True
    SendKeys "{Del}", True
ErrHandler:
End Sub

'Private Sub ProjectsEvents_ItemActivated(ByVal VBProject As VBIDE.VBProject)
'    Set ComponentEvents = oVBE.Events.VBComponentsEvents(VBProject)
'End Sub
'
'Private Sub ProjectsEvents_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
'    If VBProject Is oVBE.ActiveVBProject Then
'        Set ComponentEvents = oVBE.Events.VBComponentsEvents(VBProject)
'    End If
'End Sub
'
'Private Sub ComponentEvents_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
'    RefreshProjectReferences
'End Sub
'
'Private Sub ComponentEvents_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
'    RefreshProjectReferences
'End Sub
'
'Private Sub ComponentEvents_ItemReloaded(ByVal VBComponent As VBIDE.VBComponent)
'    RefreshProjectReferences
'End Sub
'
'Private Sub ComponentEvents_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
'    RefreshProjectReferences
'End Sub
'
'Private Sub ComponentEvents_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
'    RefreshProjectReferences
'End Sub
