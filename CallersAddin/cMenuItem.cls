VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit ' Menu Item Class - adapted from Darryl Hasieber :)

Private WithEvents MenuHandler As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1
Private oCmdBarCtl As Office.CommandBarControl

Public Sub Add(oPopup As Office.CommandBar, sCaption As String, ByVal Id As Long, Optional Bitmap As Object)
  On Error GoTo ErrHandler
    Set oCmdBarCtl = oPopup.Controls.Add(Type:=msoControlButton, Id:=Id) ' Create sub-item in the menu
    oCmdBarCtl.Caption = sCaption       ' Assign the specified caption
    Set MenuHandler = oVBE.Events.CommandBarEvents(oCmdBarCtl)
    If Not Bitmap Is Nothing Then
        Dim oTemp As Object, sClipText As String
      On Error GoTo ErrWith
        With Clipboard
            sClipText = .GetText
            Set oTemp = .GetData
            .SetData Bitmap, vbCFBitmap ' Copy the icon to the clipboard
            oCmdBarCtl.PasteFace        ' Set the icon for the button
            .Clear
            If Not oTemp Is Nothing Then .SetData oTemp
            .SetText sClipText
ErrWith:
        End With
    End If
ErrHandler:
  If Err Then LogError "cMenuItem.Add", sCaption
End Sub

Public Sub Remove()
    If Not oCmdBarCtl Is Nothing Then
       oCmdBarCtl.Delete
       Set oCmdBarCtl = Nothing
       Set MenuHandler = Nothing
    End If
End Sub

Private Sub Class_Terminate()
    Call Remove
End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)
    With CommandBarControl
        CodePaneMenuItem_Click .Caption, .Id
    End With
End Sub

