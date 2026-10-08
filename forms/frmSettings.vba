' =====================================================================
'  Code for the frmSettings UserForm.
'
'  Paste this into the FORM'S code window (right-click the form in the
'  Project pane and choose View Code), not into the main module.
'
'  The form needs two CommandButtons on it, named cmdOK and cmdCancel.
'  Position and captions don't matter - the macro sets those. Every other
'  control is built at run time by BuildSettingsControls in the main module.
' =====================================================================
Option Explicit

' True when the user clicked Export, False when they cancelled or closed
' the window. Read by ShowSettingsDialog in the main module.
Public Confirmed As Boolean

Private Sub UserForm_Initialize()
    Confirmed = False
    BuildSettingsControls Me
End Sub

Private Sub cmdOK_Click()
    Confirmed = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Confirmed = False
    Me.Hide
End Sub

' Clicking the X counts as cancelling. Hide rather than unload, so the
' macro can still read Confirmed afterwards.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Confirmed = False
        Cancel = True
        Me.Hide
    End If
End Sub
