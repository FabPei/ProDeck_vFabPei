Attribute VB_Name = "modRandomizerLottery"
' [2025-12-14]
Option Explicit

Sub OpenRandomizerLottery()
    ' Functionality: Launches the Randomizer UserForm.
    ' Key Behavior: Opens the form immediately (Modeless), regardless of current selection.
    '               User can select shapes after the form is open.
    
    ' Create and show the UserForm (non-modal allowing interaction with PPT)
    Dim uf As New ufRandomizerLottery
    uf.Show vbModeless
End Sub
