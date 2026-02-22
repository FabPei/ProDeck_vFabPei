VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufCreateGridOfShapes 
   Caption         =   "Create Grid of Shapes"
   ClientHeight    =   2848
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   4088
   OleObjectBlob   =   "ufCreateGridOfShapes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufCreateGridOfShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Code for the UserForm: ufCreateGridOfShapes

' --- Form Control Events ---

Private Sub UserForm_Initialize()
    Me.txtColumns.Value = "2"
    Me.txtRows.Value = "2"
    Me.txtSpacing.Value = "5"
    Me.chkHorizontalLines.Value = False
    Me.chkVerticalLines.Value = False
    Me.Caption = "Split Shape"
End Sub

Private Sub btnOK_Click()
    ' --- Input Validation ---
    If Not IsNumeric(Me.txtColumns.Value) Or CInt(Me.txtColumns.Value) < 1 Then
        MsgBox "Please enter a valid number of columns (minimum 1).", vbExclamation
        Me.txtColumns.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Me.txtRows.Value) Or CInt(Me.txtRows.Value) < 1 Then
        MsgBox "Please enter a valid number of rows (minimum 1).", vbExclamation
        Me.txtRows.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(Me.txtSpacing.Value) Or CSng(Me.txtSpacing.Value) < 0 Then
        MsgBox "Please enter a valid spacing value (minimum 0).", vbExclamation
        Me.txtSpacing.SetFocus
        Exit Sub
    End If
    
    ' --- Run the split logic directly ---
    ExecuteSplit
    
    ' --- Hide the form after execution ---
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.Hide
    End If
End Sub


' --- Core Logic Procedure ---

Private Sub ExecuteSplit()
    ' Retrieve values directly from form controls
    Dim numColumns As Integer, numRows As Integer, spacing As Single
    Dim addHorzLines As Boolean, addVertLines As Boolean
    
    numColumns = CInt(Me.txtColumns.Value)
    numRows = CInt(Me.txtRows.Value)
    spacing = CSng(Me.txtSpacing.Value)
    addHorzLines = Me.chkHorizontalLines.Value
    addVertLines = Me.chkVerticalLines.Value

    ' --- Shape Splitting & Line Drawing ---
    Dim originalShape As Shape
    Set originalShape = ActiveWindow.Selection.ShapeRange(1)
    
    Dim sld As Slide
    Set sld = ActiveWindow.View.Slide

    Dim origLeft As Single, origTop As Single
    Dim origWidth As Single, origHeight As Single
    origLeft = originalShape.Left
    origTop = originalShape.Top
    origWidth = originalShape.Width
    origHeight = originalShape.Height

    Dim newWidth As Single, newHeight As Single
    newWidth = (origWidth - (spacing * (numColumns - 1))) / numColumns
    newHeight = (origHeight - (spacing * (numRows - 1))) / numRows

    If newWidth <= 0 Or newHeight <= 0 Then
        MsgBox "The spacing is too large for the number of rows/columns.", vbExclamation
        Exit Sub
    End If

    ' This line was removed: Application.ScreenUpdating = False

    ' Create shape grid
    Dim r As Integer, c As Integer
    For r = 1 To numRows
        For c = 1 To numColumns
            Dim currentLeft As Single, currentTop As Single
            currentLeft = origLeft + ((c - 1) * (newWidth + spacing))
            currentTop = origTop + ((r - 1) * (newHeight + spacing))
            
            If r = 1 And c = 1 Then
                With originalShape
                    .Width = newWidth
                    .Height = newHeight
                    .Left = currentLeft
                    .Top = currentTop
                End With
            Else
                With originalShape.Duplicate
                    .Left = currentLeft
                    .Top = currentTop
                End With
            End If
        Next c
    Next r

    ' --- Add Horizontal Lines ---
    If addHorzLines And numRows > 1 Then
        For r = 1 To numRows - 1
            Dim lineTop As Single
            lineTop = origTop + (r * newHeight) + (r * spacing) - (spacing / 2)
            sld.Shapes.AddLine origLeft, lineTop, origLeft + origWidth, lineTop
        Next r
    End If

    ' --- Add Vertical Lines ---
    If addVertLines And numColumns > 1 Then
        For c = 1 To numColumns - 1
            Dim lineLeft As Single
            lineLeft = origLeft + (c * newWidth) + (c * spacing) - (spacing / 2)
            sld.Shapes.AddLine lineLeft, origTop, lineLeft, origTop + origHeight
        Next c
    End If

    ' This line was removed: Application.ScreenUpdating = True
    
    MsgBox "Shape successfully split.", vbInformation
End Sub

