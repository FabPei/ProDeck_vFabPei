Attribute VB_Name = "modSplitMergeObjectByParagraph"
Option Explicit
'20251214

Sub SmartSplitOrMerge()
    ' Functionality: Intelligent dispatcher that decides whether to Split or Merge.
    ' Key Behavior: If 1 shape is selected -> Calls SplitTextByParagraph.
    '               If >1 shapes are selected -> Calls MergeSelectedShapes.
    '               If 0 shapes selected -> Prompts user.

    ' Check what is selected
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select one or more shapes.", vbExclamation
        Exit Sub
    End If
    
    Dim count As Long
    count = ActiveWindow.Selection.ShapeRange.count
    
    If count = 1 Then
        ' Single object selected: Run Split
        SplitTextByParagraph
    ElseIf count > 1 Then
        ' Multiple objects selected: Run Merge
        MergeSelectedShapes
    Else
        MsgBox "Please select a shape.", vbExclamation
    End If

End Sub

Sub SplitTextByParagraph()
    ' Functionality: Checks if one object is selected and splits text by paragraph.
    ' Key Behavior: Creates duplicates of the original shape for every paragraph found.
    '               Preserves formatting of the original shape.
    '               Stacks the new shapes vertically below the original position.
    
    Dim shSource As Shape
    Dim shNew As Shape
    Dim para As TextRange
    Dim currentTop As Single
    Dim paraText As String
    Dim i As Long

    ' Guard clause for direct execution
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Sub
    If ActiveWindow.Selection.ShapeRange.count <> 1 Then
        MsgBox "Select exactly one object to split.", vbExclamation
        Exit Sub
    End If

    Set shSource = ActiveWindow.Selection.ShapeRange(1)

    ' Check for text frame presence
    If Not shSource.HasTextFrame Then Exit Sub
    If shSource.TextFrame.TextRange.text = "" Then Exit Sub

    currentTop = shSource.Top
    
    ' Loop through paragraphs
    For i = 1 To shSource.TextFrame.TextRange.Paragraphs.count
        Set para = shSource.TextFrame.TextRange.Paragraphs(i)
        paraText = para.text
        
        ' Clean trailing carriage returns
        If Right(paraText, 1) = vbCr Then
            paraText = Left(paraText, Len(paraText) - 1)
        End If
        
        ' Create shape if text is not empty
        If Len(Trim(paraText)) > 0 Then
            ' FIX: Duplicate returns a ShapeRange, so we access Item(1) to get the Shape object
            Set shNew = shSource.Duplicate.Item(1)
            
            shNew.TextFrame.TextRange.text = paraText
            shNew.Left = shSource.Left
            shNew.Top = currentTop
            shNew.TextFrame.AutoSize = ppAutoSizeShapeToFitText
            
            ' Increment position for next shape
            currentTop = shNew.Top + shNew.Height + 10
            shNew.Select msoFalse
        End If
    Next i

    ' Remove the original bulk shape
    shSource.Delete
End Sub

Sub MergeSelectedShapes()
    ' Functionality: Merges text from multiple selected shapes into the top-most shape.
    ' Key Behavior: Sorts selected shapes by vertical position (.Top) to ensure logical text flow.
    '               Combines text with line breaks.
    '               Deletes the lower shapes after merging their text into the top shape.
    
    Dim shR As ShapeRange
    Dim shTarget As Shape
    Dim i As Long, j As Long
    Dim tempSh As Shape
    Dim sortedShapes() As Shape
    
    ' Guard clause for direct execution
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Sub
    If ActiveWindow.Selection.ShapeRange.count < 2 Then
        MsgBox "Select at least two objects to merge.", vbExclamation
        Exit Sub
    End If
    
    Set shR = ActiveWindow.Selection.ShapeRange
    
    ' 1. Sort Shapes by .Top position (Bubble Sort)
    ReDim sortedShapes(1 To shR.count)
    For i = 1 To shR.count
        Set sortedShapes(i) = shR(i)
    Next i
    
    For i = 1 To UBound(sortedShapes) - 1
        For j = i + 1 To UBound(sortedShapes)
            If sortedShapes(j).Top < sortedShapes(i).Top Then
                Set tempSh = sortedShapes(i)
                Set sortedShapes(i) = sortedShapes(j)
                Set sortedShapes(j) = tempSh
            End If
        Next j
    Next i
    
    ' 2. Define the Target (The top-most shape)
    Set shTarget = sortedShapes(1)
    
    ' 3. Loop through the rest and append text
    For i = 2 To UBound(sortedShapes)
        If sortedShapes(i).HasTextFrame Then
            If sortedShapes(i).TextFrame.HasText Then
                ' Append a new line and the text from the lower shape
                shTarget.TextFrame.TextRange.InsertAfter vbCr & sortedShapes(i).TextFrame.TextRange.text
            End If
        End If
        ' Delete the shape after merging
        sortedShapes(i).Delete
    Next i
    
    ' Select the final merged shape
    shTarget.Select
End Sub
