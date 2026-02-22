Attribute VB_Name = "modIdentifyNodesOfArrow"
Option Explicit

Private Type PentPoint
    x As Double
    y As Double
    Name As String
    DistanceToTip As Double
    id As Integer
End Type

Sub GetPentagonCoordinates_StrictLogic()
    Dim originalShp As Shape, tempShp As Shape
    Dim sld As slide
    Dim i As Integer, j As Integer
    Dim nodePoints As Variant
    Dim x As Double, y As Double
    Dim nodeCount As Integer
    
    ' Arrays and Logic
    Dim uniquePoints() As PentPoint
    Dim uniqueCount As Integer
    Dim isDuplicate As Boolean
    Dim tipIndex As Integer
    Dim rot As Double
    
    ' 1. Validation: Check Shape
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select a Pentagon shape first.", vbExclamation
        Exit Sub
    End If
    
    Set originalShp = ActiveWindow.Selection.ShapeRange(1)
    Set sld = ActiveWindow.View.slide
    
    ' 2. Validation: Check Rotation (Strict 0, 90, 180, 270)
    rot = originalShp.Rotation
    ' Normalize rotation to 0-360 range just in case
    Do While rot < 0: rot = rot + 360: Loop
    Do While rot >= 360: rot = rot - 360: Loop
    
    If Not (rot = 0 Or rot = 90 Or rot = 180 Or rot = 270) Then
        MsgBox "This script only supports pentagons rotated at exactly 0, 90, 180, or 270 degrees.", vbExclamation
        Exit Sub
    End If

    ' 3. Create Temp Freeform & FIX OFFSET
    Set tempShp = originalShp.Duplicate(1)
    
    ' --- OFFSET CORRECTION ---
    ' Duplicate() shifts the shape. We must force alignment.
    tempShp.Left = originalShp.Left
    tempShp.Top = originalShp.Top
    ' -------------------------
    
    tempShp.Select
    On Error Resume Next
    Application.CommandBars.ExecuteMso ("ShapeConvertToFreeform")
    If Err.Number <> 0 Then
        MsgBox "Could not convert shape.", vbCritical
        tempShp.Delete: Exit Sub
    End If
    On Error GoTo 0
    Set tempShp = ActiveWindow.Selection.ShapeRange(1)
    
    ' 4. Extract and Deduplicate Nodes
    nodeCount = tempShp.Nodes.Count
    ReDim uniquePoints(1 To nodeCount)
    uniqueCount = 0
    
    For i = 1 To nodeCount
        nodePoints = tempShp.Nodes(i).Points
        x = nodePoints(1, 1)
        y = nodePoints(1, 2)
        
        ' Deduplication Loop
        isDuplicate = False
        For j = 1 To uniqueCount
            ' Threshold of 1.0 to catch floating point variance
            If Abs(uniquePoints(j).x - x) < 1 And Abs(uniquePoints(j).y - y) < 1 Then
                isDuplicate = True
                Exit For
            End If
        Next j
        
        If Not isDuplicate Then
            uniqueCount = uniqueCount + 1
            uniquePoints(uniqueCount).x = x
            uniquePoints(uniqueCount).y = y
            uniquePoints(uniqueCount).id = uniqueCount
        End If
    Next i
    
    ' Validation: Ensure we actually have a Pentagon (5 points)
    If uniqueCount <> 5 Then
        MsgBox "Error: Found " & uniqueCount & " corners. This logic requires exactly 5 unique corners.", vbCritical
        tempShp.Delete: Exit Sub
    End If

    ' 5. LOGIC: Find the TIP
    ' Rule: "The tip point never shares the same x or y coordinate of the other points"
    tipIndex = 0
    Dim sharesX As Boolean, sharesY As Boolean
    
    For i = 1 To uniqueCount
        sharesX = False
        sharesY = False
        
        For j = 1 To uniqueCount
            If i <> j Then
                ' Check X overlap
                If Abs(uniquePoints(i).x - uniquePoints(j).x) < 1 Then sharesX = True
                ' Check Y overlap
                If Abs(uniquePoints(i).y - uniquePoints(j).y) < 1 Then sharesY = True
            End If
        Next j
        
        ' If it shares neither X nor Y, it is the Tip
        If sharesX = False And sharesY = False Then
            tipIndex = i
            uniquePoints(i).Name = "Tip"
            Exit For
        End If
    Next i
    
    If tipIndex = 0 Then
        MsgBox "Could not identify a unique Tip. Ensure the shape is a standard Pentagon.", vbCritical
        tempShp.Delete: Exit Sub
    End If
    
    ' 6. LOGIC: Identify Shoulders vs Base
    ' Rule: Shoulders are closer to the Tip than the Base points.
    
    Dim dist As Double
    Dim tipX As Double, tipY As Double
    tipX = uniquePoints(tipIndex).x
    tipY = uniquePoints(tipIndex).y
    
    ' Calculate Euclidean distance to Tip for all points
    For i = 1 To uniqueCount
        If i <> tipIndex Then
            ' a^2 + b^2 = c^2
            dist = Sqr((uniquePoints(i).x - tipX) ^ 2 + (uniquePoints(i).y - tipY) ^ 2)
            uniquePoints(i).DistanceToTip = dist
        Else
            uniquePoints(i).DistanceToTip = 0
        End If
    Next i
    
    ' Sort non-tip points by distance (Bubble sort for simplicity with 5 items)
    ' We only sort an index array to keep the original array intact
    Dim sortedIndices(1 To 4) As Integer
    Dim idxPtr As Integer
    idxPtr = 1
    For i = 1 To uniqueCount
        If i <> tipIndex Then
            sortedIndices(idxPtr) = i
            idxPtr = idxPtr + 1
        End If
    Next i
    
    Dim tempIdx As Integer
    For i = 1 To 3
        For j = i + 1 To 4
            If uniquePoints(sortedIndices(i)).DistanceToTip > uniquePoints(sortedIndices(j)).DistanceToTip Then
                tempIdx = sortedIndices(i)
                sortedIndices(i) = sortedIndices(j)
                sortedIndices(j) = tempIdx
            End If
        Next j
    Next i
    
    ' Now:
    ' sortedIndices(1) and (2) are SHOULDERS (Closest)
    ' sortedIndices(3) and (4) are BASE (Furthest)
    
    ' 7. LOGIC: Assign Left/Right or Top/Bottom names
    ' We compare the coordinates of the pair relative to the Tip or each other
    
    Dim s1 As Integer, s2 As Integer ' Shoulders
    Dim b1 As Integer, b2 As Integer ' Base
    s1 = sortedIndices(1): s2 = sortedIndices(2)
    b1 = sortedIndices(3): b2 = sortedIndices(4)
    
    ' Helper to name a pair
    NamePair uniquePoints, s1, s2, "Shoulder", rot
    NamePair uniquePoints, b1, b2, "Base", rot

    ' 8. Output
    Debug.Print "--- Geometry Logic for " & originalShp.Name & " (Rot: " & rot & ") ---"
    For i = 1 To uniqueCount
        Debug.Print Format(uniquePoints(i).Name, "@@@@@@@@@@@@@@@") & _
                    " | X: " & Format(uniquePoints(i).x, "0.00") & _
                    " | Y: " & Format(uniquePoints(i).y, "0.00")
        
        DrawMarker sld, uniquePoints(i).x, uniquePoints(i).y, uniquePoints(i).Name
    Next i
    
    ' Cleanup
    tempShp.Delete
    originalShp.Select
    MsgBox "Coordinates extracted. Check Immediate Window.", vbInformation

End Sub

Sub NamePair(ByRef pts() As PentPoint, idx1 As Integer, idx2 As Integer, baseName As String, rot As Double)
    ' This sub decides if a pair is "Left/Right" or "Top/Bottom"
    ' based on rotation and coordinate comparison.
    
    Dim p1 As PentPoint, p2 As PentPoint
    p1 = pts(idx1)
    p2 = pts(idx2)
    
    If rot = 0 Or rot = 180 Then
        ' Vertical Orientation: Shoulders/Base are distinguished by Left/Right (X)
        If p1.x < p2.x Then
            pts(idx1).Name = baseName & " Left"
            pts(idx2).Name = baseName & " Right"
        Else
            pts(idx1).Name = baseName & " Right"
            pts(idx2).Name = baseName & " Left"
        End If
    Else
        ' Horizontal Orientation (90/270): Shoulders/Base are distinguished by Top/Bottom (Y)
        ' Remember: Lower Y value = Top visually
        If p1.y < p2.y Then
            pts(idx1).Name = baseName & " Top"
            pts(idx2).Name = baseName & " Bottom"
        Else
            pts(idx1).Name = baseName & " Bottom"
            pts(idx2).Name = baseName & " Top"
        End If
    End If
End Sub

Sub DrawMarker(sld As slide, x As Double, y As Double, lbl As String)
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeOval, x - 3, y - 3, 6, 6)
    shp.Fill.ForeColor.RGB = RGB(255, 0, 0)
    shp.Line.Visible = msoFalse
    
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, x + 5, y - 10, 150, 20)
    shp.TextFrame.TextRange.Text = lbl
    shp.TextFrame.TextRange.Font.Size = 8
    shp.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
End Sub

