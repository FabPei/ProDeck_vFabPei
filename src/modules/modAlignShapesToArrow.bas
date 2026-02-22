Attribute VB_Name = "modAlignShapesToArrow"
Option Explicit

Private Type PentPoint
    x As Double
    y As Double
    DistanceToTip As Double
    id As Integer
End Type

Sub AlignShapeToArrow()
    Dim shpRef As Shape   ' The Arrow/Pentagon
    Dim shpMove As Shape  ' The shape to move
    Dim rot As Double
    
    ' Coordinates
    Dim shoulderX As Double, shoulderY As Double
    Dim tipX As Double, tipY As Double
    
    ' 1. Selection Validation
    If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
        MsgBox "Please select exactly two shapes:" & vbCrLf & _
               "1. The Arrow/Pentagon (Reference)" & vbCrLf & _
               "2. The Shape to move", vbExclamation
        Exit Sub
    End If
    
    Set shpRef = ActiveWindow.Selection.ShapeRange(1)
    Set shpMove = ActiveWindow.Selection.ShapeRange(2)
    
    ' 2. Check Rotation (Strict 90-degree increments)
    rot = shpRef.Rotation
    Do While rot < 0: rot = rot + 360: Loop
    Do While rot >= 360: rot = rot - 360: Loop
    
    If Not (rot = 0 Or rot = 90 Or rot = 180 Or rot = 270) Then
        MsgBox "The reference shape must be rotated at exactly 0, 90, 180, or 270 degrees.", vbExclamation
        Exit Sub
    End If
    
    ' 3. Extract Geometry (Tip and Shoulders)
    If Not GetPentagonGeometry(shpRef, tipX, tipY, shoulderX, shoulderY) Then
        Exit Sub ' Error message is inside the function
    End If
    
    ' 4. Determine Orientation and Align
    ' We compare Tip coordinates to Shoulder coordinates to know direction
    
    Dim epsilon As Double
    epsilon = 1 ' Tolerance for floating point comparison
    
    If Abs(tipX - shoulderX) < epsilon Then
        ' --- VERTICAL ORIENTATION (Tip and Shoulders share Center X) ---
        
        If tipY < shoulderY Then
            ' Case: Arrow Points UP (Tip Y is smaller/higher than Shoulders)
            ' Logic: Connect Target Top to Shoulder Line
            shpMove.Top = shoulderY
            ' Center horizontally? (Optional)
            shpMove.Left = shoulderX - (shpMove.Width / 2)
            
        Else
            ' Case: Arrow Points DOWN (Tip Y is larger/lower than Shoulders)
            ' Logic: Connect Target Bottom to Shoulder Line
            shpMove.Top = shoulderY - shpMove.Height
            ' Center horizontally? (Optional)
            shpMove.Left = shoulderX - (shpMove.Width / 2)
        End If
        
    ElseIf Abs(tipY - shoulderY) < epsilon Then
        ' --- HORIZONTAL ORIENTATION (Tip and Shoulders share Center Y) ---
        
        If tipX > shoulderX Then
            ' Case: Arrow Points RIGHT (Tip X is larger than Shoulders)
            ' Logic: Connect Target Right Edge to Shoulder Line
            shpMove.Left = shoulderX - shpMove.Width
            ' Center vertically? (Optional)
            shpMove.Top = shoulderY - (shpMove.Height / 2)
            
        Else
            ' Case: Arrow Points LEFT (Tip X is smaller than Shoulders)
            ' Logic: Connect Target Left Edge to Shoulder Line
            shpMove.Left = shoulderX
            ' Center vertically? (Optional)
            shpMove.Top = shoulderY - (shpMove.Height / 2)
        End If
        
    Else
        MsgBox "Geometry error: Could not determine clear orientation.", vbCritical
    End If

End Sub

' ==============================================================================
' HELPER: Returns Tip X/Y and Average Shoulder X/Y
' ==============================================================================
Function GetPentagonGeometry(originalShp As Shape, ByRef tX As Double, ByRef tY As Double, ByRef sX As Double, ByRef sY As Double) As Boolean
    Dim tempShp As Shape
    Dim i As Integer, j As Integer
    Dim nodePoints As Variant
    Dim x As Double, y As Double
    Dim nodeCount As Integer
    
    ' Data Structures
    Dim uniquePoints() As PentPoint
    Dim uniqueCount As Integer
    Dim isDuplicate As Boolean
    Dim tipIndex As Integer
    
    ' 1. Duplicate & Align
    Set tempShp = originalShp.Duplicate(1)
    tempShp.Left = originalShp.Left
    tempShp.Top = originalShp.Top
    
    ' 2. Convert to Freeform
    tempShp.Select
    On Error Resume Next
    Application.CommandBars.ExecuteMso ("ShapeConvertToFreeform")
    If Err.Number <> 0 Then
        MsgBox "Could not process shape geometry.", vbCritical
        tempShp.Delete
        GetPentagonGeometry = False
        Exit Function
    End If
    On Error GoTo 0
    Set tempShp = ActiveWindow.Selection.ShapeRange(1)
    
    ' 3. Extract Nodes
    nodeCount = tempShp.Nodes.Count
    ReDim uniquePoints(1 To nodeCount)
    uniqueCount = 0
    
    For i = 1 To nodeCount
        nodePoints = tempShp.Nodes(i).Points
        x = nodePoints(1, 1)
        y = nodePoints(1, 2)
        
        isDuplicate = False
        For j = 1 To uniqueCount
            If Abs(uniquePoints(j).x - x) < 1 And Abs(uniquePoints(j).y - y) < 1 Then
                isDuplicate = True
                Exit For
            End If
        Next j
        
        If Not isDuplicate Then
            uniqueCount = uniqueCount + 1
            uniquePoints(uniqueCount).x = x
            uniquePoints(uniqueCount).y = y
        End If
    Next i
    
    If uniqueCount <> 5 Then
        MsgBox "Shape must have exactly 5 corners.", vbCritical
        tempShp.Delete
        GetPentagonGeometry = False
        Exit Function
    End If
    
    ' 4. Find TIP (Unique X and Unique Y)
    tipIndex = 0
    Dim sharesX As Boolean, sharesY As Boolean
    
    For i = 1 To uniqueCount
        sharesX = False: sharesY = False
        For j = 1 To uniqueCount
            If i <> j Then
                If Abs(uniquePoints(i).x - uniquePoints(j).x) < 1 Then sharesX = True
                If Abs(uniquePoints(i).y - uniquePoints(j).y) < 1 Then sharesY = True
            End If
        Next j
        
        If sharesX = False And sharesY = False Then
            tipIndex = i
            Exit For
        End If
    Next i
    
    If tipIndex = 0 Then
        MsgBox "Could not identify a unique Tip.", vbCritical
        tempShp.Delete
        GetPentagonGeometry = False
        Exit Function
    End If
    
    ' Return Tip Coordinates
    tX = uniquePoints(tipIndex).x
    tY = uniquePoints(tipIndex).y
    
    ' 5. Find SHOULDERS (2 Closest points to Tip)
    Dim dist As Double
    For i = 1 To uniqueCount
        If i <> tipIndex Then
            dist = Sqr((uniquePoints(i).x - tX) ^ 2 + (uniquePoints(i).y - tY) ^ 2)
            uniquePoints(i).DistanceToTip = dist
        Else
            uniquePoints(i).DistanceToTip = 999999
        End If
    Next i
    
    ' Find indices of 2 smallest distances
    Dim idx1 As Integer, idx2 As Integer
    Dim min1 As Double, min2 As Double
    min1 = 999999: min2 = 999999
    
    For i = 1 To uniqueCount
        If i <> tipIndex Then
            If uniquePoints(i).DistanceToTip < min1 Then
                min2 = min1
                idx2 = idx1
                min1 = uniquePoints(i).DistanceToTip
                idx1 = i
            ElseIf uniquePoints(i).DistanceToTip < min2 Then
                min2 = uniquePoints(i).DistanceToTip
                idx2 = i
            End If
        End If
    Next i
    
    ' Return Average Shoulder Coordinates
    sX = (uniquePoints(idx1).x + uniquePoints(idx2).x) / 2
    sY = (uniquePoints(idx1).y + uniquePoints(idx2).y) / 2
    
    tempShp.Delete
    GetPentagonGeometry = True
    
End Function

