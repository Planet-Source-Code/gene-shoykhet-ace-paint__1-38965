Attribute VB_Name = "modPaint"
' *****************************************************************************
' Project:          PaintPro
' Version:          1.2
' Module:           modPaint
' Original Author:  Gene Shoykhet
' Modified By:
' Date:             9/11/02 11:08:33 AM
' *****************************************************************************

Option Explicit

Public Const strTitle = "Paint Pro"
Public Const strVersion = "1.2"
Public Const strNewFile = "Untitled"

Public lngColor As Long
Public blnModified As Boolean

Public strFilename As String

Public Type UDT_Tool   'UDT for tool used
    Line As Boolean
    FreeLine As Boolean
    Circle As Boolean
    Point As Boolean
    Square As Boolean
    Eraser As Boolean
    Fan As Boolean
End Type

Public Type UDT_Square     'UDT for square
    mintSquareX1 As Single
    mintSquareX2 As Single
    mintSquareY1 As Single
    mintSquareY2 As Single
    blnFill As Boolean
End Type

Public Type UDT_Line       'UDT for Line
    mintLineX1 As Single
    mintLineX2 As Single
    mintLineY1 As Single
    mintLineY2 As Single
End Type

Public Type UDT_Circle     'UDT for Circle
    mintCircleX1 As Single
    mintCircleY1 As Single
    mintCircleX2 As Single
    mintCircleY2 As Single
    mdblCircleR As Single
End Type

Public Function SetAllFalse(ByRef rudtTool As UDT_Tool) As Boolean
    'reset the UDT tool
    With rudtTool
        .Circle = False
        .FreeLine = False
        .Line = False
        .Point = False
        .Square = False
        .Eraser = False
        .Fan = False
    End With
    SetAllFalse = True
End Function

Public Function distancePoints(X1 As Single, X2 As Single, Y1 As Single, Y2 As Single) As Double
    On Error GoTo Err
    
    Dim xses As Double
    Dim yses As Double
    
    'find the distance between two points
    xses = (X2 - X1) ^ 2
    yses = (Y2 - Y1) ^ 2
    'return the distance
    distancePoints = Sqr(xses + yses)
    Exit Function
Err:
    'do nothing
End Function
    
Public Function RandomizeBackground(ByRef frmName As Form)
    Dim x As Integer
    Dim y As Integer
    Dim color As Integer
    Dim Index As Integer
    
    DoEvents
    For Index = 1 To 10000
        With frmName
            x = Rnd * .picMain.ScaleWidth
            y = Rnd * .picMain.ScaleHeight
            color = Rnd * 15
            .picMain.PSet (x, y), QBColor(color)
        End With
    Next
End Function
