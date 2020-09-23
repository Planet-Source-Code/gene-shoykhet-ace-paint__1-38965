VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.OCX"
Begin VB.Form frmPaint 
   Caption         =   "PaintPro"
   ClientHeight    =   8730
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10050
   Icon            =   "PaintPro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Background"
      Height          =   1575
      Left            =   5520
      TabIndex        =   30
      Top             =   7080
      Width           =   1455
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   32
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   31
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Random"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Solid Fill"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   375
      Left            =   8520
      TabIndex        =   18
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Save &As"
      Height          =   375
      Left            =   7200
      TabIndex        =   17
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton cmdTool 
      Caption         =   "F&an"
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Brush Size"
      Height          =   1575
      Left            =   3480
      TabIndex        =   24
      Top             =   7080
      Width           =   1815
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   14
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   13
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   12
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "GIT!!!"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Larger Yet"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Large"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Medium"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Small"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color"
      Height          =   1575
      Left            =   1920
      TabIndex        =   22
      Top             =   7080
      Width           =   1335
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   825
         TabIndex        =   23
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "&Color"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog cdlColor 
      Left            =   120
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   8520
      TabIndex        =   19
      Top             =   8160
      Width           =   1095
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   1920
      MousePointer    =   2  'Cross
      ScaleHeight     =   6585
      ScaleWidth      =   7665
      TabIndex        =   21
      Top             =   360
      Width           =   7695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tools"
      Height          =   7575
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   1335
      Begin VB.CommandButton cmdTool 
         Caption         =   "&Eraser"
         Height          =   615
         Index           =   10
         Left            =   120
         Picture         =   "PaintPro.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CommandButton cmdTool 
         Caption         =   "&Free Line"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdTool 
         Caption         =   "C&ircle Fill"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdTool 
         Caption         =   "Sq&uare Fill"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdTool 
         Caption         =   "S&quare"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdTool 
         Caption         =   "&Circle"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdTool 
         Caption         =   "&Point"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdTool 
         Caption         =   "&Line"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuTool 
         Caption         =   "Point"
         Index           =   0
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Line"
         Index           =   1
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Free Line"
         Index           =   2
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Circle"
         Index           =   3
      End
      Begin VB.Menu mnuTool 
         Caption         =   "CircleFill"
         Index           =   4
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Square"
         Index           =   5
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Square Fill"
         Index           =   6
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Fan"
         Index           =   7
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Eraser"
         Index           =   10
      End
   End
   Begin VB.Menu mnuBack 
      Caption         =   "Background"
      Begin VB.Menu mnuBG 
         Caption         =   "Solid"
         Index           =   0
      End
      Begin VB.Menu mnuBG 
         Caption         =   "Random"
         Index           =   1
      End
   End
   Begin VB.Menu mnuClr 
      Caption         =   "Color"
      Begin VB.Menu mnuColor 
         Caption         =   "Choose Color"
      End
   End
   Begin VB.Menu mnuBrushSize 
      Caption         =   "Brush Size"
      Begin VB.Menu mnuBrush 
         Caption         =   "Small"
         Index           =   0
      End
      Begin VB.Menu mnuBrush 
         Caption         =   "Medium"
         Index           =   1
      End
      Begin VB.Menu mnuBrush 
         Caption         =   "Large"
         Index           =   2
      End
      Begin VB.Menu mnuBrush 
         Caption         =   "Larger Yet"
         Index           =   3
      End
      Begin VB.Menu mnuBrush 
         Caption         =   "GIT!!!"
         Index           =   4
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *****************************************************************************
' Project:          PaintPro
' Version:          1.2
' Module:           frmPaint
' Original Author:  Gene Shoykhet
' Modified By:
' Date:             9/11/02 11:08:33 AM
' *****************************************************************************

Option Explicit

Private CurentX As Single
Private CurentY As Single
Private CurentWidth As Single

Private mudtSquare As UDT_Square
Private mudtCircle As UDT_Circle
Private mudtLine As UDT_Line
Private mudtTool As UDT_Tool

Private Sub cmdColor_Click()
    'bring up the color dialog box
    cdlColor.ShowColor
    lngColor = cdlColor.color
    'set the preview box color
    picColor.BackColor = lngColor
End Sub

Private Sub cmdExit_Click()
Dim intExitChoice As Integer
'make sure user wants to exit and/or save
   If blnModified Then
      intExitChoice = MsgBox("Would you like to save before exiting?", vbYesNoCancel, "Exit")
        If intExitChoice = vbYes Then
            cmdSave_Click
            End
        ElseIf intExitChoice = vbNo Then
            End
        Else
            Exit Sub
        End If
   End If
   End
End Sub

Private Sub cmdNew_Click()
    'if picture modified, ask for a save
    If blnModified Then
        If MsgBox("Erase without saving?", vbInformation + vbOKCancel, "New") = vbOK Then
            'clear the picture control
            picMain.Cls
            picMain.Refresh
            strFilename = strNewFile
            frmPaint.Caption = strTitle & " v" & strVersion & " <" & strFilename & ">"
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdOpen_Click()
    On Error GoTo Err

    'bring up the open dialog box
    cdlColor.Filter = "Bitmap" & "(*.bmp)|*.bmp|Jpeg Files (*.jpg)|*.jpg"
    cdlColor.FilterIndex = 1
    cdlColor.ShowOpen
    cdlColor.CancelError = False
    strFilename = cdlColor.FileName
    picMain.Picture = LoadPicture(strFilename)
    frmPaint.Caption = strTitle & " " & strVersion & " <" & cdlColor.FileTitle & ">"
Exit Sub

Err:
    'do nothing here

End Sub

Private Sub cmdSaveAs_Click()
    On Error GoTo Err
    
    'bring up the save dialog box
    cdlColor.Filter = "Bitmap" & "(*.bmp)|*.bmp|Jpeg Files (*.jpg)|*.jpg"
    cdlColor.FilterIndex = 1
    cdlColor.ShowSave
    If cdlColor.Flags = vbCancel Then End
    strFilename = cdlColor.FileName
    SavePicture picMain.Image, cdlColor.FileName
    frmPaint.Caption = strTitle & " " & strVersion & " <" & cdlColor.FileTitle & ">"
Exit Sub

Err:
    'do nothing here
End Sub

Private Sub cmdTool_Click(Index As Integer)
    'set UDT tool to clear previous tool
    SetAllFalse mudtTool
    'select tool and attributes
    With mudtTool
        Select Case Index
            Case 0      'single point
                .Point = True
            Case 1      'straight line
                .Line = True
            Case 2      'freehand line
                .FreeLine = True
            Case 3      'empty circle
                picMain.FillStyle = vbFSTransparent
                .Circle = True
            Case 4      'filled circle
                picMain.FillStyle = vbFSSolid
                .Circle = True
            Case 5      'empty rectangle
                mudtSquare.blnFill = False
                picMain.FillStyle = vbFSTransparent
                .Square = True
            Case 6      'filled rectangle
                mudtSquare.blnFill = True
                picMain.FillColor = lngColor
                .Square = True
            Case 7      'fan effect
                .Fan = True
            Case 8      'fill background with current selected color
                picMain.BackColor = lngColor
            Case 9      'fill background with random color dots
                RandomizeBackground Me
            Case 10     'eraser
                picMain.DrawWidth = CurentWidth + 5
                .Eraser = True
        End Select
    End With
End Sub

Private Sub Form_Load()
    frmPaint.MousePointer = vbDefault
    strFilename = strNewFile
    
    'set the caption to static app title and version defined in modPaint.bas
    frmPaint.Caption = strTitle & " v" & strVersion & " <" & strFilename & ">"
    
    'set the default background color and point width
    picColor.BackColor = vbBlack
    picMain.DrawWidth = 1
    lngColor = vbBlack
    
    'set other defaults
    Option1(0).Value = True
    blnModified = False
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuBG_Click(Index As Integer)
    Option2(Index).Value = True
    Option2_Click Index
End Sub

Private Sub mnuBrush_Click(Index As Integer)
    Option1(Index).Value = True
    Option1_Click (Index)
End Sub

Private Sub mnuColor_Click()
    cmdColor_Click
End Sub

Private Sub mnuExit_Click()
    cmdExit_Click
End Sub

Private Sub mnuNew_Click()
    cmdNew_Click
End Sub

Private Sub mnuSave_Click()
    cmdSave_Click
End Sub

Private Sub mnuSaveAs_Click()
    cmdSaveAs_Click
End Sub

Private Sub mnuTool_Click(Index As Integer)
    cmdTool_Click Index
End Sub

Private Sub Option1_Click(Index As Integer)
    'select brush size
    With picMain
        If Index = 0 Then
            CurentWidth = 1
        ElseIf Index = 1 Then
            CurentWidth = 3
        ElseIf Index = 2 Then
            CurentWidth = 5
        ElseIf Index = 3 Then
            CurentWidth = 9
        ElseIf Index = 4 Then
            CurentWidth = 20
        End If
        If mudtTool.Eraser Then CurentWidth = CurentWidth + 5
        .DrawWidth = CurentWidth
    End With
End Sub

Private Sub Option2_Click(Index As Integer)
    Select Case Index
        Case 0  'call solid fill
            cmdTool_Click 8
        Case 1  'call random fill
            cmdTool_Click 9
    End Select
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'check to see which tool is selected and act accordingly
  If Button = vbLeftButton Then
    With mudtTool
        If .Line Then
            With mudtLine
                .mintLineX1 = x
                .mintLineY1 = y
            End With
        ElseIf .Point Then
            blnModified = True
            picMain.PSet (x, y), lngColor
        ElseIf .Circle Then
            With mudtCircle
                .mintCircleX1 = x
                .mintCircleY1 = y
            End With
        ElseIf .Square Then
            With mudtSquare
                .mintSquareX1 = x
                .mintSquareY1 = y
            End With
        ElseIf .Eraser Then
            picMain.PSet (x, y), picMain.BackColor
        ElseIf .FreeLine Then
            blnModified = True
            CurentX = x
            CurentY = y
        ElseIf .Fan Then
            blnModified = True
            CurentX = x
            CurentY = y
        End If
    End With
  End If
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        With mudtTool
            If .FreeLine Then
                picMain.Line (CurentX, CurentY)-(x, y), lngColor
                CurentX = x
                CurentY = y
            ElseIf .Fan Then
                picMain.Line (CurentX, CurentY)-(x, y), lngColor
            ElseIf .Eraser Then
                picMain.PSet (x, y), picMain.BackColor
            End If
        End With
    End If
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then
    With mudtTool
        If .Line Then
            With mudtLine
                .mintLineX2 = x
                .mintLineY2 = y
                blnModified = True
                picMain.Line (.mintLineX1, .mintLineY1)-(.mintLineX2, .mintLineY2), lngColor
            End With
        ElseIf .Circle Then
            With mudtCircle
                .mintCircleX2 = x
                .mintCircleY2 = y
                .mdblCircleR = distancePoints(.mintCircleX1, .mintCircleX2, .mintCircleY1, .mintCircleY2) / 2
                picMain.FillColor = lngColor
                blnModified = True
                picMain.Circle (.mintCircleX1, .mintCircleY1), .mdblCircleR, lngColor
            End With
        ElseIf .Square Then
            With mudtSquare
                .mintSquareX2 = x
                .mintSquareY2 = y
                If Not .blnFill Then
                    blnModified = True
                    picMain.Line (.mintSquareX1, .mintSquareY1)-(.mintSquareX2, .mintSquareY2) _
                        , lngColor, B
                Else
                    blnModified = True
                    picMain.Line (.mintSquareX1, .mintSquareY1)-(.mintSquareX2, .mintSquareY2) _
                        , lngColor, BF
                End If
            End With
        End If
    End With
  End If
End Sub

Private Sub cmdSave_Click()
    'check to see if file was already saved and has a name
    If strFilename <> strNewFile Then
        SavePicture picMain.Image, strFilename
    Else
        cmdSaveAs_Click
    End If
End Sub

