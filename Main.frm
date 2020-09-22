VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Begin VB.Form Main 
   Caption         =   "Voronoi test"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11505
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   647
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   767
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFloodFillOutline 
      Height          =   570
      Left            =   4320
      Picture         =   "Main.frx":49E2
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "GDI_API algorithm (outlines only)"
      Top             =   4080
      Width           =   570
   End
   Begin VB.TextBox comLine 
      BackColor       =   &H8000000F&
      ForeColor       =   &H80000015&
      Height          =   570
      Left            =   4680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7680
      Width           =   4695
   End
   Begin VB.PictureBox picRegion 
      AutoRedraw      =   -1  'True
      Height          =   975
      Left            =   1920
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   18
      Top             =   6120
      Width           =   825
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   6120
      ScaleHeight     =   159
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdExact 
      Height          =   570
      Left            =   3720
      Picture         =   "Main.frx":5624
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "GDI_API algorithm"
      Top             =   4080
      Width           =   570
   End
   Begin VB.HScrollBar CPRadiusScroll 
      Height          =   180
      Left            =   2880
      Max             =   200
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5400
      Value           =   20
      Width           =   1560
   End
   Begin VB.CommandButton cmdAttract 
      Height          =   570
      Left            =   7320
      Picture         =   "Main.frx":6266
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   570
   End
   Begin VB.CommandButton cmdDisperse 
      Height          =   570
      Left            =   6720
      Picture         =   "Main.frx":6EA8
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6240
      Width           =   570
   End
   Begin VB.OptionButton optAddDelete 
      Height          =   570
      Index           =   1
      Left            =   6840
      Picture         =   "Main.frx":7AEA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   570
   End
   Begin VB.OptionButton optAddDelete 
      Height          =   570
      Index           =   0
      Left            =   6240
      Picture         =   "Main.frx":872C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Value           =   -1  'True
      Width           =   570
   End
   Begin VB.CommandButton cmdClosestSample 
      Height          =   570
      Left            =   3120
      Picture         =   "Main.frx":936E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Closest sample algorithm"
      Top             =   4080
      Width           =   570
   End
   Begin VB.TextBox txtColour 
      Enabled         =   0   'False
      Height          =   975
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   5880
      Width           =   825
   End
   Begin VB.TextBox Components 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   2
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Blue"
      Top             =   6720
      Width           =   735
   End
   Begin VB.TextBox Components 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Index           =   1
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Green"
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox Components 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   0
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Red"
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton cmdSquareGrowth 
      Height          =   570
      Left            =   2520
      Picture         =   "Main.frx":9FB0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Square growth algorithm"
      Top             =   4080
      Width           =   570
   End
   Begin VB.CommandButton cmdDrawSamples 
      Height          =   570
      Left            =   7440
      Picture         =   "Main.frx":ABF2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Draw all samples"
      Top             =   4080
      Width           =   570
   End
   Begin MSComDlg.CommonDialog cmdBox 
      Left            =   8400
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "bmp"
      DialogTitle     =   "Save the image"
      FileName        =   "Voronoi_diagram"
      Filter          =   "Bitmap *.bmp|*.bmp"
   End
   Begin VB.CommandButton cmdSpiralGrowth 
      Height          =   570
      Left            =   1920
      Picture         =   "Main.frx":B834
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Spiral growth algorithm"
      Top             =   4080
      Width           =   570
   End
   Begin VB.CommandButton cmdClear 
      Height          =   570
      Left            =   6120
      Picture         =   "Main.frx":C476
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Remove all samples"
      Top             =   6240
      Width           =   570
   End
   Begin VB.PictureBox picDiagram 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3540
      Left            =   1920
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   232
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   0
      Top             =   360
      Width           =   3960
   End
   Begin VB.PictureBox ColourBar 
      AutoRedraw      =   -1  'True
      Height          =   3015
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   15
      Top             =   240
      Width           =   1560
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSaveImage 
         Caption         =   "&Save image"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPlaceImage 
         Caption         =   "&Place image"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuClearBackground 
         Caption         =   "Clear image"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveSamples 
         Caption         =   "Save samples"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLoadSamples 
         Caption         =   "Load Samples"
      End
      Begin VB.Menu mnuLoadExampleMap 
         Caption         =   "Load Example map"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuSamples 
      Caption         =   "&Samples"
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuRemoveLast 
         Caption         =   "&Remove last"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuCredits 
         Caption         =   "&Credits"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&Context help..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Source code copyrighted by Gelfling '04 aka David Rutten

Public GlobalHalt As Boolean        'Used throughout the app to cancel long running loops
Public cR As Byte                   'Red component of current colour
Public cG As Byte                   'Green component of current colour
Public cB As Byte                   'Blue component of current colour
Public cC As Long                   'Current colour

Private CPx As Long                 'x-coordinate of colour picker target
Private CPy As Long                 'y-coordinate of colour picker target
Private CPr As Long                 'radius of colour picker target

Private ArrowTarget As Boolean      'I forgot what this is for
Private UsePicker As Boolean        'Use colour from picker or use colour from image

'This function returns a colour that is compliant with the current settings.
'Note that complete black and complete white are NEVER used.
Public Function GetRegionalColour() As Long
    If UsePicker Then
        Dim X As Long
        Dim Y As Long
        
        X = CPx + (Rnd * 2 - 1) * Region(CPr, 0, ColourBar.ScaleWidth - 1)
        Y = CPy + (Rnd * 4 - 2) * Region(CPr, 0, ColourBar.ScaleHeight - 1)
        X = Region(X, 0, ColourBar.ScaleWidth - 1)
        Y = Region(Y, 0, ColourBar.ScaleHeight - 1)
        DrawColourPickerTarget
        GetRegionalColour = GetPixel(ColourBar.hdc, X, Y)
        DrawColourPickerTarget
    Else
        GetRegionalColour = cC
    End If
    
    If GetRegionalColour = 0 Then
        GetRegionalColour = 1
    ElseIf GetRegionalColour = vbWhite Then
        GetRegionalColour = RGB(255, 255, 254)
    End If
End Function

Private Sub cmdAttract_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Dim dblDemp As Double
    
    dblDemp = -10
    GlobalHalt = False
AddLogEntry "Attraction sequence initiated"
    Do
        If GlobalHalt Then Exit Sub
        Math.DisperseSamples 500, 20, dblDemp
        dblDemp = Max(dblDemp + 0.01, -2)
        
        picDiagram.Cls
        DrawSamples True
        picDiagram.Refresh
        DoEvents
    Loop
End Sub

Private Sub cmdAttract_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GlobalHalt = True
End Sub

Private Sub cmdClear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Erase arrSamples
        SampleBound = -1
        picDiagram.Cls
AddLogEntry "All samples deleted"
    ElseIf Button = 2 Then
        Set picDiagram.Picture = Nothing
        picDiagram.Cls
        DrawSamples True
    End If
End Sub

Private Sub cmdDisperse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim curBrush As Long
    Dim oldBrush As Long
    
    If Button <> 1 Then Exit Sub
    GlobalHalt = False
AddLogEntry "Dispersion sequence initiated"
    Do
        If GlobalHalt Then Exit Sub
        Math.DisperseSamples 1000, 20, 20
        
        curBrush = CreateSolidBrush(RGB(224, 224, 224))
        oldBrush = SelectObject(picDiagram.hdc, curBrush)
        DeleteObject oldBrush
        Rectangle picDiagram.hdc, -1, -1, picDiagram.ScaleWidth + 1, picDiagram.ScaleHeight + 1
        DrawSamples True
        picDiagram.Refresh
        DoEvents
    Loop
End Sub

Private Sub cmdDisperse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GlobalHalt = True
End Sub

Private Sub cmdExact_Click()
AddLogEntry "GDI_API floodfill algorithm called"
    Voronoi.FloodFillVoronoiEx picDiagram
AddLogEntry "GDI_API floodfill algorithm finished"
    DrawSamples True
End Sub

Private Sub cmdFloodFillOutline_Click()
AddLogEntry "GDI_API floodfill outlines algorithm called"
    Voronoi.FloodFillVoronoiOutline picDiagram
AddLogEntry "GDI_API floodfill outlines algorithm finished"
    DrawSamples False
End Sub

Private Sub ColourBar_GotFocus()
AddLogEntry "Colourpicker [gotfocus]"
    If Not UsePicker Then
        UsePicker = True
        CopyColourBar
        ColourBar.Refresh
    End If
    ArrowTarget = True
End Sub

Private Sub ColourBar_LostFocus()
    ArrowTarget = False
End Sub

Private Sub ColourBar_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Effect = 7 Then
        Select Case UCase(Right(Data.Files(1), 4))
        Case ".BMP", ".GIF", ".JPG", "JPEG", ".DIB"
            Set picBuffer.Picture = LoadPicture(Data.Files(1))
            CopyColourBar
            CopyColourRegion
        Case Else
            MsgBox "Unrecognized image file format...", vbOKOnly, "Sorry..."
        End Select
    End If
End Sub

Private Sub CPRadiusScroll_Change()
    DrawColourPickerTarget
    CPr = CPRadiusScroll.Value
    CopyColourRegion
    DrawColourPickerTarget
End Sub

Private Sub CPRadiusScroll_GotFocus()
    If Not UsePicker Then
        UsePicker = True
        CopyColourBar
        ColourBar.Refresh
    End If
    ArrowTarget = True
End Sub

Private Sub CPRadiusScroll_Scroll()
    DrawColourPickerTarget
    CPr = CPRadiusScroll.Value
    CopyColourRegion
    DrawColourPickerTarget
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then GlobalHalt = True
    If Not ArrowTarget Then Exit Sub
    
    DrawColourPickerTarget
    Select Case KeyCode
    Case 38 'arrow up
        If Shift = 0 Then
            CPy = Region(CPy - 1, 0, ColourBar.ScaleHeight - 1)
        Else
            CPy = Region(CPy - 10, 0, ColourBar.ScaleHeight - 1)
        End If
    Case 40 'arrow down
        If Shift = 0 Then
            CPy = Region(CPy + 1, 0, ColourBar.ScaleHeight - 1)
        Else
            CPy = Region(CPy + 10, 0, ColourBar.ScaleHeight - 1)
        End If
    Case 37 'arrow left
        If Shift = 0 Then
            CPx = Region(CPx - 1, 0, ColourBar.ScaleWidth - 1)
        Else
            CPx = Region(CPx - 10, 0, ColourBar.ScaleWidth - 1)
        End If
    Case 39 'arrow right
        If Shift = 0 Then
            CPx = Region(CPx + 1, 0, ColourBar.ScaleWidth - 1)
        Else
            CPx = Region(CPx + 10, 0, ColourBar.ScaleWidth - 1)
        End If
    Case Else
        'Nothing yet
    End Select
    LoadColourFromPicker
    DrawColourPickerTarget
End Sub

Private Sub Form_Load()
AddLogEntry "Application started on " & CStr(Now)
    Erase arrSamples
    SampleBound = -1
    ArrowTarget = False
    UsePicker = True
    GlobalHalt = False
    Randomize
AddLogEntry "Initialization completed"
    Me.Show
    
    'Draw banner
    Set picBuffer.Picture = LoadPicture(App.Path & "\Banner.gif")
    Call StretchBlt(picDiagram.hdc, _
                    picDiagram.ScaleWidth \ 2 - picBuffer.ScaleWidth \ 2, _
                    picDiagram.ScaleHeight \ 2 - picBuffer.ScaleHeight \ 2, _
                    picBuffer.ScaleWidth, picBuffer.ScaleHeight, _
                    picBuffer.hdc, 0, 0, picBuffer.ScaleWidth, picBuffer.ScaleHeight, &HCC0020)
    picDiagram.Refresh
AddLogEntry "Banner loaded"

    'Draw colour picker image
    Set picBuffer.Picture = LoadPicture(App.Path & "\ColourPicker.bmp")
AddLogEntry "Colour picker palette loaded"
    CopyColourBar
    ColourBar.DrawMode = vbInvert
    ColourBar.FillStyle = 1
    CPx = 50
    CPy = 250
    CPr = 20
    LoadColourFromPicker
    DrawColourPickerTarget
    cC = RGB(cR, cG, cB)
AddLogEntry "Colourpicker initialized"
End Sub

'draw the target of the colour picker using vbINVERT mode. If the colourpicker is not active
'then draw the hatch.
Private Sub DrawColourPickerTarget()
    If UsePicker Then
        ColourBar.Line (CPx - 10, CPy)-(CPx + 11, CPy)
        ColourBar.Line (CPx, CPy - 10)-(CPx, CPy + 11)
        ColourBar.Line (CPx - CPr, CPy - 2 * CPr)-(CPx + CPr, CPy - 2 * CPr)
        ColourBar.Line (CPx + CPr, CPy - 2 * CPr)-(CPx + CPr, CPy + 2 * CPr)
        ColourBar.Line (CPx + CPr, CPy + 2 * CPr)-(CPx - CPr, CPy + 2 * CPr)
        ColourBar.Line (CPx - CPr, CPy + 2 * CPr)-(CPx - CPr, CPy - 2 * CPr)
        picRegion.Visible = True
    Else
        Dim i As Long
        Dim h As Long
        
        h = ColourBar.ScaleHeight
        For i = 0 To h + ColourBar.ScaleWidth Step 4
            ColourBar.Line (0, i)-(h, i - h)
        Next
        picRegion.Visible = False
    End If
End Sub

'Retrieve the colour of the pixel directly under the target (ignore radius)
Private Sub LoadColourFromPicker()
    Dim rgbPT As Long
    rgbPT = GetPixel(ColourBar.hdc, CPx, CPy)
    RGB_Components rgbPT, cR, cG, cB
    Components(0).Text = cR
    Components(1).Text = cG
    Components(2).Text = cB
    txtColour.BackColor = RGB(cR, cG, cB)
    
    CopyColourRegion
End Sub

Private Sub ColourBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    DrawColourPickerTarget
    CPx = Region(X, 0, ColourBar.ScaleWidth - 1)
    CPy = Region(Y, 0, ColourBar.ScaleHeight - 1)
    LoadColourFromPicker
    CopyColourRegion
    DrawColourPickerTarget
    DoEvents
End Sub

Private Sub ColourBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    DrawColourPickerTarget
    CPx = Region(X, 0, ColourBar.ScaleWidth - 1)
    CPy = Region(Y, 0, ColourBar.ScaleHeight - 1)
    LoadColourFromPicker
    CopyColourRegion
    DrawColourPickerTarget
    DoEvents
End Sub

Private Sub ColourBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    DrawColourPickerTarget
    CPx = Region(X, 0, ColourBar.ScaleWidth - 1)
    CPy = Region(Y, 0, ColourBar.ScaleHeight - 1)
    LoadColourFromPicker
    CopyColourRegion
    DrawColourPickerTarget
    DoEvents
End Sub

Private Sub cmdClosestSample_Click()
AddLogEntry "Solution per pixel algorithm called"
    Voronoi.ClosestSampleVoronoi picDiagram
AddLogEntry "Solution per pixel algorithm finished"
    DrawSamples True
End Sub

Private Sub cmdDrawSamples_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        DrawSamples True
    ElseIf Button = 2 Then
        picDiagram.Cls
    End If
End Sub

Private Sub cmdSpiralGrowth_Click()
AddLogEntry "Spiral growth algorithm called"
    Call Voronoi.CircularVoronoi(picDiagram)
AddLogEntry "Spiral growth algorithm finished"
    DrawSamples True
End Sub

Private Sub cmdSquareGrowth_Click()
AddLogEntry "Square growth algorithm called"
    Call Voronoi.RectangularVoronoi(picDiagram)
AddLogEntry "Square growth algorithm finished"
    DrawSamples True
End Sub

'Returns the index of the sample closest to the specified coordinates
Public Function FindClosestSample(ByVal X As Long, ByVal Y As Long) As Long
    FindClosestSample = -1
    If SampleBound < 0 Then Exit Function
    Dim d As Double
    Dim minD As Double
    Dim i As Long, lowI As Long

    If SampleBound = 0 Then
        FindClosestSample = 0
        Exit Function
    End If
    
    minD = Voronoi.Distance(Array(X, Y), Array(arrSamples(0).X, arrSamples(0).Y))
    lowI = 0
    For i = 1 To SampleBound
        d = Distance(Array(X, Y), Array(arrSamples(i).X, arrSamples(i).Y))
        If d < minD Then
            minD = d
            lowI = i
        End If
    Next
    FindClosestSample = lowI
End Function

'Draw all samples. Samples are drawn as filled circles or crosses.
Public Sub DrawSamples(Optional blnDrawAsCircle As Boolean = True, _
                       Optional bytRadius As Long = 2)
    On Error GoTo ErrorTrap
    
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    
    picDiagram.FillStyle = 0
    If blnDrawAsCircle Then
        For i = 0 To UBound(arrSamples)
            picDiagram.FillColor = arrSamples(i).C
            picDiagram.Circle (arrSamples(i).X, arrSamples(i).Y), bytRadius, 0
        Next
    Else
        For i = 0 To UBound(arrSamples)
            picDiagram.Line (arrSamples(i).X - bytRadius, arrSamples(i).Y)-(arrSamples(i).X + bytRadius + 1, arrSamples(i).Y)
            picDiagram.Line (arrSamples(i).X, arrSamples(i).Y - bytRadius)-(arrSamples(i).X, arrSamples(i).Y + bytRadius + 1)
        Next
    End If
ErrorTrap:
    picDiagram.Refresh
    Err.Clear
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrorTrap
    
    'Reposition and resize all interface elements
    ColourBar.Left = 0
    ColourBar.Top = 0
    ColourBar.Height = Main.ScaleHeight - txtColour.Height - CPRadiusScroll.Height
    picDiagram.Top = 0
    picDiagram.Left = ColourBar.Width + 8
    picDiagram.Height = ColourBar.Height
    picDiagram.Width = Main.ScaleWidth - picDiagram.Left
    CPRadiusScroll.Top = ColourBar.Height
    CPRadiusScroll.Left = 0
    txtColour.Top = ColourBar.Height + CPRadiusScroll.Height + 1
    txtColour.Left = 0
    picRegion.Top = txtColour.Top
    picRegion.Left = txtColour.Left
    Components(0).Left = txtColour.Width
    Components(0).Top = txtColour.Top
    Components(1).Left = txtColour.Width
    Components(1).Top = txtColour.Top + Components(0).Height
    Components(2).Left = txtColour.Width
    Components(2).Top = txtColour.Top + 2 * Components(0).Height
    
    cmdSpiralGrowth.Left = picDiagram.Left
    cmdSpiralGrowth.Top = picDiagram.Height + 1
    cmdSquareGrowth.Left = cmdSpiralGrowth.Left + cmdSpiralGrowth.Width
    cmdSquareGrowth.Top = cmdSpiralGrowth.Top
    cmdClosestSample.Left = cmdSquareGrowth.Left + cmdSquareGrowth.Width
    cmdClosestSample.Top = cmdSquareGrowth.Top
    cmdExact.Left = cmdClosestSample.Left + cmdClosestSample.Width
    cmdExact.Top = cmdClosestSample.Top
    cmdFloodFillOutline.Left = cmdExact.Left + cmdExact.Width
    cmdFloodFillOutline.Top = cmdExact.Top
    
    optAddDelete(0).Left = cmdFloodFillOutline.Left + 1.5 * cmdFloodFillOutline.Width
    optAddDelete(0).Top = cmdFloodFillOutline.Top
    optAddDelete(1).Left = optAddDelete(0).Left + optAddDelete(0).Width
    optAddDelete(1).Top = optAddDelete(0).Top
    cmdDrawSamples.Left = optAddDelete(1).Left + optAddDelete(1).Width
    cmdDrawSamples.Top = optAddDelete(1).Top
    
    cmdClear.Left = cmdDrawSamples.Left + 1.5 * cmdDrawSamples.Width
    cmdClear.Top = cmdDrawSamples.Top
    cmdDisperse.Left = cmdClear.Left + cmdClear.Width
    cmdDisperse.Top = cmdClear.Top
    cmdAttract.Left = cmdDisperse.Left + cmdDisperse.Width
    cmdAttract.Top = cmdDisperse.Top
    
    comLine.Left = cmdSpiralGrowth.Left
    comLine.Top = cmdSpiralGrowth.Top + cmdSpiralGrowth.Height
    comLine.Width = Main.ScaleWidth - comLine.Left
    
    CopyColourBar
ErrorTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GlobalHalt = True
    End
End Sub

Private Sub mnuAbout_Click()
    ShellExecuteA Me.hwnd, "open", App.Path & "\Help.chm", "", App.Path, 1
End Sub

Private Sub mnuClear_Click()
    Erase arrSamples
    SampleBound = -1
    picDiagram.Cls
    mnuSaveSamples.Enabled = False
End Sub

Private Sub mnuClearBackground_Click()
    Set picDiagram.Picture = Nothing
    picDiagram.Cls
    DrawSamples True
End Sub

Private Sub mnuCredits_Click()
    MsgBox "Application written and copyrighted by:" & vbNewLine & _
           "Gelfling '04 aka. David Rutten" & vbNewLine & vbNewLine & _
           "Many thanks to Teake de Jong for asking questions." & vbNewLine & _
           "Many thanks to Andrew leBihan for providing answers." & vbNewLine & _
           "Many thanks to Kris & Pieter Phillipaerts for API-Guide.", _
           vbOKOnly, "Copyright and credits"
End Sub

Private Sub mnuLoadExampleMap_Click()
    On Error GoTo ErrorTrap
   
    Erase arrSamples
    SampleBound = -1
    Math.LoadSolution App.Path & "\World.samples"
   
    Exit Sub
ErrorTrap:
    MsgBox "Non-compatible image data..."
    Set picBuffer.Picture = LoadPicture(App.Path & "\ColourPicker.bmp")
    Err.Clear
End Sub

Private Sub mnuLoadSamples_Click()
    On Error GoTo ErrorTrap
    With cmdBox
        .DialogTitle = "Import existing voronoi samples..."
        .Filter = "Voronoi Samples (*.samples)|*.samples|All Files (*.*)|*.*||"
        .FileName = ""
        .ShowOpen
        If Len(cmdBox.FileName) = 0 Then Exit Sub
    End With
    
    Math.LoadSolution cmdBox.FileName
    
    Exit Sub
ErrorTrap:
    MsgBox "Non-compatible image data..."
    Set picBuffer.Picture = LoadPicture(App.Path & "\ColourPicker.bmp")
    Err.Clear
End Sub

Private Sub mnuPlaceImage_Click()
    On Error GoTo ErrorTrap
    With cmdBox
        .DialogTitle = "Import existing image..."
        .Filter = "Bitmaps (*.bmp)|*.bmp|GIF images (*.gif)|*.gif|JPEG images (*.jpg)|*.jpg|All Files (*.*)|*.*||"
        .FileName = ""
        .ShowOpen
        If Len(cmdBox.FileName) = 0 Then Exit Sub
    End With
    
    ''Uncomment these lines to place the image directly onto the viewport without scaling
    'Set picDiagram.Picture = LoadPicture(cmdBox.FileName)
    'picDiagram.Refresh
    'Exit sub
    
    Set picBuffer.Picture = LoadPicture(cmdBox.FileName)
    SetStretchBltMode picDiagram.hdc, 3
    StretchBlt picDiagram.hdc, 0, 0, picDiagram.ScaleWidth, picDiagram.ScaleHeight, _
               picBuffer.hdc, 0, 0, picBuffer.ScaleWidth, picBuffer.ScaleHeight, &HCC0020
    picDiagram.Refresh
    BurnCurrentImage
    
    Set picBuffer.Picture = LoadPicture(App.Path & "\ColourPicker.bmp")
    
    Exit Sub
ErrorTrap:
    MsgBox "Non-compatible image data..."
    Set picBuffer.Picture = LoadPicture(App.Path & "\ColourPicker.bmp")
    Err.Clear
End Sub

Private Sub mnuQuit_Click()
    End
End Sub

Private Sub mnuRemoveLast_Click()
    If SampleBound < 0 Then Exit Sub
    If SampleBound = 0 Then
        Erase arrSamples
        SampleBound = -1
        picDiagram.Cls
        mnuSaveSamples.Enabled = False
    Else
        ReDim Preserve arrSamples(SampleBound - 1)
        SampleBound = SampleBound - 1
        picDiagram.Cls
        DrawSamples True
    End If
End Sub

Private Sub mnuSaveImage_Click()
    With cmdBox
        .DialogTitle = "Save current voronoi diagram image..."
        .Filter = "Bitmap|*.bmp||"
        .FileName = "Voronoi_Diagram"
        .ShowSave
        If Len(cmdBox.FileName) = 0 Then Exit Sub
    End With
    SavePicture picDiagram.Image, CStr(cmdBox.FileName)
End Sub

Private Sub mnuSaveSamples_Click()
    With cmdBox
        .DialogTitle = "Save current voronoi samples..."
        .Filter = "Voronoi Samples (*.samples)|*.samples|All Files(*.*)|*.*||"
        .FileName = "Voronoi_Samples"
        .ShowSave
        If Len(cmdBox.FileName) = 0 Then Exit Sub
    End With
    Math.SaveSolution cmdBox.FileName
End Sub

Private Sub picDiagram_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If Button = 1 Then
        If optAddDelete(0).Value Then
            mnuSaveSamples.Enabled = True
            ReDim Preserve arrSamples(SampleBound + 1)
            SampleBound = SampleBound + 1
            
            If Not UsePicker Then
                cC = GetPixel(picDiagram.hdc, X, Y)
                RGB_Components cC, cR, cG, cB
                Components(0).Text = cR
                Components(1).Text = cG
                Components(2).Text = cB
                txtColour.BackColor = cC
            End If
            
            arrSamples(SampleBound).X = X
            arrSamples(SampleBound).Y = Y
            arrSamples(SampleBound).C = GetRegionalColour
AddLogEntry "Sample added on " & X & "," & Y & " with pal.index=" & arrSamples(SampleBound).C
            DrawSamples True
        Else
            If SampleBound < 0 Then Exit Sub
            Dim iclosest As Long
            iclosest = FindClosestSample(X, Y)
            If iclosest < 0 Then Exit Sub
            If SampleBound = 0 Then
                Erase arrSamples
                SampleBound = -1
                mnuSaveSamples.Enabled = False
AddLogEntry "Sample deleted"
            Else
AddLogEntry "Sample #" & iclosest + 1 & " deleted from " & arrSamples(iclosest).X & "," & arrSamples(iclosest).Y
                If iclosest < SampleBound Then
                    For i = iclosest To SampleBound - 1
                        arrSamples(i) = arrSamples(i + 1)
                    Next
                End If
                ReDim Preserve arrSamples(SampleBound - 1)
                SampleBound = SampleBound - 1
            End If
            picDiagram.Cls
            DrawSamples True
        End If
    End If
    
    If Button = 2 Then
AddLogEntry "Colourpicker source image switched from colourpickerbar to main viewport."
        cC = GetPixel(picDiagram.hdc, X, Y)
        RGB_Components cC, cR, cG, cB
        Components(0).Text = cR
        Components(1).Text = cG
        Components(2).Text = cB
        txtColour.BackColor = cC
        DrawColourPickerTarget
        UsePicker = False
        DrawColourPickerTarget
    End If
End Sub

'calculate the red, green and blue components of a long colour.
Public Function RGB_Components(ByVal rgbInput As Long, ByRef Red As Byte, ByRef Green As Byte, ByRef Blue As Byte)
    If rgbInput = -1 Then
        Red = 0
        Green = 0
        Blue = 0
    Else
        Red = (rgbInput And 255)
        Green = (rgbInput And 65280) / 256
        Blue = (rgbInput And 16711680) / 65536
    End If
End Function

'I don't seemto be using this one yet.
Public Sub WriteText(ByRef pic As PictureBox, _
                     ByVal X As Long, _
                     ByVal Y As Long, _
                     ByVal strText As String, _
                     Optional ByVal rgbColour As Long = 0)
    Call SetTextColor(pic.hdc, rgbColour)
    Call TextOut(pic.hdc, X, Y, strText, Len(strText))
End Sub

'Return the lowest of two values
Public Function Min(ByVal dblInput, ByVal dblMinValue) As Double
    If dblInput < dblMinValue Then
        Min = dblMinValue
    Else
        Min = dblInput
    End If
End Function

'Return the highest of two values
Public Function Max(ByVal dblInput, ByVal dblMaxValue) As Double
    If dblInput > dblMaxValue Then
        Max = dblMaxValue
    Else
        Max = dblInput
    End If
End Function

'Limit a numeric value to a numeric region
Public Function Region(ByVal dblInput, ByVal dblMinValue, ByVal dblMaxValue) As Double
    Region = dblInput
    Region = Min(Region, dblMinValue)
    Region = Max(Region, dblMaxValue)
End Function

'Place the colourbar banner into the variable size colourbar (it is stretched to fit)
Public Sub CopyColourBar()
    Call SetStretchBltMode(ColourBar.hdc, 3)
    Call StretchBlt(ColourBar.hdc, 0, 0, ColourBar.ScaleWidth, ColourBar.ScaleHeight, _
                    picBuffer.hdc, 0, 0, picBuffer.ScaleWidth, picBuffer.ScaleHeight, _
                    &HCC0020)
    ColourBar.Refresh
    LoadColourFromPicker
    DrawColourPickerTarget
End Sub

'Copy the rectangle defined by crX, crY and crR into the small colour feedback picturebox
Public Sub CopyColourRegion()
    Dim X1 As Long, Y1 As Long
    Dim X2 As Long, Y2 As Long
    
    X1 = CPx - CPr
    Y1 = CPy - 2 * CPr
    X2 = CPx + CPr
    Y2 = CPy + 2 * CPr
    
    X1 = Region(X1, 0, ColourBar.ScaleWidth - 1)
    X2 = Region(X2, 0, ColourBar.ScaleWidth - 1)
    Y1 = Region(Y1, 0, ColourBar.ScaleHeight - 1)
    Y2 = Region(Y2, 0, ColourBar.ScaleHeight - 1)
    
    If X1 = X2 Then
        picRegion.BackColor = RGB(cR, cG, cB)
        picRegion.Cls
    Else
        Call SetStretchBltMode(picRegion.hdc, 3)
        Call StretchBlt(picRegion.hdc, 0, 0, picRegion.ScaleWidth, picRegion.ScaleHeight, _
                        ColourBar.hdc, X1, Y1, X2 - X1, Y2 - Y1, &HCC0020)
    End If
    picRegion.Refresh
End Sub

'Save the current viewport as a bitmap and load it as background image (cls will no longer destroy these pixels.)
Public Sub BurnCurrentImage()
    SavePicture picDiagram.Image, App.Path & "\Temp_Graph.bmp"
    Set picDiagram.Picture = LoadPicture(App.Path & "\Temp_Graph.bmp")
End Sub

'Add a string of text to the logger textbox
Public Sub AddLogEntry(ByVal strEntry As String)
    comLine.Text = comLine.Text & vbNewLine & strEntry
    comLine.Text = Right(comLine.Text, 2000)
    comLine.SelStart = 2000
End Sub

Private Sub picDiagram_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorTrap
    If Effect = 7 Then
        Select Case UCase(Right(Data.Files(1), 4))
        Case ".BMP", ".GIF", ".JPG", "JPEG", ".DIB"
            Set picDiagram.Picture = LoadPicture(Data.Files(1))
            picDiagram.Refresh
            BurnCurrentImage
        Case Else
            MsgBox "Unrecognized image file format...", vbOKOnly, "Sorry..."
        End Select
    End If

    Exit Sub
ErrorTrap:
    MsgBox "Error in placing image", vbOKOnly Or vbExclamation, "Darn!"
End Sub
