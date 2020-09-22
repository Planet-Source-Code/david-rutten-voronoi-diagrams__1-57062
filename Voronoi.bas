Attribute VB_Name = "Voronoi"
Option Explicit

Public arrSamples() As Sample
Public SampleBound As Long

'Spiral growth algorithm
Public Sub CircularVoronoi(ByVal DC As PictureBox)
    On Error GoTo ErrorTrap
    If SampleBound < 0 Then Exit Sub
    Main.GlobalHalt = False
    
    Dim i As Long, j As Long
    Dim A As Long, b As Long
    Dim t As Double
    
    Dim pi As Double
    Dim idDC As Long
    
    Dim rgbSource As Long
    Dim blnFilter() As Boolean
    Dim r As Double
    Dim blnCompleted As Boolean
    Dim blnActive As Boolean
    Dim blnSamples() As Boolean
    
    'Set all samples to active
    ReDim blnSamples(UBound(arrSamples))
    For i = 0 To UBound(arrSamples)
        blnSamples(i) = True
    Next
    
    pi = 4 * Atn(1)
    idDC = DC.hdc
    
    'Create a boolean two dimensional array so we can keep track of every pixel we place.
    ReDim blnFilter(DC.ScaleWidth, DC.ScaleHeight)
    For i = 0 To UBound(blnFilter, 1)
    For j = 0 To UBound(blnFilter, 2)
        blnFilter(i, j) = True
    Next
    Next
    
    'Start to spiral outwards from every sample
    For r = 0.5 To Sqr((DC.ScaleWidth ^ 2) + (DC.ScaleHeight ^ 2)) Step 0.4
        'Assume we are finished
        blnCompleted = True
        For i = 0 To UBound(arrSamples)
            'Make sure the sample is active
            If blnSamples(i) Then
                'Assume the sample becomes inactive
                blnActive = False
           
                'Make a complete circle
                For t = 0 To 2 * pi Step 1 / r
                    'Calculate sampling coordinates
                    A = Sin(t) * r + arrSamples(i).X
                    b = Cos(t) * r + arrSamples(i).Y
                    'Make sure the pixel is visible
                    If A >= 0 And A <= UBound(blnFilter, 1) And b >= 0 And b <= UBound(blnFilter, 2) Then
                        'Make sure the pixel hasn't already been taken
                        If blnFilter(A, b) Then
                            'Draw the pixel and leave mark
                            'Since we drew at least one pixel, this sample is still active
                            SetPixelV idDC, A, b, arrSamples(i).C
                            blnFilter(A, b) = False
                            blnCompleted = False
                            blnActive = True
                        End If
                    End If
                Next
                
                If Not blnActive Then
                    blnSamples(i) = False
                End If
            End If
            DC.Refresh
            DoEvents
            If Main.GlobalHalt Then Exit Sub
        Next
        If blnCompleted Then Exit For
    Next
    
    DC.Refresh
    Main.BurnCurrentImage
    Exit Sub
ErrorTrap:
    MsgBox "An error occured in the algorithm:" & vbNewLine & _
            "ErrorNumber;" & Err.Number & vbNewLine & _
            "ErrorType;" & Err.Description & vbNewLine & vbNewLine & _
            "The algorithm will be stopped....", vbOKOnly Or vbCritical, _
            "Error"
    Err.Clear
End Sub

Public Sub RectangularVoronoi(ByVal DC As PictureBox)
    On Error GoTo ErrorTrap
    If SampleBound < 0 Then Exit Sub
    Main.GlobalHalt = False
    
    Dim i As Long, j As Long, pi As Double
    Dim idDC As Long
    Dim A As Long
    Dim t As Double
    Dim blnFilter() As Boolean
    Dim r As Long
    Dim blnCompleted As Boolean
    Dim blnActive As Boolean
    Dim blnSamples() As Boolean
    
    ReDim blnSamples(UBound(arrSamples))
    For i = 0 To UBound(arrSamples)
        blnSamples(i) = True
    Next
    
    pi = 4 * Atn(1)
    idDC = DC.hdc
    
    ReDim blnFilter(DC.ScaleWidth, DC.ScaleHeight)
    For i = 0 To UBound(blnFilter, 1)
    For j = 0 To UBound(blnFilter, 2)
        blnFilter(i, j) = True
    Next
    Next
    
    For r = 0 To 10000
        blnCompleted = True
        For i = 0 To UBound(arrSamples)
            If blnSamples(i) Then
                blnActive = False
              
                For A = -r To r
                    If arrSamples(i).X + A >= 0 And arrSamples(i).X + A <= UBound(blnFilter, 1) And _
                       arrSamples(i).Y - r >= 0 And arrSamples(i).Y - r <= UBound(blnFilter, 2) Then
                        If blnFilter(arrSamples(i).X + A, arrSamples(i).Y - r) Then
                            blnFilter(arrSamples(i).X + A, arrSamples(i).Y - r) = False
                            Call SetPixelV(idDC, arrSamples(i).X + A, arrSamples(i).Y - r, arrSamples(i).C)
                            blnCompleted = False
                            blnActive = True
                        End If
                    End If
                    
                    If arrSamples(i).X + A >= 0 And arrSamples(i).X + A <= UBound(blnFilter, 1) And _
                       arrSamples(i).Y + r >= 0 And arrSamples(i).Y + r <= UBound(blnFilter, 2) Then
                        If blnFilter(arrSamples(i).X + A, arrSamples(i).Y + r) Then
                            blnFilter(arrSamples(i).X + A, arrSamples(i).Y + r) = False
                            Call SetPixelV(idDC, arrSamples(i).X + A, arrSamples(i).Y + r, arrSamples(i).C)
                            blnCompleted = False
                            blnActive = True
                        End If
                    End If
                    
                    If arrSamples(i).X + r >= 0 And arrSamples(i).X + r <= UBound(blnFilter, 1) And _
                       arrSamples(i).Y + A >= 0 And arrSamples(i).Y + A <= UBound(blnFilter, 2) Then
                        If blnFilter(arrSamples(i).X + r, arrSamples(i).Y + A) Then
                            blnFilter(arrSamples(i).X + r, arrSamples(i).Y + A) = False
                            Call SetPixelV(idDC, arrSamples(i).X + r, arrSamples(i).Y + A, arrSamples(i).C)
                            blnCompleted = False
                            blnActive = True
                        End If
                    End If
                    
                    If arrSamples(i).X - r >= 0 And arrSamples(i).X - r <= UBound(blnFilter, 1) And _
                       arrSamples(i).Y + A >= 0 And arrSamples(i).Y + A <= UBound(blnFilter, 2) Then
                        If blnFilter(arrSamples(i).X - r, arrSamples(i).Y + A) Then
                            blnFilter(arrSamples(i).X - r, arrSamples(i).Y + A) = False
                            Call SetPixelV(idDC, arrSamples(i).X - r, arrSamples(i).Y + A, arrSamples(i).C)
                            blnCompleted = False
                            blnActive = True
                        End If
                    End If
                Next
            
                If Not blnActive Then
                    blnSamples(i) = False
                End If
            End If
            DC.Refresh
            DoEvents
            If Main.GlobalHalt Then Exit Sub
        Next
        If blnCompleted Then Exit For
    Next
    Main.BurnCurrentImage
    Exit Sub
ErrorTrap:
    MsgBox "An error occured in the algorithm:" & vbNewLine & _
            "ErrorNumber;" & Err.Number & vbNewLine & _
            "ErrorType;" & Err.Description & vbNewLine & vbNewLine & _
            "The algorithm will be stopped....", vbOKOnly Or vbCritical, _
            "Error"
    Err.Clear
End Sub

'Find closest sample per pixel
Public Sub ClosestSampleVoronoi(ByVal DC As PictureBox)
    On Error GoTo ErrorTrap
    If SampleBound < 0 Then Exit Sub
    Main.GlobalHalt = False
    
    Dim X As Long
    Dim Y As Long
    Dim i As Integer
    Dim arrD() As Double
    Dim ClosestIndex As Integer
    Dim idDC As Long
    
    ReDim arrD(UBound(arrSamples))
    idDC = DC.hdc
    'Loop through all pixels
    For X = 0 To DC.ScaleWidth
        For Y = 0 To DC.ScaleHeight
            For i = 0 To UBound(arrSamples)
                arrD(i) = Distance(Array(X, Y), Array(arrSamples(i).X, arrSamples(i).Y))
            Next
            ClosestIndex = iMin(arrD)
            SetPixelV idDC, X, Y, arrSamples(ClosestIndex).C
        Next
        DC.Line (X + 1, 0)-(X + 1, DC.ScaleHeight), 0
        DC.Refresh
        DoEvents
        If Main.GlobalHalt Then Exit Sub
    Next
    Main.BurnCurrentImage
    Exit Sub
    
ErrorTrap:
    MsgBox "An error occured in the algorithm:" & vbNewLine & _
            "ErrorNumber;" & Err.Number & vbNewLine & _
            "ErrorType;" & Err.Description & vbNewLine & vbNewLine & _
            "The algorithm will be stopped....", vbOKOnly Or vbCritical, _
            "Error"
    Err.Clear
End Sub

'Use smart API calls to create a voronoi diagram
Public Sub FloodFillVoronoiOutline(ByVal DC As PictureBox)
    On Error GoTo ErrorTrap
    If SampleBound < 0 Then Exit Sub
    Main.GlobalHalt = False
    
    Dim hColDC As Long, hColBmp As Long, hBmpPrev As Long
    Dim srcW As Long, srcH As Long
    
    srcW = DC.ScaleWidth
    srcH = DC.ScaleHeight
    
    hColDC = CreateCompatibleDC(DC.hdc)
    hColBmp = CreateCompatibleBitmap(DC.hdc, srcW, srcH)
    hBmpPrev = SelectObject(hColDC, hColBmp)
    DeleteObject hBmpPrev

    Dim i As Long, j As Long
    Dim nBrush As Long, cBrush As Long
    Dim nPen As Long, cPen As Long
    Dim diameter As Double

    diameter = Sqr(srcW ^ 2 + srcH ^ 2)
    nBrush = CreateSolidBrush(0)
    cBrush = SelectObject(DC.hdc, nBrush)
    Rectangle DC.hdc, -1, -1, srcW + 1, srcH + 1
    SelectObject DC.hdc, cBrush
    DeleteObject nBrush
    
    For i = 0 To SampleBound
        nBrush = CreateSolidBrush(0)
        cBrush = SelectObject(hColDC, nBrush)
        DeleteObject cBrush
        Rectangle hColDC, -1, -1, srcW + 1, srcH + 1
        
        nPen = CreatePen(0, 1, vbRed)
        cPen = SelectObject(hColDC, nPen)
        DeleteObject cPen
        For j = 0 To SampleBound
            If i <> j Then
                Math.DrawSlantLine hColDC, arrSamples(i).X, arrSamples(i).Y, arrSamples(j).X, arrSamples(j).Y, diameter
            End If
        Next
        
        nPen = CreatePen(0, 0, 0)
        cPen = SelectObject(hColDC, nPen)
        DeleteObject cPen
        nBrush = CreateSolidBrush(vbWhite)
        cBrush = SelectObject(hColDC, nBrush)
        DeleteObject cBrush
        ExtFloodFill hColDC, arrSamples(i).X, arrSamples(i).Y, 0, 1
        nPen = CreatePen(0, 1, 0)
        cPen = SelectObject(hColDC, nPen)
        DeleteObject cPen
        For j = 0 To SampleBound
            If i <> j Then
                Math.DrawSlantLine hColDC, arrSamples(i).X, arrSamples(i).Y, arrSamples(j).X, arrSamples(j).Y, diameter
            End If
        Next
        TransparentBlt DC.hdc, 0, 0, srcW, srcH, hColDC, 0, 0, srcW, srcH, 0
        DC.Refresh
        DoEvents
        If Main.GlobalHalt Then Exit Sub
    Next
    
    DeleteObject hColBmp
    DeleteDC hColDC
    Main.BurnCurrentImage
    Exit Sub
    
ErrorTrap:
    MsgBox "An error occured in the algorithm:" & vbNewLine & _
            "ErrorNumber;" & Err.Number & vbNewLine & _
            "ErrorType;" & Err.Description & vbNewLine & vbNewLine & _
            "The algorithm will be stopped....", vbOKOnly Or vbCritical, _
            "Error"
    Err.Clear
End Sub

'Okay this one is heavily commented
Public Sub FloodFillVoronoiEx(ByVal DC As PictureBox)
    On Error GoTo ErrorTrap
    'Do not run if there are no samples
    If SampleBound < 0 Then Exit Sub
    Main.GlobalHalt = False
    
    Dim hColDC As Long, hColBmp As Long, hBmpPrev As Long
    Dim srcW As Long, srcH As Long
    
    'Store the width and height of the viewport in easy accessible variables
    srcW = DC.ScaleWidth
    srcH = DC.ScaleHeight
    
    'Create a memory device context
    hColDC = CreateCompatibleDC(DC.hdc)
    'Create a memory bitmap
    hColBmp = CreateCompatibleBitmap(DC.hdc, srcW, srcH)
    'Load the memory bitmap into the memry device context
    hBmpPrev = SelectObject(hColDC, hColBmp)
    'Delete the old bitmap
    DeleteObject hBmpPrev

    Dim i As Long, j As Long
    Dim nBrush As Long, cBrush As Long
    Dim nPen As Long, cPen As Long
    Dim diameter As Double

    'Calculate the diameter of the viewport
    diameter = Sqr(srcW ^ 2 + srcH ^ 2)
    
    'For every sample...
    For i = 0 To SampleBound Step 1
        'Create a black brush (we will treat black as transparant). This is why samples cannot be completely black
        nBrush = CreateSolidBrush(0)
        'Load the black brush into the memory device context
        cBrush = SelectObject(hColDC, nBrush)
        'Delete the old brush
        DeleteObject cBrush
        'Draw a rectangle that makes the entire memory device context black (transparant)
        Rectangle hColDC, -1, -1, srcW + 1, srcH + 1
        
        'Create a pen with a red colour
        nPen = CreatePen(0, 1, vbRed)
        'Select the pen into the memory DC
        cPen = SelectObject(hColDC, nPen)
        'Delete the old pen
        DeleteObject cPen
        'For every other sample draw a slantline
        For j = 0 To i
            If i <> j Then
                Math.DrawSlantLineEx hColDC, arrSamples(i).X, arrSamples(i).Y, arrSamples(j).X, arrSamples(j).Y, diameter
            End If
        Next
        
        'Create a null pen (zero width) and load it into the memory DC
        nPen = CreatePen(0, 0, 0)
        cPen = SelectObject(hColDC, nPen)
        DeleteObject cPen
        
        'Create a brush that has the fill colour of the sample we are now using
        nBrush = CreateSolidBrush(arrSamples(i).C)
        cBrush = SelectObject(hColDC, nBrush)
        DeleteObject cBrush
        
        'Run the floodfill api so we get our voronoi cell in a non-transparaent colour
        ExtFloodFill hColDC, arrSamples(i).X, arrSamples(i).Y, 0, 1
        
        'Create a black (transparant) pen
        nPen = CreatePen(0, 1, 0)
        cPen = SelectObject(hColDC, nPen)
        DeleteObject cPen
        
        'For every other sample redraw the slant line in black (erasing the red slantlines we drew before)
        For j = 0 To i
            If i <> j Then
                Math.DrawSlantLineEx hColDC, arrSamples(i).X, arrSamples(i).Y, arrSamples(j).X, arrSamples(j).Y, diameter
            End If
        Next
        
        'Perform a transparant bit block transfer from the memory DC to the viewport picturebox
        TransparentBlt DC.hdc, 0, 0, srcW, srcH, hColDC, 0, 0, srcW, srcH, 0
        
        'Update the viewport
        DC.Refresh
        DoEvents
        If Main.GlobalHalt Then Exit Sub
    Next
    
    'Delete the memory bitmap
    DeleteObject hColBmp
    'Delete the memoryDC
    DeleteDC hColDC
    'Make the new diagram permanent
    Main.BurnCurrentImage
    Exit Sub
    
ErrorTrap:
    MsgBox "An error occured in the algorithm:" & vbNewLine & _
            "ErrorNumber;" & Err.Number & vbNewLine & _
            "ErrorType;" & Err.Description & vbNewLine & vbNewLine & _
            "The algorithm will be stopped....", vbOKOnly Or vbCritical, _
            "Error"
    Err.Clear
End Sub

'Draw a single linesegment using GDI API calls
Public Sub LineSegment(ByVal Context As Long, _
                        ByVal X1 As Double, ByVal Y1 As Double, _
                        ByVal X2 As Double, ByVal Y2 As Double)
    Dim apiPt As PointApi
    MoveToEx Context, X1, Y1, apiPt
    LineTo Context, X2, Y2
End Sub

'Calculate the square distance between two points
Public Function Distance(ByVal arrPt1 As Variant, ByVal arrPt2 As Variant) As Double
    Distance = (arrPt2(0) - arrPt1(0)) * (arrPt2(0) - arrPt1(0)) + _
               (arrPt2(1) - arrPt1(1)) * (arrPt2(1) - arrPt1(1))
End Function

'Find the lowest value in an array
Private Function iMin(Numbers() As Double) As Long
    Dim tmpMin As Double
    Dim minIndex As Long
    Dim i As Long
    
    tmpMin = Numbers(0)
    minIndex = 0
    For i = 1 To UBound(Numbers)
        If Numbers(i) < tmpMin Then
            tmpMin = Numbers(i)
            minIndex = i
        End If
    Next
    iMin = minIndex
End Function
