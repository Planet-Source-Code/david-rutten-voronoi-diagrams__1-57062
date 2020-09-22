Attribute VB_Name = "Math"
Option Explicit
'A whole bunch of API GDI functions
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As PointApi) As Long
Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
                                     ByVal nWidth As Long, ByVal nHeight As Long, _
                                     ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
                                     ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
                                        ByVal nWidth As Long, ByVal nHeight As Long, _
                                        ByVal hSrcDC As Long, _
                                        ByVal xSrc As Long, ByVal ySrc As Long, _
                                        ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
                                        ByVal dwRop As Long) As Long
Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
                                                   ByVal nWidth As Long, ByVal nHeight As Long, _
                                                   ByVal hSrcDC As Long, _
                                                   ByVal xSrc As Long, ByVal ySrc As Long, _
                                                   ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
                                                   ByVal crTransparent As Long) As Boolean
Declare Function ShellExecuteA Lib "shell32.dll" (ByVal hwnd As Long, ByVal lpOperation As String, _
                                                  ByVal lpFile As String, ByVal lpParameters As String, _
                                                  ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'The typewe use to store samples
Public Type Sample
    X As Long
    Y As Long
    C As Long
End Type

'Used in several API GDI calls
Public Type PointApi
    X As Long
    Y As Long
End Type

'Treat sample collection as particle cloud.
Public Function DisperseSamples(Optional ByVal Radius As Long = 100, _
                                Optional ByVal maxRegion As Long = 20, _
                                Optional ByVal Demp As Double = 1) As Boolean
    DisperseSamples = False
    If SampleBound < 0 Then Exit Function

    Dim i As Long
    Dim j As Long
    Dim mapHeight As Long
    Dim mapWidth As Long
    
    Dim arrMirror(3) As Variant
    Dim arrX() As Double
    Dim arrY() As Double
    
    ReDim arrX(SampleBound)
    ReDim arrY(SampleBound)
        
    Dim d As Double
    Dim F As Double

    mapWidth = Main.picDiagram.ScaleWidth
    mapHeight = Main.picDiagram.ScaleHeight
    Radius = Radius ^ 2

    For i = 0 To SampleBound
        arrX(i) = 0#
        arrY(i) = 0#
    Next

    For i = 0 To SampleBound
        For j = 0 To SampleBound
            If j <> i Then
                d = Distance(Array(arrSamples(i).X, arrSamples(i).Y), _
                             Array(arrSamples(j).X, arrSamples(j).Y))
                
                If d < Radius Then
                    If d = 0 Then
                        arrX(i) = arrX(i) + (2 * Rnd - 0.5)
                        arrY(i) = arrY(i) + (2 * Rnd - 0.5)
                    Else
                        F = (1 / d) * Demp
                        arrX(i) = arrX(i) + (F * (arrSamples(i).X - arrSamples(j).X))
                        arrY(i) = arrY(i) + (F * (arrSamples(i).Y - arrSamples(j).Y))
                    End If
                End If
            End If
        Next
        
        If Demp > 0 Then
            arrMirror(0) = Array(arrSamples(i).X, -arrSamples(i).Y - 1)
            arrMirror(1) = Array(2 * mapWidth - arrSamples(i).X, arrSamples(i).Y)
            arrMirror(2) = Array(arrSamples(i).X, 2 * mapHeight + 1 - arrSamples(i).Y)
            arrMirror(3) = Array(-arrSamples(i).X - 1, arrSamples(i).Y)
            For j = 0 To 3
                d = Distance(Array(arrSamples(i).X, arrSamples(i).Y), Array(arrMirror(j)(0), arrMirror(j)(1)))
                If d < Radius Then
                    F = (10 / d) * Demp
                    arrX(i) = arrX(i) + (F * (arrSamples(i).X - arrMirror(j)(0)))
                    arrY(i) = arrY(i) + (F * (arrSamples(i).Y - arrMirror(j)(1)))
                End If
            Next
        End If
    Next

    For i = 0 To SampleBound
        arrSamples(i).X = Main.Region(arrSamples(i).X + Main.Region(arrX(i), -maxRegion, maxRegion), 0, mapWidth - 1)
        arrSamples(i).Y = Main.Region(arrSamples(i).Y + Main.Region(arrY(i), -maxRegion, maxRegion), 0, mapHeight - 1)
    Next
    
    DisperseSamples = True
End Function

'Draws a line halfway between two samples at 90 degrees
Public Function DrawSlantLine(ByVal picDest As Long, _
                              ByVal X1 As Long, ByVal Y1 As Long, _
                              ByVal X2 As Long, ByVal Y2 As Long, _
                              ByVal diameter As Long) As Boolean
    DrawSlantLine = False
    If X1 = X2 And Y1 = Y2 Then Exit Function
    Dim Xm As Long, Ym As Long
    Dim Xt As Long, Yt As Long
    Dim d As Double
    
    d = Sqr((Y2 - Y1) ^ 2 + (X2 - X1) ^ 2)
    Xm = (X1 + X2) / 2
    Ym = (Y1 + Y2) / 2
    Xt = Xm - (Y2 - Y1)
    Yt = Ym + (X2 - X1)
    Xt = Xm + (Xt - Xm) / d * diameter
    Yt = Ym + (Yt - Ym) / d * diameter
    Xm = Xt + (Xm - Xt) / d * diameter
    Ym = Yt + (Ym - Yt) / d * diameter
    Voronoi.LineSegment picDest, Xm, Ym, Xt, Yt
    DrawSlantLine = True
End Function

'This function is identical to the one above with the exception that it offsets all lines by one pixel
'awayfrom the X1,Y1 coordinate. ie. it removes the voronoi cell outlines you get when using the above function.
Public Function DrawSlantLineEx(ByVal picDest As Long, _
                                ByVal X1 As Long, ByVal Y1 As Long, _
                                ByVal X2 As Long, ByVal Y2 As Long, _
                                ByVal diameter As Long) As Boolean
    DrawSlantLineEx = False
    If X1 = X2 And Y1 = Y2 Then Exit Function
    Dim Xm As Long, Ym As Long
    Dim Xt As Long, Yt As Long
    Dim d As Double
    
    d = Sqr((Y2 - Y1) ^ 2 + (X2 - X1) ^ 2)

    Xm = (X1 + X2) / 2 + Sgn(X2 - X1)
    Ym = (Y1 + Y2) / 2 + Sgn(Y2 - Y1)
    Xt = Xm - (Y2 - Y1)
    Yt = Ym + (X2 - X1)
    Xt = Xm + (Xt - Xm) / d * diameter
    Yt = Ym + (Yt - Ym) / d * diameter
    Xm = Xt + (Xm - Xt) / d * diameter
    Ym = Yt + (Ym - Yt) / d * diameter
    Voronoi.LineSegment picDest, Xm, Ym, Xt, Yt
    DrawSlantLineEx = True
End Function

Public Function LineLine_INT_Absolute(ByVal X1 As Double, ByVal Y1 As Double, _
                                      ByVal X2 As Double, ByVal Y2 As Double, _
                                      ByVal X3 As Double, ByVal Y3 As Double, _
                                      ByVal X4 As Double, ByVal Y4 As Double, _
                                      Optional ByVal ClipSegment As Boolean = False) As Double()
    Dim t1 As Double, t2 As Double
    Dim C(1) As Double
    
    t1 = LineLine_INT_Parameter(X1, Y1, X2, Y2, X3, Y3, X4, Y4)
    If IsNull(t1) Then Exit Function
    
    If ClipSegment Then
        t2 = LineLine_INT_Parameter(X3, Y3, X4, Y4, X1, Y1, X2, Y2)
        If t1 < 0 Or t1 > 1 Or t2 < 0 Or t2 > 1 Then Exit Function
    End If
    
    C(0) = X1 + t1 * (X2 - X1)
    C(1) = Y1 + t1 * (Y2 - Y1)
    LineLine_INT_Absolute = C
End Function

Public Function LineLine_INT_Parameter(ByVal X1 As Double, ByVal Y1 As Double, _
                                       ByVal X2 As Double, ByVal Y2 As Double, _
                                       ByVal X3 As Double, ByVal Y3 As Double, _
                                       ByVal X4 As Double, ByVal Y4 As Double) As Double
    LineLine_INT_Parameter = -1
    Dim Enumerator As Double
    Dim Denominator As Double
    
    Denominator = (Y4 - Y3) * (X2 - X1) - (X4 - X3) * (Y2 - Y1)
    'Parallel lines or singular line(s)
    If Denominator = 0 Then Exit Function
    
    Enumerator = (X4 - X3) * (Y1 - Y3) - (Y4 - Y3) * (X1 - X3)
    LineLine_INT_Parameter = Enumerator / Denominator
End Function

'Write the sample array to a file
Public Function SaveSolution(ByVal strFilePath As String) As Boolean
    SaveSolution = False
    On Error GoTo ErrorTrap
    If SampleBound < 0 Then Exit Function
    
    Dim i As Long
    
    Open strFilePath For Output As #1
    Write #1, "Voronoi Application DataStream"
    Write #1, "File created on " & CStr(Now)
    Write #1,
    
    For i = 0 To SampleBound
        Write #1, arrSamples(i).X & ";" & arrSamples(i).Y & ";" & arrSamples(i).C
    Next
    
    Write #1,
    Write #1, "_eof"
    Close #1   ' Close file.

    SaveSolution = True
    Exit Function
ErrorTrap:
    SaveSolution = False
End Function

'Load samples from a file and ADD these to the sample array
Public Function LoadSolution(ByVal strFilePath As String) As Long
    LoadSolution = -1
    
    Dim datLine As String
    Dim N As Long
    Dim arrParts As Variant
    Dim X As Long, Y As Long, C As Long
    
    N = 0
    Open strFilePath For Input As #1
    Do While Not EOF(1)
        Line Input #1, datLine
        datLine = Replace(datLine, Chr(34), "")
        arrParts = Split(datLine, ";")
        If IsArray(arrParts) Then
            If UBound(arrParts) = 2 Then
                X = Abs(CLng(Val(arrParts(0))))
                Y = Abs(CLng(Val(arrParts(1))))
                C = Abs(CLng(Val(arrParts(2))))
                
                ReDim Preserve arrSamples(SampleBound + 1)
                arrSamples(SampleBound + 1).X = X
                arrSamples(SampleBound + 1).Y = Y
                arrSamples(SampleBound + 1).C = C
                SampleBound = SampleBound + 1
                N = N + 1
            End If
        End If
    Loop
    Close #1
    Main.DrawSamples True
    LoadSolution = N
    
ErrorTrap:
    LoadSolution = -1
End Function
