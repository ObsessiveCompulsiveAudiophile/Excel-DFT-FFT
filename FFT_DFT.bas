Attribute VB_Name = "Module1"
'Serkan Guer 2023
'All rights reserved
Option Explicit
Type Complex
   Re As Double
   Im As Double
End Type
'Radix 2 recursive
Sub F(N As Long, s As Long, q As Long, d As Long, x() As Complex)

   Dim m As Long, p As Long, theta0 As Double
   Dim wp As Complex, a As Complex, b As Complex
   
   m = N / 2
   theta0 = 2 * Application.pi / N
   
   If N > 1 Then
      For p = 0 To m - 1
         wp.Re = Cos(p * theta0)
         wp.Im = -Sin(p * theta0)
         
         a = x(q + p)
         b = x(q + p + m)
         
         x(q + p).Re = a.Re + b.Re
         x(q + p).Im = a.Im + b.Im
         
         x(q + p + m).Re = (a.Re - b.Re) * wp.Re - (a.Im - b.Im) * wp.Im
         x(q + p + m).Im = (a.Re - b.Re) * wp.Im + (a.Im - b.Im) * wp.Re
      Next p
      
      Call F(N / 2, 2 * s, q, d, x)
      Call F(N / 2, 2 * s, q + m, d + s, x)
   
   ElseIf q > d Then
      Call Swap(x(q), x(d))
   End If
   
End Sub

Sub Swap(a As Complex, b As Complex)
   Dim temp As Complex
   temp = a
   a = b
   b = temp
End Sub

Sub fft(N As Long, x() As Complex)
   Call F(N, 1, 0, 0, x)
End Sub

Sub ifft(N As Long, x() As Complex)
   Dim k As Long
   For k = 0 To N - 1
      x(k).Im = -x(k).Im
   Next k
   
   Call F(N, 1, 0, 0, x)
   
   For k = 0 To N - 1
      x(k).Re = x(k).Re / N
      x(k).Im = -x(k).Im / N
   Next k
End Sub

Sub GenerateFR() 'fast fourier transform
    Columns("B:H").Select
    Selection.ClearContents
    Range("A1").Select
   
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim N As Long, i As Long, rw As Long, SampleRate As Long, rng As String
    
    SampleRate = 48000 'Change to correct sample rate
    
    rw = Range("A1").End(xlDown).Row
    Dim x(1048575) As Complex
    
    N = rw
    
    For i = 0 To N - 1
       x(i).Re = Cells(i + 1, 1).Value
       x(i).Im = 0
    Next i
    
    Call fft(N, x())
    
    For i = 0 To N - 1
       'Cells(i + 1, 2).Value = WorksheetFunction.Complex(x(i).Re, x(i).Im)
       Cells(i + 1, 2).Value = x(i).Re
       Cells(i + 1, 3).Value = x(i).Im
    Next i
    
    'Extract magnitude and phase
    Cells(1, 7) = rw
    Cells(1, 8) = SampleRate
    Cells(2, 4).Formula = "=sqrt(B2^2+c2^2)"
    Cells(2, 5).Formula = "=atan2(b2,c2)"
    Cells(2, 6).Formula = "=F1+$H$1/$G$1"
    Cells(2, 7).Formula = "=20*LOG10(D2)+100" 'Change 100 to the preferred SPL dB Offset value
    Cells(2, 8).Formula = "=-180*E2/PI()"
     
    rng = "D2:H" & rw / 2
    Range("D2:H2").Select
    Selection.AutoFill Destination:=Range(rng)
    Range("G2").Select
   
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub GenerateImpulse() 'inverse fast fourier transform
    Columns("D:H").Select
    Selection.ClearContents
    Columns("A:A").Select
    Selection.ClearContents
    Range("B1").Select
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Dim N As Long, i As Long, rw As Long
    
    rw = Range("B1").End(xlDown).Row
    Dim x(1048575) As Complex
    
    N = rw
    
    For i = 0 To N - 1
        x(i).Re = Cells(i + 1, 2).Value
        x(i).Im = Cells(i + 1, 3).Value
    Next i
    'Inverse fft
    Call ifft(N, x)
    
    'Write ifft result to column D
    For i = 0 To N - 1
        Cells(i + 1, 4) = x(i).Re
    Next i
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Sub dft() 'basic discrete fourier transform
Dim r As Long, s As Long, k As LongLong, N(1048576) As Double, SumReal As Double, SumImag As Double, pi As Double, w As Double
Columns("B:H").Select
    Selection.ClearContents
    Range("B1").Select
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
pi = Application.pi

For r = 1 To 1048576
    N(r - 1) = Cells(r, 1).Value
    If Cells(r, 1) = "" Then Exit For
Next r
r = r - 1

For k = 0 To (r / 2) - 1
    SumReal = 0: SumImag = 0
    For s = 0 To r - 1
        w = -2 * pi * k * s / r
        SumReal = SumReal + N(s) * Cos(w)
        SumImag = SumImag + N(s) * Sin(w)
    Next s
    Cells(k + 1, 2) = SumReal: Cells(k + 1, 3) = SumImag
Next k

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub
