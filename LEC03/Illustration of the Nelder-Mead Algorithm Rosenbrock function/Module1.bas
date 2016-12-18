Attribute VB_Name = "Module1"
Option Base 1
Function Rosenbrock(params As Variant, A As Double, B As Double) As Variant

' Rosenbrock banana function.
' Usual values are A=100 and B=1 with the minimum at (1,1), where the value is 0.

Dim x As Double, y As Double
x = params(1)
y = params(2)

Rosenbrock = A * (y - x ^ 2) ^ 2 + (B - x) ^ 2

End Function
Function NelderMead(fname As String, StartValues As Variant, Tolerance As Double, Shock As Double, MaxIters As Integer, _
    A As Double, B As Double) As Variant

' Nelder Mead Algorithm
' From "Option Pricing Models and Volatility Using Excel-VBA"
' By Rouah, F., and G. Vainberg, John Wiley & Sons (2007)
' Algorithm from "Numerical Analysis" (1995), Griffiths and Watson, page 194.
' INPUTS
'   fname = name of the objective function
'   StartValues = a vector of starting values
'   Tolerance = Tolerance on objective function value (stops when value < tolerance)
'   Shock = shock to create starting vertices (start = start + start*RandomShock%
'   MaxIter = Maximum number of iterations
' OUTPUTS
'   Vector of parameter estimates
'   Value of objective function
'   Actual number of iterations used

' Required settings
Dim rho As Double, xi As Double, gam As Double, sigma As Double
rho = 1
xi = 2
gam = 0.5
sigma = 0.5
    
' Number of parameters
N = Application.count(StartValues)
    
' Initialize the results matrix and the points representing the
' best (x1), second-to-worst (xn), worst (xn1), mean (xbar), reflective (xr), expansion (xe), outside contraction (xoc) and inside contraction (xic)
Dim x1() As Double, xn() As Double, xn1() As Double, xbar() As Double, xr() As Double, xe() As Double, xoc() As Double, xic() As Double
Dim passParams() As Double
ReDim resmat(N + 1, N + 1) As Double
ReDim x1(N) As Double, xn(N) As Double, xn1(N) As Double, xbar(N) As Double, xr(N) As Double, xe(N) As Double, xoc(N) As Double, xic(N) As Double
ReDim passParams(N)
    
For i = 1 To N
    resmat(1, i + 1) = StartValues(i)
Next i
resmat(1, 1) = Run(fname, StartValues, A, B)
    
' Randomize and initialize the starting values
For j = 1 To N
    For i = 1 To N
        Randomize
        random = 2 * Shock * Rnd - Shock
        resmat(j + 1, i + 1) = (1 + random) * StartValues(i)
        passParams(i) = resmat(j + 1, i + 1)
    Next i
resmat(j + 1, 1) = Run(fname, passParams, A, B)
Next j
    
For iter = 1 To MaxIters
    ' Sort the functional values
    resmat = BubSortRows(resmat)
    If (Abs(resmat(1, 1) - resmat(N + 1, 1)) < Tolerance) Then
        Exit For
    End If
        
    ' The best functional value and point
    f1 = resmat(1, 1)
    For i = 1 To N
        x1(i) = resmat(1, i + 1)
    Next i
                
    ' The second-to-worst functional value and point
    fn = resmat(N, 1)
    For i = 1 To N
        xn(i) = resmat(N, i + 1)
    Next i
        
    ' The worst functional value and point
    fn1 = resmat(N + 1, 1)
    For i = 1 To N
        xn1(i) = resmat(N + 1, i + 1)
    Next i
   
    ' The center of gravity
    For i = 1 To N
        xbar(i) = 0
        For j = 1 To N
            xbar(i) = xbar(i) + resmat(j, i + 1)
        Next j
        xbar(i) = xbar(i) / N
    Next i
        
    ' Reflection point
    For i = 1 To N
        xr(i) = xbar(i) + rho * (xbar(i) - xn1(i))
    Next i
    fr = Run(fname, xr, A, B)
        
    shrink = 0
    If ((fr >= f1) And (fr < fn)) Then
        newpoint = xr
        newf = fr
    ElseIf (fr < f1) Then
        ' Expansion point
        For i = 1 To N
            xe(i) = xbar(i) + xi * (xr(i) - xbar(i))
        Next i
        fe = Run(fname, xe, A, B)
        If (fe < fr) Then
            newpoint = xe
            newf = fe
        Else
            newpoint = xr
            newf = fr
        End If
    ElseIf (fr >= fn) Then
        If ((fr >= fn) And (fr < fn1)) Then
            ' Outside contraction
            For i = 1 To N
                xoc(i) = xbar(i) + gam * (xr(i) - xbar(i))
            Next i
        foc = Run(fname, xoc, A, B)
        If (foc <= fr) Then
            newpoint = xoc
            newf = foc
        Else
            shrink = 1
        End If
    Else
        ' Inside contraction
        For i = 1 To N
            xic(i) = xbar(i) - gam * (xbar(i) - xn1(i))
        Next i
        fic = Run(fname, xic, A, B)
        If (fic < fn1) Then
            newpoint = xic
            newf = fic
        Else
            shrink = 1
        End If
            End If
    End If
    If (shrink = 1) Then
        For scnt = 2 To N + 1
            For i = 1 To N
                ' Shrinkage step
                resmat(scnt, i + 1) = x1(i) + sigma * (resmat(scnt, i + 1) - x1(1))
                passParams(i) = resmat(scnt, i + 1)
            Next i
            resmat(scnt, 1) = Run(fname, passParams, A, B)
        Next scnt
    Else
        For i = 1 To N
            resmat(N + 1, i + 1) = newpoint(i)
        Next i
        resmat(N + 1, 1) = newf
    End If
Next iter

' Sort the results
resmat = BubSortRows(resmat)

' Output the parameter estimates
ReDim output(N + 2) As Double
For i = 1 To N
    output(i) = resmat(1, i + 1)
Next i
    
' Output the value of the objective function
output(N + 1) = resmat(1, 1)
    
' Output the actual number of iterations used
output(N + 2) = iter
    
NelderMead = Application.Transpose(output)
    
End Function
Function BubSortRows(passVec)
    
' Bubble Sort Algorithm -- Extension for Row Sorting
' From "Option Pricing Models and Volatility Using Excel-VBA"
' By Rouah, F., and G. Vainberg, John Wiley & Sons (2007)
    
Dim tmpVec() As Double, temp() As Double
uVec = passVec
rownum = UBound(uVec, 1)
colnum = UBound(uVec, 2)
ReDim tmpVec(rownum, colnum) As Double
ReDim temp(colnum) As Double
    
For i = rownum - 1 To 1 Step -1
    For j = 1 To i
        If (uVec(j, 1) > uVec(j + 1, 1)) Then
            For k = 1 To colnum
                temp(k) = uVec(j + 1, k)
                uVec(j + 1, k) = uVec(j, k)
                uVec(j, k) = temp(k)
            Next k
        End If
    Next j
Next i

BubSortRows = uVec

End Function


