Type Rough
U As Double
L As Double
End Type

Sub RoughISM()

Dim x, z As Integer





x = InputBox("Enter the number of factors:")
z = InputBox("Enter the number of experts:")



Dim karar() As Double
ReDim karar(x, x, z) As Double

Dim Rkarar() As Rough
ReDim Rkarar(x, x, z) As Rough

Dim RSkarar() As Rough
ReDim RSkarar(x, x) As Rough

Dim RM() As Double
ReDim RM(x, x) As Double

Dim BRM() As Double
ReDim BRM(1 To x, 1 To x) As Double

Dim FRM() As Double
ReDim FRM(1 To x, 1 To x) As Double

Dim YRM() As Double
ReDim YRM(1 To x, 1 To x) As Double

Dim MICRM() As Double
ReDim MICRM(1 To x, 1 To x) As Double

Dim DR() As Double
ReDim DR(1 To x) As Double

Dim DE() As Double
ReDim DE(1 To x) As Double

Dim Total As Rough

For i = 1 To x * z

    For j = 1 To x
    
    
    
    If i Mod x = 0 Then
    t = i / x
    karar(x, j, t) = 0.25 * Cells(i, j)
    Else
    t = Int(i / x) + 1
    karar(i Mod x, j, t) = 0.25 * Cells(i, j)
    End If
    
   Next j
    
Next i


For t = 1 To z
For i = 1 To x
For j = 1 To x

Cells(i + x * (t - 1), j + x + 3) = Round(karar(i, j, t), 3)
Next j
Next i
Next t

For i = 1 To x
    For j = 1 To x
    
        For t = 1 To z
        up = 0
        lo = 0
        'up = 1
        'lo = 1
        sayu = 0
        sayl = 0
            For s = 1 To z
                If karar(i, j, t) <= karar(i, j, s) Then
                sayu = sayu + 1
                up = up + karar(i, j, s)
                'up = up * karar(i, j, s)
                End If
                If karar(i, j, t) >= karar(i, j, s) Then
                sayl = sayl + 1
                lo = lo + karar(i, j, s)
                'lo = lo * karar(i, j, s)
                End If
              
                Next s
        Rkarar(i, j, t).L = lo / sayl
        Rkarar(i, j, t).U = up / sayu
        
        'Rkarar(i, j, t).L = lo ^ (1 / sayl)
        'Rkarar(i, j, t).U = up ^ (1 / sayu)
        
        Next t
    
    
    
    Next j
Next i


Cells(1, 2 * x + 6) = "Experts' opinions in rough numbers form"
Cells(1, 4 * x + 10) = "Rough Decision Matrix"
Cells(1, 5 * x + 12) = "Relationship Matrix"
For i = 1 To x
For j = 1 To x

'RSkarar(i, j).U = 1
'RSkarar(i, j).L = 1

For t = 1 To z

Cells(i + 1 + x * (t - 1), j + 3 * x + 6) = "[" & CStr(Round(Rkarar(i, j, t).L, 3)) & ";" & CStr(Round(Rkarar(i, j, t).U, 3)) & "]"
If t = 1 Then
Cells(i + 1, j + 2 * x + 4) = "(" & "[" & CStr(Round(Rkarar(i, j, t).L, 2)) & ";" & CStr(Round(Rkarar(i, j, t).U, 2)) & "]" & ";"
ElseIf t > 1 And t < z Then
Cells(i + 1, j + 2 * x + 4) = Cells(i + 1, j + 2 * x + 4).Value & "[" & CStr(Round(Rkarar(i, j, t).L, 2)) & ";" & CStr(Round(Rkarar(i, j, t).U, 2)) & "]" & ";"
ElseIf t = z Then
Cells(i + 1, j + 2 * x + 4) = Cells(i + 1, j + 2 * x + 4).Value & "[" & CStr(Round(Rkarar(i, j, t).L, 2)) & ";" & CStr(Round(Rkarar(i, j, t).U, 2)) & "]" & ")"
End If

RSkarar(i, j).U = RSkarar(i, j).U + Rkarar(i, j, t).U
RSkarar(i, j).L = RSkarar(i, j).L + Rkarar(i, j, t).L

Next t
RSkarar(i, j).U = RSkarar(i, j).U / z
RSkarar(i, j).L = RSkarar(i, j).L / z


Cells(i + 1, j + 4 * x + 8) = "[" & CStr(Round(RSkarar(i, j).L, 3)) & ";" & CStr(Round(RSkarar(i, j).U, 3)) & "]"

RM(i, j) = (RSkarar(i, j).L + RSkarar(i, j).U) / 2

Cells(i + 1, j + 5 * x + 10) = Round(RM(i, j), 2)

TotalRM = TotalRM + RM(i, j)

Next j
Next i

TrH = TotalRM / (x * x)
Cells(1, 6 * x + 12) = "The threshold value "
Cells(2, 6 * x + 12) = TrH
Cells(1, 6 * x + 16) = "The initial relationship matrix  "

For i = 1 To x
For j = 1 To x

If RM(i, j) > TrH Then
BRM(i, j) = 1
Else
BRM(i, j) = 0

End If
If i = j Then
BRM(i, j) = 1
End If

Cells(i + 1, j + 6 * x + 14) = BRM(i, j)
FRM(i, j) = BRM(i, j)
Next j
Next i

For i = 1 To x

For j = 1 To x

If FRM(i, j) = 1 Then
For k = 1 To x
If FRM(j, k) = 1 Then
FRM(i, k) = 1
End If
Next k
End If



Cells(i + 1, j + 7 * x + 16) = FRM(i, j)

Next j

Next i





For i = 1 To x


For j = 1 To x





Cells(i + 1, j + 7 * x + 16) = FRM(i, j)

If BRM(i, j) <> FRM(i, j) Then
Cells(i + 1, j + 8 * x + 18) = FRM(i, j) & "*"
Else
Cells(i + 1, j + 8 * x + 18) = FRM(i, j)
End If

YRM(i, j) = FRM(i, j)

MICRM(i, j) = FRM(i, j)

Next j

Next i




lev = 0

100: lev = lev + 1

For i = 1 To x
For j = 1 To x
YRM(i, j) = FRM(i, j)
Next j
Next i




For i = 1 To x
Cells((lev - 1) * x + i + 1, 9 * x + 20) = i
sayx = 0
sayy = 0
sayxy = 0

For j = 1 To x

If YRM(i, j) = 1 Then

sayx = sayx + 1

If sayx = 1 Then
Cells((lev - 1) * x + i + 1, 9 * x + 21) = j
ElseIf sayx > 1 Then

Cells((lev - 1) * x + i + 1, 9 * x + 21) = Cells((lev - 1) * x + i + 1, 9 * x + 21).Value & ";" & j

End If

End If

If YRM(j, i) = 1 Then
sayy = sayy + 1

If sayy = 1 Then
Cells((lev - 1) * x + i + 1, 9 * x + 22) = j
ElseIf sayy > 1 Then

Cells((lev - 1) * x + i + 1, 9 * x + 22) = Cells((lev - 1) * x + i + 1, 9 * x + 22).Value & ";" & j

End If

End If


If YRM(i, j) = 1 And YRM(j, i) = 1 Then
sayxy = sayxy + 1

If sayxy = 1 Then
Cells((lev - 1) * x + i + 1, 9 * x + 23) = j
ElseIf sayxy > 1 Then

Cells((lev - 1) * x + i + 1, 9 * x + 23) = Cells((lev - 1) * x + i + 1, 9 * x + 23).Value & ";" & j

End If

End If



Next j

If Cells((lev - 1) * x + i + 1, 9 * x + 21) <> 0 And Cells((lev - 1) * x + i + 1, 9 * x + 21).Value = Cells((lev - 1) * x + i + 1, 9 * x + 23).Value Then
Cells((lev - 1) * x + i + 1, 9 * x + 24) = lev

For s = 1 To x
FRM(i, s) = 0
FRM(s, i) = 0
Next s
End If



Next i

sayson = 0
For i = 1 To x


For j = 1 To x



If YRM(i, j) > 0 Then
sayson = sayson + 1
End If

Next j

Next i

If sayson > 0 Then
GoTo 100
End If




For i = 1 To x


For j = 1 To x
DR(i) = MICRM(i, j) + DR(i)
DE(i) = MICRM(j, i) + DE(i)




Next j
Cells(i + 1, 9 * x + 27) = DR(i)
Cells(i + 1, 9 * x + 26) = DE(i)
'Cells(i + 1, 9 * x + 26) = "(" & DE(i) & "," & DR(i) & ")"
Next i





Cells(1, 9 * x + 20) = "Element (Pi)"

Cells(1, 9 * x + 21) = "Reachability set: R (Pi)"
Cells(1, 9 * x + 22) = "Antecedent set: A (Pi)"
Cells(1, 9 * x + 23) = "Intersection R (Pi)n A (Pi)"
Cells(1, 9 * x + 24) = "Level"

Cells(1, 7 * x + 18) = "The final relationship matrix  "

Cells(1, 9 * x + 27) = "The driving power "
Cells(1, 9 * x + 26) = "The dependence power"







End Sub
