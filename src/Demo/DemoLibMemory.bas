Attribute VB_Name = "DemoLibMemory"
Option Explicit
Option Private Module

#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
    #Else
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
    #End If
#Else 'Windows
    'https://msdn.microsoft.com/en-us/library/mt723419(v=vs.85).aspx
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    #End If
#End If

Private Const LOOPS As Long = 1000

Sub DemoMain()
    Debug.Print String(24, "-") & " Speed " & String(24, "-")
    DemoMemByteSpeed
    DemoMemIntSpeed
    DemoMemLongSpeed
    DemoMemLongPtrSpeed
    DemoMemObjectSpeed
    Debug.Print String(21, "-") & " Redirection " & String(21, "-")
    DemoInstanceRedirection
    Debug.Print String(23, "-") & " MemCopy " & String(23, "-")
    DemoMemCopySpeed
End Sub

Private Sub DemoInstanceRedirection()
    Const loopsCount As Long = 100000
    Dim i As Long
    Dim t As Double
    '
    Debug.Print Format$(loopsCount, "#,##0") & " times"
    t = Timer
    For i = 1 To loopsCount
        Debug.Assert DemoClass.Factory2(i).ID = i
    Next i
    Debug.Print "Public  Init (seconds): " & VBA.Round(Timer - t, 3)
    '
    t = Timer
    For i = 1 To loopsCount
        Debug.Assert DemoClass.Factory(i).ID = i
    Next i
    Debug.Print "Private Init (seconds): " & VBA.Round(Timer - t, 3)
End Sub

Private Sub DemoMemByteSpeed()
    Dim x1 As Byte: x1 = 1
    Dim x2 As Byte: x2 = 2
    Dim i As Long
    Dim t As Double
    '
    t = Timer
    For i = 1 To LOOPS
        MemByte(VarPtr(x1)) = MemByte(VarPtr(x2))
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By Ref " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
    Debug.Assert x1 = x2
    '
    t = Timer
    For i = 1 To LOOPS
        CopyMemory x1, x2, 1
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By API " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
    DoEvents
End Sub

Private Sub DemoMemIntSpeed()
    Dim x1 As Integer: x1 = 11111
    Dim x2 As Integer: x2 = 22222
    Dim i As Long
    Dim t As Double
    '
    t = Timer
    For i = 1 To LOOPS
        MemInt(VarPtr(x1)) = MemInt(VarPtr(x2))
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By Ref " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
    Debug.Assert x1 = x2
    '
    t = Timer
    For i = 1 To LOOPS
        CopyMemory x1, x2, 2
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By API " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
    DoEvents
End Sub

Private Sub DemoMemLongSpeed()
    Dim x1 As Long: x1 = 111111111
    Dim x2 As Long: x2 = 222222222
    Dim i As Long
    Dim t As Double
    '
    t = Timer
    For i = 1 To LOOPS
        MemLong(VarPtr(x1)) = MemLong(VarPtr(x2))
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By Ref " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
    Debug.Assert x1 = x2
    '
    t = Timer
    For i = 1 To LOOPS
        CopyMemory x1, x2, 4
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By API " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
    DoEvents
End Sub

Private Sub DemoMemLongPtrSpeed()
    #If Win64 Then
        Dim x1 As LongLong: x1 = 111111111111111^
        Dim x2 As LongLong: x2 = 111111111111112^
    #Else
        Dim x1 As Long: x1 = 111111111
        Dim x2 As Long: x2 = 222222222
    #End If
    Dim i As Long
    Dim t As Double
    '
    t = Timer
    For i = 1 To LOOPS
        MemLongPtr(VarPtr(x1)) = MemLongPtr(VarPtr(x2))
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By Ref " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
    Debug.Assert x1 = x2
    '
    t = Timer
    For i = 1 To LOOPS
        CopyMemory x1, x2, PTR_SIZE
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By API " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
    DoEvents
End Sub

Private Sub DemoMemObjectSpeed()
    Dim i As Long
    Dim t As Double
    Dim d As DemoClass: Set d = New DemoClass
    Dim obj As Object
    #If Win64 Then
        Dim ptr As LongLong
    #Else
        Dim ptr As Long
    #End If
    '
    ptr = ObjPtr(d)
    t = Timer
    For i = 1 To LOOPS
        Set obj = MemObj(ptr)
    Next i
    Debug.Print "Dereferenced an Object " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
    DoEvents
End Sub

Private Sub DemoMemCopySpeed()
    Dim t As Double
    Dim a1() As Byte
    Dim a2() As Byte
    Dim size As Long
    Dim iterations As Long
    Dim i As Long
    Dim src As LongPtr
    Dim dest As LongPtr
    Dim res1 As Double
    Dim res2 As Double
    Dim slowFactor As Long
    '
    size = 2
    iterations = 2 ^ 21
    Debug.Print "Size", "Iterations", "MemCopy", "CopyMemory"
    Do
        ReDim a1(0 To size - 1)
        ReDim a2(0 To size - 1)
        '
        For i = 3 To UBound(a2) - 2 ^ 16 - 1 Step 2 ^ 16
            a2(i) = 128 'This forces BSTR copy rather than SAFEARRAY
        Next i
        '
        src = VarPtr(a2(0))
        dest = VarPtr(a1(0))
        '
        t = Timer
        For i = 1 To iterations
            MemCopy dest, src, size
        Next i
        res1 = Round(Timer - t, 3)
        '
        slowFactor = 10000 'In case API call is too slow
        Do
            t = Timer
            For i = 1 To iterations \ slowFactor
                CopyMemory ByVal dest, ByVal src, size
            Next i
            res2 = Round(Timer - t, 3)
            If res2 < 0.1 Then
                slowFactor = slowFactor / 10
            Else
                Exit Do
            End If
        Loop Until slowFactor = 1
        '
        Debug.Print size, iterations, res1, res2 * slowFactor _
                  & IIf(slowFactor > 1, " (extrapolated from " & iterations _
                  \ slowFactor & " iterations that took " & res2 & ")", "")
        '
        Const maxLong As Long = 2147483647
        If CDbl(size) * 2 > maxLong Then
            If CDbl(size) * 2 - 1 > maxLong Then
                iterations = 2
            Else
                size = CDbl(size) * 2 - 1
            End If
        Else
            size = size * 2
        End If
        iterations = iterations / 1.6
        DoEvents
    Loop Until iterations = 1
End Sub
