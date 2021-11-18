Attribute VB_Name = "DemoLibMemory"
Option Explicit
Option Private Module

Private Const LOOPS As Long = 10000

Sub DemoMain()
    Debug.Print String(24, "-") & " Speed " & String(24, "-")
    DemoMemByteSpeed
    DemoMemIntSpeed
    DemoMemLongSpeed
    DemoMemLongPtrSpeed
    DemoMemObjectSpeed
    Debug.Print String(21, "-") & " Redirection " & String(21, "-")
    DemoInstanceRedirection
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
