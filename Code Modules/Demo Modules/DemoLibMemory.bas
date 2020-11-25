Attribute VB_Name = "DemoLibMemory"
Option Explicit
Option Private Module

Private Const LOOPS As Long = 100000

Sub DemoMain()
    DemoMemByte
    DemoMemInt
    DemoMemLong
    DemoMemLongPtr
End Sub

Sub DemoMemByte()
    Dim x1 As Byte: x1 = 1
    Dim x2 As Byte: x2 = 2
    Dim i As Long
    Dim t As Double
    
    t = Timer
    For i = 1 To LOOPS
        MemByte(VarPtr(x1)) = MemByte(VarPtr(x2))
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By Ref " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
    Debug.Assert x1 = x2

    t = Timer
    For i = 1 To LOOPS
        CopyMemory x1, x2, 1
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By API " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
End Sub

Sub DemoMemInt()
    Dim x1 As Integer: x1 = 11111
    Dim x2 As Integer: x2 = 22222
    Dim i As Long
    Dim t As Double
    
    t = Timer
    For i = 1 To LOOPS
        MemInt(VarPtr(x1)) = MemInt(VarPtr(x2))
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By Ref " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
    Debug.Assert x1 = x2

    t = Timer
    For i = 1 To LOOPS
        CopyMemory x1, x2, 1
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By API " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
End Sub

Sub DemoMemLong()
    Dim x1 As Long: x1 = 111111111
    Dim x2 As Long: x2 = 222222222
    Dim i As Long
    Dim t As Double
    
    t = Timer
    For i = 1 To LOOPS
        MemLong(VarPtr(x1)) = MemLong(VarPtr(x2))
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By Ref " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
    Debug.Assert x1 = x2

    t = Timer
    For i = 1 To LOOPS
        CopyMemory x1, x2, 1
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By API " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
End Sub

Sub DemoMemLongPtr()
    #If Win64 Then
        Dim x1 As LongLong: x1 = 111111111111111^
        Dim x2 As LongLong: x2 = 111111111111112^
    #Else
        Dim x1 As Long: x1 = 111111111
        Dim x2 As Long: x2 = 222222222
    #End If
    Dim i As Long
    Dim t As Double
    
    t = Timer
    For i = 1 To LOOPS
        MemLongPtr(VarPtr(x1)) = MemLongPtr(VarPtr(x2))
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By Ref " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
    Debug.Assert x1 = x2

    t = Timer
    For i = 1 To LOOPS
        CopyMemory x1, x2, 1
    Next i
    Debug.Print "Copy <" & TypeName(x1) & "> By API " & Format$(LOOPS, "#,##0") _
        & " times in " & Round(Timer - t, 3) & " seconds"
End Sub
