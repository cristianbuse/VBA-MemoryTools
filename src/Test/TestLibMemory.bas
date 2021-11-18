Attribute VB_Name = "TestLibMemory"
'''=============================================================================
''' VBA MemoryTools
''' -----------------------------------------------
''' https://github.com/cristianbuse/VBA-MemoryTools
''' -----------------------------------------------
''' MIT License
'''
''' Copyright (c) 2020 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

Option Explicit
Option Private Module

Public Sub RunAllTests()
    TestReadByte
    TestWriteByte
    TestReadInteger
    TestWriteInteger
    TestReadLong
    TestWriteLong
    TestReadLongLong
    TestWriteLongLong
    '
    Debug.Print "Finished running tests at " & Now()
End Sub

Private Sub TestReadByte()
    Dim b As Byte
    Dim i As Integer
    Dim l1 As Long, l2 As Long, l3 As Long, l4 As Long
    Dim l As Long
    Dim s As String
    '
    For l1 = 0 To 255
        b = l1
        Debug.Assert MemByte(VarPtr(b)) = b
    Next l1
    '
    For l1 = 0 To 255
        For l2 = 0 To 255
            i = ((l1 Xor &H8000) + l2 * 256) Xor &H8000
            Debug.Assert MemByte(VarPtr(i)) = l1
            Debug.Assert MemByte(VarPtr(i) + 1) = l2
        Next l2
    Next l1
    '
    For l1 = 0 To 255 Step 25
        For l2 = 0 To 255 Step 25
            For l3 = 0 To 255 Step 5
                For l4 = 0 To 255 Step 5
                    l = (((l1 + l2 * 256 + l3 * 256 ^ 2) Xor &H80000000) + l4 * 256 ^ 3) Xor &H80000000
                    Debug.Assert MemByte(VarPtr(l)) = l1
                    Debug.Assert MemByte(VarPtr(l) + 1) = l2
                    Debug.Assert MemByte(VarPtr(l) + 2) = l3
                    Debug.Assert MemByte(VarPtr(l) + 3) = l4
                Next l4
            Next l3
        Next l2
    Next l1
    '
    s = Chr$(66)
    Debug.Assert MemByte(StrPtr(s)) = 66
    Debug.Assert MemByte(StrPtr(s) + 1) = 0
    Debug.Assert MemByte(StrPtr(s) - 4) = 2 'Byte Count in BSTR
End Sub

Private Sub TestWriteByte()
    Dim b As Byte
    Dim i As Integer
    Dim l1 As Long, l2 As Long, l3 As Long, l4 As Long
    Dim l As Long
    Dim s As String
    '
    For l1 = 0 To 255
        MemByte(VarPtr(b)) = l1
        Debug.Assert b = l1
    Next l1
    '
    For l1 = 0 To 255
        For l2 = 0 To 255
            MemByte(VarPtr(i)) = l1
            MemByte(VarPtr(i) + 1) = l2
            Debug.Assert i = ((l1 Xor &H8000) + l2 * 256) Xor &H8000
        Next l2
    Next l1
    '
    For l1 = 0 To 255 Step 25
        For l2 = 0 To 255 Step 25
            For l3 = 0 To 255 Step 5
                For l4 = 0 To 255 Step 5
                    MemByte(VarPtr(l)) = l1
                    MemByte(VarPtr(l) + 1) = l2
                    MemByte(VarPtr(l) + 2) = l3
                    MemByte(VarPtr(l) + 3) = l4
                    Debug.Assert l = (((l1 + l2 * 256 + l3 * 256 ^ 2) Xor &H80000000) + l4 * 256 ^ 3) Xor &H80000000
                Next l4
            Next l3
        Next l2
    Next l1
    '
    s = Space(1)
    MemByte(StrPtr(s)) = 66
    Debug.Assert s = "B"
    MemByte(StrPtr(s) - 4) = 4
    Debug.Assert LenB(s) = 4
    MemByte(StrPtr(s) - 4) = 2
    Debug.Assert LenB(s) = 2
End Sub

Private Sub TestReadInteger()
    Dim i As Integer
    Dim l1 As Long, l2 As Long
    Dim l As Long
    Dim s As String
    '
    For l1 = &H8000 To &H7FFF
        i = l1
        Debug.Assert MemInt(VarPtr(i)) = i
    Next l1
    '
    For l1 = &H8000 To &H7FFF Step 128
        For l2 = &H8000 To &H7FFF Step 128
            l = l1 + IIf(l1 And &H8000, &H10000, 0) + l2 * &H10000
            Debug.Assert MemInt(VarPtr(l)) = l1
            Debug.Assert MemInt(VarPtr(l) + 2) = l2
        Next l2
    Next l1
    '
    s = Chr$(66) & Chr$(65)
    Debug.Assert MemInt(StrPtr(s)) = 66
    Debug.Assert MemInt(StrPtr(s) + 2) = 65
    Debug.Assert MemInt(StrPtr(s) - 4) = 4 'Byte Count in BSTR
End Sub

Private Sub TestWriteInteger()
    Dim i As Integer
    Dim l1 As Long, l2 As Long
    Dim l As Long
    Dim s As String
    '
    For l1 = &H8000 To &H7FFF
        MemInt(VarPtr(i)) = l1
        Debug.Assert i = l1
    Next l1
    '
    For l1 = &H8000 To &H7FFF Step 128
        For l2 = &H8000 To &H7FFF Step 128
            MemInt(VarPtr(l)) = l1
            MemInt(VarPtr(l) + 2) = l2
            Debug.Assert l = l1 + IIf(l1 And &H8000, &H10000, 0) + l2 * &H10000
        Next l2
    Next l1
    '
    s = Space(2)
    MemInt(StrPtr(s)) = 66
    Debug.Assert Mid$(s, 1, 1) = "B"
    MemInt(StrPtr(s) + 2) = 65
    Debug.Assert Mid$(s, 2, 1) = "A"
    MemInt(StrPtr(s) - 4) = 2
    Debug.Assert LenB(s) = 2
    MemInt(StrPtr(s) - 4) = 4
    Debug.Assert LenB(s) = 4
End Sub

Private Sub TestReadLong()
    Dim l As Long
    Dim s As String
    Dim c As Currency
    '
    For l = &H80000000 To &H7FFFFFFF - &H1000 Step &H1000
        Debug.Assert MemLong(VarPtr(l)) = l
        c = l / 10000
        Debug.Assert MemLong(VarPtr(c)) = l
        Debug.Assert MemLong(VarPtr(c) + 4) = IIf(l And &H80000000, -1, 0)
    Next l
    l = &H7FFFFFFF
    Debug.Assert MemLong(VarPtr(l)) = l
    c = l / 10000
    Debug.Assert MemLong(VarPtr(c)) = l
    '
    s = Chr$(65) & Chr$(66)
    Debug.Assert MemLong(StrPtr(s)) = 65 + 66 * 256 ^ 2
    Debug.Assert MemLong(StrPtr(s) + 2) = 66
    Debug.Assert MemLong(StrPtr(s) - 4) = 4 'Byte Count in BSTR
End Sub

Private Sub TestWriteLong()
    Dim l As Long, l1 As Long
    Dim s As String
    Dim c As Currency
    '
    For l1 = &H80000000 To &H7FFFFFFF - &H1000 Step &H1000
        MemLong(VarPtr(l)) = l1
        Debug.Assert l = l1
        '
        MemLong(VarPtr(c)) = l1
        MemLong(VarPtr(c) + 4) = IIf(l And &H80000000, -1, 0)
        Debug.Assert c = l1 / 10000
    Next l1
    l1 = &H7FFFFFFF
    MemLong(VarPtr(l)) = l1
    Debug.Assert l = l1
    MemLong(VarPtr(c)) = l1
    Debug.Assert c = l1 / 10000
    '
    s = Chr$(65) & Chr$(66)
    Debug.Assert MemLong(StrPtr(s)) = 65 + 66 * 256 ^ 2
    Debug.Assert MemLong(StrPtr(s) + 2) = 66
    Debug.Assert MemLong(StrPtr(s) - 4) = 4 'Byte Count in BSTR
    '
    s = Space(2)
    MemLong(StrPtr(s)) = 65 + 66 * 256 ^ 2
    Debug.Assert Mid$(s, 1, 2) = "AB"
    MemLong(StrPtr(s) - 4) = 2
    Debug.Assert LenB(s) = 2
    MemInt(StrPtr(s) - 4) = 4
    Debug.Assert LenB(s) = 4
End Sub

Private Sub TestReadLongLong()
#If Win64 Then
    Dim l As Long
    Dim ll As LongLong
    Dim s As String
    '
    For l = 1 To 63
        ll = -2 ^ l
        Debug.Assert MemLongLong(VarPtr(ll)) = ll
        ll = 2 ^ l - 1
        Debug.Assert MemLongLong(VarPtr(ll)) = ll
    Next l
    '
    s = Chr$(65) & Chr$(66) & Chr$(67) & Chr$(68)
    Debug.Assert MemLongLong(StrPtr(s)) = 65 + 66 * 256 ^ 2 + 67 * 256 ^ 4 + 68 * 256 ^ 6
    Debug.Assert MemLongLong(VarPtr(s)) = StrPtr(s)
#End If
End Sub

Private Sub TestWriteLongLong()
#If Win64 Then
    Dim l As Long
    Dim ll As LongLong, ptr As LongPtr
    Dim s As String, s2 As String
    '
    For l = 1 To 63
        MemLongLong(VarPtr(ll)) = -2 ^ l
        Debug.Assert ll = -2 ^ l
        MemLongLong(VarPtr(ll)) = 2 ^ l - 1
        Debug.Assert ll = 2 ^ l - 1
    Next l
    '
    s = Space(4)
    MemLongLong(StrPtr(s)) = 65 + 66 * 256 ^ 2 + 67 * 256 ^ 4 + 68 * 256 ^ 6
    Debug.Assert Mid$(s, 1, 4) = "ABCD"
    '
    s2 = "TEST"
    ptr = StrPtr(s)
    MemLongLong(VarPtr(s)) = StrPtr(s2)
    Debug.Assert Mid$(s, 1, 4) = "TEST"
    MemLongLong(VarPtr(s)) = ptr
    Debug.Assert Mid$(s, 1, 4) = "ABCD"
#End If
End Sub
