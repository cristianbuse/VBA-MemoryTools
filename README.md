# VBA-MemoryTools
Native memory manipulation in VBA

Using CopyMemory API (RtlMoveMemory on Windows and MemMove on Mac) is quite slow when used many times. Moreover, on some systems this Memory API is even slower due to certain software (e.g. Windows Defender - see [article](https://stackoverflow.com/questions/57885185/windows-defender-extremly-slowing-down-macro-only-on-windows-10)). The API can become so slow that is pretty much unusable (e.g. on my x32 machine it is 600 times slower than it used to be). Using the **LibMemory** module presented here overcomes the speed issues for reading and writing 1, 2, 4 and 8 bytes from and into memory.

## Implementation
Same technique used [here](https://github.com/cristianbuse/VBA-WeakReference) was implemented. A remote Variant allows the changing of the VarType on a second Variant which in turn reads memory remotely as well (has VT_BYREF flag set). A single CopyMemory API call is done when initializing the mentioned remote VarType. Subsequent usage relies on native VBA code only.

4 main parametric properties (Get/Let) are exposed:
 1. MemByte
 2. MetInt 
 3. MemLong
 4. MemLongPtr

A function for UnsignedAddition of Integers is also exposed.

## Installation
Just import the following code modules in your VBA Project:
* **LibMemory.cls**

## Demo

```VBA
Sub DemoMem()
    #If VBA7 Then
        Dim ptr As LongPtr
    #Else
        Dim ptr As Long
    #End If
    Dim i As Long
    Dim arr() As Variant
    ptr = ObjPtr(Application)
    '
    'Read Memory using MemByte
    ReDim arr(0 To PTR_SIZE - 1)
    For i = LBound(arr) To UBound(arr)
        arr(i) = MemByte(UnsignedAddition(ptr, i))
    Next i
    Debug.Print Join(arr, " ")
    '
    'Read Memory using MemInt
    ReDim arr(0 To PTR_SIZE / 2 - 1)
    For i = LBound(arr) To UBound(arr)
        arr(i) = MemInt(UnsignedAddition(ptr, i * 2))
    Next i
    Debug.Print Join(arr, " ")
    '
    'Read Memory using MemLong
    ReDim arr(0 To PTR_SIZE / 4 - 1)
    For i = LBound(arr) To UBound(arr)
        arr(i) = MemLong(UnsignedAddition(ptr, i * 4))
    Next i
    Debug.Print Join(arr, " ")
    '
    'Read Memory using MemLongPtr
    Debug.Print MemLongPtr(ptr)
    '
    'Write Memory using MemByte
    ptr = 0
    MemByte(VarPtr(ptr)) = 24
    Debug.Assert ptr = 24
    MemByte(UnsignedAddition(VarPtr(ptr), 2)) = 24
    Debug.Assert ptr = 1572888
    '
    'Write Memory using MemInt
    ptr = 0
    MemInt(UnsignedAddition(VarPtr(ptr), 2)) = 300
    Debug.Assert ptr = 19660800
    '
    'Write Memory using MemLong
    ptr = 0
    MemLong(VarPtr(ptr)) = 77777
    Debug.Assert ptr = 77777
    '
    'Write Memory using MemLongPtr
    MemLongPtr(VarPtr(ptr)) = ObjPtr(Application)
    Debug.Assert ptr = ObjPtr(Application)
End Sub
```

## Notes
* CopyMemory API is also exposed just in case the 4 main methods are not satisfying the requirement (e.g. copy 50 bytes at once)

## License
MIT License

Copyright (c) 2020 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.