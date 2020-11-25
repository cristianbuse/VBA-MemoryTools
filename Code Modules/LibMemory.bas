Attribute VB_Name = "LibMemory"
'''=============================================================================
''' VBA MemoryTools
'''------------------------------------------------
''' https://github.com/cristianbuse/VBA-MemoryTools
'''------------------------------------------------
'''
''' Copyright (c) 2020 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to deal
''' in the Software without restriction, including without limitation the rights
''' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
''' copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in all
''' copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
''' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
''' SOFTWARE.
'''=============================================================================

Option Explicit
Option Private Module

'*******************************************************************************
'' Methods in this library module allow direct native memory manipulation in VBA
'' regardless of:
''  - the host Application (Excel, Word, AutoCAD etc.)
''  - the operating system (Mac, Windows)
''  - application environment (x32, x64)
'' A single API call to RtlMoveMemory/MemMove is needed to start the remote
''  referencing mechanism. Inside a REMOTE_MEMORY type, a 'remoteVT' Variant is
''  used to manipulate the VarType of the 'memValue' that we want to read/write.
''  The remote manipulation of the VarType is done by setting the VT_BYREF flag
''  on the 'remoteVT' Variant. This is done by using a CopyMemory API but only
''  once per initialization of the REMOTE_MEMORY variable. Once the flag is set,
''  the 'remoteVT' is used to change the VarType of the first Variant just by
''  using a native VBA assignment operation (needs a utility method for correct
''  redirection). In order for the 'memValue' Variant to point to a specific
''  memory address, 2 steps are needed:
''   1) the required address is assigned to the 'memValue' Variant
''   2) the VarType of the 'memValue' Variant is remotely changed via the
''      'remoteVT' Variant while making sure the VT_BYREF flag is also set
''
'' Note that the 'DeRefMem' method could have been a function returning a
''  REMOTE_MEMORY type and thus greatly improving readability but this approach
''  proved to be 2x slower in testing
'*******************************************************************************

'Used for raising errors
Private Const MODULE_NAME As String = "LibMemory"

#If Mac Then
    #If VBA7 Then
        Public Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, source As Any, ByVal Length As LongPtr) As LongPtr
    #Else
        Public Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, source As Any, ByVal Length As Long) As Long
    #End If
#Else 'Windows
    'https://msdn.microsoft.com/en-us/library/mt723419(v=vs.85).aspx
    #If VBA7 Then
        Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As LongPtr)
    #Else
        Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
    #End If
#End If

#If VBA7 Then
    Public Declare PtrSafe Function VarPtrArray Lib "VBE7.dll" Alias "VarPtr" (ByRef ptr() As Any) As LongPtr
#Else
    Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef ptr() As Any) As Long
#End If

'The size in bytes of a memory address
#If Win64 Then
    Public Const PTR_SIZE As Long = 8
#Else
    Public Const PTR_SIZE As Long = 4
#End If

#If Win64 Then
    #If Mac Then
        Public Const vbLongLong As Long = 20 'Apparently missing for x64 on Mac
    #End If
    Public Const vbLongPtr As Long = vbLongLong
#Else
    Public Const vbLongPtr As Long = vbLong
#End If

Private Type REMOTE_MEMORY
    memValue As Variant
    remoteVT As Variant
    isInitialized As Boolean 'In case state is lost
End Type

'https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/3fe7db9f-5803-4dc4-9d14-5425d3f5461f
'https://docs.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-variant?redirectedfrom=MSDN
'Flag used to simulate ByRef Variants
Public Const VT_BYREF As Long = &H4000

Private m_remoteMemory As REMOTE_MEMORY

'*******************************************************************************
'Read/Write a Byte from/to memory
'*******************************************************************************
#If VBA7 Then
Public Property Get MemByte(ByVal memAddress As LongPtr) As Byte
#Else
Public Property Get MemByte(ByVal memAddress As Long) As Byte
#End If
    DeRefMem m_remoteMemory, memAddress, vbByte
    MemByte = m_remoteMemory.memValue
End Property
#If VBA7 Then
Public Property Let MemByte(ByVal memAddress As LongPtr, ByVal newValue As Byte)
#Else
Public Property Let MemByte(ByVal memAddress As Long, ByVal newValue As Byte)
#End If
    DeRefMem m_remoteMemory, memAddress, vbByte
    LetByRef(m_remoteMemory.memValue) = newValue
End Property

'*******************************************************************************
'Read/Write 2 Bytes (Integer) from/to memory
'*******************************************************************************
#If VBA7 Then
Public Property Get MemInt(ByVal memAddress As LongPtr) As Integer
#Else
Public Property Get MemInt(ByVal memAddress As Long) As Integer
#End If
    DeRefMem m_remoteMemory, memAddress, vbInteger
    MemInt = m_remoteMemory.memValue
End Property

#If VBA7 Then
Public Property Let MemInt(ByVal memAddress As LongPtr, ByVal newValue As Integer)
#Else
Public Property Let MemInt(ByVal memAddress As Long, ByVal newValue As Integer)
#End If
    DeRefMem m_remoteMemory, memAddress, vbInteger
    LetByRef(m_remoteMemory.memValue) = newValue
End Property

'*******************************************************************************
'Read/Write 4 Bytes (Long) from/to memory
'*******************************************************************************
#If VBA7 Then
Public Property Get MemLong(ByVal memAddress As LongPtr) As Long
#Else
Public Property Get MemLong(ByVal memAddress As Long) As Long
#End If
    DeRefMem m_remoteMemory, memAddress, vbLong
    MemLong = m_remoteMemory.memValue
End Property
#If VBA7 Then
Public Property Let MemLong(ByVal memAddress As LongPtr, ByVal newValue As Long)
#Else
Public Property Let MemLong(ByVal memAddress As Long, ByVal newValue As Long)
#End If
    DeRefMem m_remoteMemory, memAddress, vbLong
    LetByRef(m_remoteMemory.memValue) = newValue
End Property

'*******************************************************************************
'Read/Write 8 Bytes (LongLong) from/to memory
'*******************************************************************************
#If VBA7 Then
Public Property Get MemLongPtr(ByVal memAddress As LongPtr) As LongPtr
#Else
Public Property Get MemLongPtr(ByVal memAddress As Long) As Long
#End If
    DeRefMem m_remoteMemory, memAddress, vbLongPtr
    MemLongPtr = m_remoteMemory.memValue
End Property
#If VBA7 Then
Public Property Let MemLongPtr(ByVal memAddress As LongPtr, ByVal newValue As LongPtr)
#Else
Public Property Let MemLongPtr(ByVal memAddress As Long, ByVal newValue As Long)
#End If
    #If Win64 Then
        'Cannot set Variant/LongLong ByRef so we use a Currency instead
        Const currDivider As Currency = 10000
        DeRefMem m_remoteMemory, memAddress, vbCurrency
        LetByRef(m_remoteMemory.memValue) = CCur(newValue / currDivider)
    #Else
        MemLong(memAddress) = newValue
    #End If
End Property

'*******************************************************************************
'Redirects the rm.memValue Variant to the new memory address so that the value
'   can be read ByRef
'*******************************************************************************
Private Sub DeRefMem(ByRef rm As REMOTE_MEMORY, ByRef memAddress As LongPtr, ByRef vt As VbVarType)
    With rm
        If Not .isInitialized Then
            'Link .remoteVt to the first 2 bytes of the .memValue Variant
            .remoteVT = VarPtr(.memValue)
            CopyMemory .remoteVT, vbInteger + VT_BYREF, 2
            '
            .isInitialized = True
        End If
        'Link .memValue to the desired address
        .memValue = memAddress
        LetByRef(.remoteVT) = vt + VT_BYREF 'Faster than: CopyMemory .memValue, vt + VT_BYREF, 2
    End With
End Sub

'*******************************************************************************
'Utility for updating remote values that have the VT_BYREF flag set
'*******************************************************************************
Private Property Let LetByRef(ByRef v As Variant, ByRef newValue As Variant)
    v = newValue
End Property

'*******************************************************************************
'Unsigned Addition
'
'VBA does not allow the declaration of unsigned integers. The integers are
'   always signed and can store both positive and negative numbers.
'
'-------------------------------------------------
'Basic information on bits and bytes
'-------------------------------------------------
'Bit: a basic unit of information used in computing.
'   Can only have one of two values: 0 or 1
'Nibble: a set of 4 bits
'   Can have binary values from 0000 to 1111 (0 to 15 in decimal notation)
'   Ex. 1001 (bin) = 1*2^3 + 0*2^2 + 0*2^1 + 1*2^0 = 8 + 0 + 0 + 1 = 9 (dec)
'Byte: unit of digital information that consists of 8 bits (or 2 nibbles)
'   Can have binary values from 00000000 to 11111111 (0 to 255 in decimal)
'   In VBA a Byte is an unsigned type
'
'-------------------------------------------------
'Signed VBA Integer Data Types
'-------------------------------------------------
'Integer: 2 Bytes (16 bits)
'   Can store values from -32,768 to 32,767 (decimal)
'Long: 4 Bytes (32 bits)
'   Can store values from -2,147,483,648 to 2,147,483,647 (decimal)
'LongLong: 8 Bytes (or 64 bits)
'   Values from -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 (decimal)
'   Available only in x64 versions of Applications using VBA
'
'In VBA the "Two's complement" mathematical operation method is used to
'   represent both negative and positive numbers (signed numbers)
'   See: https://en.wikipedia.org/wiki/Two%27s_complement
'   Ex. for 0101 (5), we reverse bits and add 1 and we get 1011 (-5)
'In signed integers the left-most bit is used to indicate the sign, so, a value
'   of 1 in the left-most bit indicates a negative number and a value of 0 in
'   the left-most bit indicates a non-negative number (zero or positive)
'
'-------------------------------------------------
'Unsigned vs Signed example
'-------------------------------------------------
'A 2-byte unsigned Integer would have binary values from 0000 0000 0000 0000 to
'   1111 1111 1111 1111 (0 to 65,535 in decimal)
'A 2-byte signed Integer has binary values from 0000 0000 0000 0000 to
'   0111 1111 1111 1111 (positive 0 to 32,767 in decimal) and binary values from
'   1000 0000 0000 0000 to 1111 1111 1111 1111 (negative -32,768 to -1 in decimal)
'
'-------------------------------------------------
'Hexadecimal notation
'-------------------------------------------------
'Hex notation is a base-16 system where each digit is a value between 0 an 15
'In Decimal, each digit is between 0 an 9 and in Binary each digit is 0 or 1
'
'In hex notation values from 0 to 9 are the same as in Decimal but values
'   10 to 15 are written as A to F
'So, a nibble corresponds to a digit in hex (0 to F) and a byte can be written
'   as a 2 digit hex number with values from 00 to FF
'
'In VBA the hex numbers are prefixed by &H characters. 00 -> &H00; FF -> &HFF
'   Ex.: &H7E = 7*16^1 + 15*16^0 = 126 (dec) or 0111 1110 in binary
'
'Hex notation provides a very convenient way to write byte values
'
'If a 2-byte Integer Type would be unsigned then it's values 0 to 65,535 could be
'   written as &H0000 to &HFFFF but because the Integer Type is signed then it's
'   values -32,768 to 32,767 could be written in hex as follows:
'       1000000000000000 to 1111111111111111 (-32,768 to -1) as &H8000 to &HFFFF
'       0000000000000000 to 0111111111111111 (0 to 32,767) as &H0000 to &H7FFF
'
'-------------------------------------------------
'LongPtr
'-------------------------------------------------
'https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/longptr-data-type
'Not a true data type. It transforms to Long(x32) or LongLong(x64)
'
'Long:
'   -2,147,483,648 to -1 (dec) corresponds to &H80000000 to &HFFFFFFFF (hex)
'    0 to +2,147,483,647 (dec) corresponds to &H00000000 to &H7FFFFFFF (hex)
'LongLong:
'    0 to +9,223,372,036,854,775,807 corresponds to &H0000000000000000 to &H7FFFFFFFFFFFFFFF
'   -9,223,372,036,854,775,808 to -1 corresponds to &H8000000000000000 to &HFFFFFFFFFFFFFFFF
'   https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/longlong-data-type
'
'Notes:
'   To declare a value as Long in VBA the following suffix is used: &
'   To declare a value as LongLong in VBA the following suffix is used: ^
'   Example:
'       &H8000  (Integer) = -32,768
'       &H8000& (Long)    = +32,768
'       &H80000000  (Long) =  -2,147,483,648 (can't be Integer; "&" not needed)
'       &H80000000& (Long) =  -2,147,483,648
'       &H80000000^ (LongLong) =  +2,147,483,648 (64-bit platforms only)
'       &H8000000000000000 (LongLong) = will not compile!
'       &H8000000000000000^ (LongLong) = -9,223,372,036,854,775,808
'
'-------------------------------------------------
'Memory address
'-------------------------------------------------
'Memory addresses are fixed-length sequences of digits conventionally displayed
'   and manipulated as unsigned integers
'
'When given a memory address as a signed integer and a positive increment, in
'   order to find the correct address+increment, the 2 values must take
'   into account the limits described above.
'
'-------------------------------------------------
'Overflow Example
'-------------------------------------------------
'   Assume the 2 Long Integer numbers:
'       A = &H7FFFFFFD (hex) = 2147483645 (dec)
'       B = &H0000000C (hex) =         12 (dec)
'   If the 2 integers would be Unsigned then their sum would be:
'       S = A + B = &H7FFFFFFD + &H0000000C = &H80000009
'       or in decimal
'       S = A + B = 2147483645 +         12 = +2147483657
'   But because the 2 integers are Signed then their sum exceeds the limit
'       of 2,147,483,648 available in a Long data type
'   In VBA, the signed number S = &H80000009 = -2147483639
'
'-------------------------------------------------
'Aim of the function
'-------------------------------------------------
'The "UnsignedAddition" function avoids overflow errors as in the example above
'   by adding the minimum negative value as needed
'*******************************************************************************
#If VBA7 Then
Public Function UnsignedAddition(ByVal val1 As LongPtr, ByVal val2 As LongPtr) As LongPtr
#Else
Public Function UnsignedAddition(ByVal val1 As Long, ByVal val2 As Long) As Long
#End If
    'The minimum negative integer value of a Long Integer in VBA
    #If Win64 Then
    Const minNegative As LongLong = &H8000000000000000^ '-9,223,372,036,854,775,808 (dec)
    #Else
    Const minNegative As Long = &H80000000 '-2,147,483,648 (dec)
    #End If
    '
    If val1 > 0 Then
        If val2 > 0 Then
            'Overflow could occur
            If (val1 + minNegative + val2) < 0 Then
                'The sum will not overflow
                UnsignedAddition = val1 + val2
            Else
                'Example for Long data type (x32):
                '   &H7FFFFFFD + &H0000000C =  &H80000009
                '   2147483645 +         12 = -2147483639
                UnsignedAddition = val1 + minNegative + val2 + minNegative
            End If
        Else 'Val2 <= 0
            'Sum cannot overflow
            UnsignedAddition = val1 + val2
        End If
    Else 'Val1 <= 0
        If val2 > 0 Then
            'Sum cannot overflow
            UnsignedAddition = val1 + val2
        Else 'Val2 <= 0
            'Overflow could occur
            On Error GoTo ErrorHandler
            UnsignedAddition = val1 + val2
        End If
    End If
Exit Function
ErrorHandler:
    Err.Raise 6, MODULE_NAME & ".UnsignedAddition", "Overflow"
End Function
