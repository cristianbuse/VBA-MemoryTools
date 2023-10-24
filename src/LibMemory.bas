Attribute VB_Name = "LibMemory"
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

'*******************************************************************************
'' Methods in this library module allow direct native memory manipulation in VBA
'' regardless of:
''  - the host Application (Excel, Word, AutoCAD etc.)
''  - the operating system (Mac, Windows)
''  - application environment (x32, x64)
'' A single API call to RtlMoveMemory API is needed to start the remote
''  referencing mechanism. Inside a REMOTE_MEMORY type, a 'remoteVT' Variant is
''  used to manipulate the VarType of the 'memValue' that we want to read/write.
''  The remote manipulation of the VarType is done by setting the VT_BYREF flag
''  on the 'remoteVT' Variant. This is done by using a CopyMemory API but only
''  once per initialization of the REMOTE_MEMORY variable (see MemIntAPI). Note
''  that besides the Static REMOTE_MEMORY used in MemIntAPI all the other memory
''  structs are initialized through InitRemoteMemory with no API calls. Once the
''  flag is set, the 'remoteVT' is used to change the VarType of the first
''  Variant just by using a native VBA assignment operation (needs a utility
''  method for correct redirection). In order for the 'memValue' Variant to
''  point to aspecific memory address, 2 steps are needed:
''   1) the required address is assigned to the 'memValue' Variant
''   2) the VarType of the 'memValue' Variant is remotely changed via the
''      'remoteVT' Variant (must be done ByRef in any utility method)
'*******************************************************************************

'Used for raising errors
Private Const MODULE_NAME As String = "LibMemory"

#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
        Private Declare PtrSafe Function FillMemory Lib "/usr/lib/libc.dylib" Alias "memset" (Destination As Any, ByVal Fill As Byte, ByVal Length As LongPtr) As LongPtr
    #Else
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
        Private Declare Function FillMemory Lib "/usr/lib/libc.dylib" Alias "memset" (Destination As Any, ByVal Fill As Byte, ByVal Length As Long) As Long
    #End If
#Else 'Windows
    'https://msdn.microsoft.com/en-us/library/mt723419(v=vs.85).aspx
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
        Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
        Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
    #End If
#End If

#If VBA7 = 0 Then       'LongPtr trick discovered by @Greedo (https://github.com/Greedquest)
    Public Enum LongPtr
        [_]
    End Enum
#End If                 'https://github.com/cristianbuse/VBA-MemoryTools/issues/3

#If Win64 Then
    Public Const PTR_SIZE As Long = 8
    Public Const VARIANT_SIZE As Long = 24
#Else
    Public Const PTR_SIZE As Long = 4
    Public Const VARIANT_SIZE As Long = 16
#End If

Private Const BYTE_SIZE As Long = 1
Private Const INT_SIZE As Long = 2
Private Const VT_SPACING As Long = VARIANT_SIZE / INT_SIZE 'VarType spacing in an array of Variants

#If Win64 Then
    #If Mac Then
        Public Const vbLongLong As Long = 20 'Apparently missing for x64 on Mac
    #End If
    Public Const vbLongPtr As Long = vbLongLong
#Else
    Public Const vbLongLong As Long = 20 'Useful in Select Case logic
    Public Const vbLongPtr As Long = vbLong
#End If

Public Type REMOTE_MEMORY
    memValue As Variant
    remoteVT As Variant 'Will be linked to the first 2 bytes of 'memValue' - see 'InitRemoteMemory'
    isInitialized As Boolean 'In case state is lost
End Type

'https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/3fe7db9f-5803-4dc4-9d14-5425d3f5461f
'https://docs.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-variant?redirectedfrom=MSDN
'Flag used to simulate ByRef Variants
Public Const VT_BYREF As Long = &H4000

Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Public Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    #If Win64 Then
        dummyPadding As Long
        pvData As LongLong
    #Else
        pvData As Long
    #End If
    rgsabound0 As SAFEARRAYBOUND
End Type
Private Const FADF_HAVEVARTYPE As Long = &H80

'*******************************************************************************
'Returns an initialized (linked) REMOTE_MEMORY struct
'Links .remoteVt to the first 2 bytes of .memValue
'*******************************************************************************
Public Sub InitRemoteMemory(ByRef rm As REMOTE_MEMORY)
    rm.remoteVT = VarPtr(rm.memValue)
    MemIntAPI(VarPtr(rm.remoteVT)) = vbInteger + VT_BYREF
    rm.isInitialized = True
End Sub

'*******************************************************************************
'The only method in this module that uses CopyMemory!
'Assures that InitRemoteMemory can link the Var Type for new structs
'*******************************************************************************
Private Property Let MemIntAPI(ByVal memAddress As LongPtr, ByVal newValue As Integer)
    Static rm As REMOTE_MEMORY
    If Not rm.isInitialized Then 'Link .remoteVt to .memValue's first 2 bytes
        rm.remoteVT = VarPtr(rm.memValue)
        CopyMemory rm.remoteVT, vbInteger + VT_BYREF, 2
        rm.isInitialized = True
    End If
    RemoteAssign rm, memAddress, rm.remoteVT, vbInteger + VT_BYREF, rm.memValue, newValue
End Property

'*******************************************************************************
'This method assures the required redirection for both the remote varType and
'   the remote value at the same time thus removing any additional stack frames
'It can be used to both read from and write to memory by swapping the order of
'   the last 2 parameters
'*******************************************************************************
Private Sub RemoteAssign(ByRef rm As REMOTE_MEMORY _
                       , ByRef memAddress As LongPtr _
                       , ByRef remoteVT As Variant _
                       , ByVal newVT As VbVarType _
                       , ByRef targetVariable As Variant _
                       , ByRef newValue As Variant)
    rm.memValue = memAddress
    If Not rm.isInitialized Then InitRemoteMemory rm
    remoteVT = newVT
    targetVariable = newValue
    remoteVT = vbEmpty 'Stop linking to remote address, for safety
End Sub

'*******************************************************************************
'Read/Write a Byte from/to memory
'*******************************************************************************
Public Property Get MemByte(ByVal memAddress As LongPtr) As Byte
    #If Mac Or (VBA7 = 0) Then
        CopyMemory MemByte, ByVal memAddress, 1
    #ElseIf TWINBASIC Then
        GetMem1 memAddress, MemByte
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbByte + VT_BYREF, MemByte, rm.memValue
    #End If
End Property
Public Property Let MemByte(ByVal memAddress As LongPtr, ByVal newValue As Byte)
    #If Mac Or (VBA7 = 0) Then
        CopyMemory ByVal memAddress, newValue, 1
    #ElseIf TWINBASIC Then
        PutMem1 memAddress, newValue
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbByte + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 2 Bytes (Integer) from/to memory
'*******************************************************************************
Public Property Get MemInt(ByVal memAddress As LongPtr) As Integer
    #If Mac Or (VBA7 = 0) Then
        CopyMemory MemInt, ByVal memAddress, 2
    #ElseIf TWINBASIC Then
        GetMem2 memAddress, MemInt
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbInteger + VT_BYREF, MemInt, rm.memValue
    #End If
End Property
Public Property Let MemInt(ByVal memAddress As LongPtr, ByVal newValue As Integer)
    #If Mac Or (VBA7 = 0) Then
        CopyMemory ByVal memAddress, newValue, 2
    #ElseIf TWINBASIC Then
        PutMem2 memAddress, newValue
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbInteger + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 2 Bytes (Boolean) from/to memory
'*******************************************************************************
Public Property Get MemBool(ByVal memAddress As LongPtr) As Boolean
    #If Mac Or TWINBASIC Or (VBA7 = 0) Then
        CopyMemory MemBool, ByVal memAddress, 2
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbBoolean + VT_BYREF, MemBool, rm.memValue
    #End If
End Property
Public Property Let MemBool(ByVal memAddress As LongPtr, ByVal newValue As Boolean)
    #If Mac Or TWINBASIC Or (VBA7 = 0) Then
        CopyMemory ByVal memAddress, newValue, 2
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbBoolean + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 4 Bytes (Long) from/to memory
'*******************************************************************************
Public Property Get MemLong(ByVal memAddress As LongPtr) As Long
    #If Mac Or (VBA7 = 0) Then
        CopyMemory MemLong, ByVal memAddress, 4
    #ElseIf TWINBASIC Then
        GetMem4 memAddress, MemLong
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbLong + VT_BYREF, MemLong, rm.memValue
    #End If
End Property
Public Property Let MemLong(ByVal memAddress As LongPtr, ByVal newValue As Long)
    #If Mac Or (VBA7 = 0) Then
        CopyMemory ByVal memAddress, newValue, 4
    #ElseIf TWINBASIC Then
        PutMem4 memAddress, newValue
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbLong + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 4 Bytes (Single) from/to memory
'*******************************************************************************
Public Property Get MemSng(ByVal memAddress As LongPtr) As Single
    #If Mac Or TWINBASIC Or (VBA7 = 0) Then
        CopyMemory MemSng, ByVal memAddress, 4
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbSingle + VT_BYREF, MemSng, rm.memValue
    #End If
End Property
Public Property Let MemSng(ByVal memAddress As LongPtr, ByVal newValue As Single)
    #If Mac Or TWINBASIC Or (VBA7 = 0) Then
        CopyMemory ByVal memAddress, newValue, 4
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbSingle + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 8 Bytes (LongLong) from/to memory
'*******************************************************************************
#If Win64 Or TWINBASIC Then
Public Property Get MemLongLong(ByVal memAddress As LongLong) As LongLong
    #If Mac Or TWINBASIC Or (VBA7 = 0) Then
        CopyMemory MemLongLong, ByVal memAddress, 8
    #Else
        'Cannot set Variant/LongLong ByRef so we cannot use 'RemoteAssign'
        Static rm As REMOTE_MEMORY: rm.memValue = memAddress
        MemLongLong = ByRefLongLong(rm, rm.remoteVT, rm.memValue)
    #End If
End Property
Public Property Let MemLongLong(ByVal memAddress As LongLong, ByVal newValue As LongLong)
    #If Mac Or TWINBASIC Or (VBA7 = 0) Then
        CopyMemory ByVal memAddress, newValue, 8
    #Else
        'Cannot set Variant/LongLong ByRef so we use Currency instead
        Static rmSrc As REMOTE_MEMORY: rmSrc.memValue = VarPtr(newValue)
        Static rmDest As REMOTE_MEMORY: rmDest.memValue = memAddress
        LetByRefLongLong rmDest, rmDest.remoteVT, rmDest.memValue _
                       , rmSrc, rmSrc.remoteVT, rmSrc.memValue
    #End If
End Property
Private Property Get ByRefLongLong(ByRef rm As REMOTE_MEMORY _
                                 , ByRef vt As Variant _
                                 , ByRef memValue As Variant) As LongLong
    If Not rm.isInitialized Then InitRemoteMemory rm
    vt = vbLongLong + VT_BYREF
    ByRefLongLong = memValue
    vt = vbEmpty
End Property
Private Sub LetByRefLongLong(ByRef rmDest As REMOTE_MEMORY _
                           , ByRef vtDest As Variant _
                           , ByRef memValueDest As Variant _
                           , ByRef rmSrc As REMOTE_MEMORY _
                           , ByRef vtSrc As Variant _
                           , ByRef memValueSrc As Variant)
    If Not rmSrc.isInitialized Then InitRemoteMemory rmSrc
    If Not rmDest.isInitialized Then InitRemoteMemory rmDest
    vtDest = vbCurrency + VT_BYREF
    vtSrc = vbCurrency + VT_BYREF
    memValueDest = memValueSrc
    vtDest = vbEmpty
    vtSrc = vbEmpty
End Sub
#End If

'*******************************************************************************
'Read/Write 4 Bytes (Long on x32) or 8 Bytes (LongLong on x64) from/to memory
'Note that wrapping MemLong and MemLongLong is about 25% slower because of the
'   extra stack frame! Performance was chosen over code repetition!
'*******************************************************************************
Public Property Get MemLongPtr(ByVal memAddress As LongPtr) As LongPtr
    #If Mac Or (VBA7 = 0) Then
        CopyMemory MemLongPtr, ByVal memAddress, PTR_SIZE
    #ElseIf TWINBASIC Then
        GetMemPtr memAddress, MemLongPtr
    #ElseIf Win64 Then
        Static rm As REMOTE_MEMORY: rm.memValue = memAddress
        MemLongPtr = ByRefLongLong(rm, rm.remoteVT, rm.memValue)
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbLong + VT_BYREF, MemLongPtr, rm.memValue
    #End If
End Property
Public Property Let MemLongPtr(ByVal memAddress As LongPtr, ByVal newValue As LongPtr)
    #If Mac Or (VBA7 = 0) Then
        CopyMemory ByVal memAddress, newValue, PTR_SIZE
    #ElseIf TWINBASIC Then
        PutMemPtr memAddress, newValue
    #ElseIf Win64 Then
        Static rmSrc As REMOTE_MEMORY: rmSrc.memValue = VarPtr(newValue)
        Static rmDest As REMOTE_MEMORY: rmDest.memValue = memAddress
        LetByRefLongLong rmDest, rmDest.remoteVT, rmDest.memValue, rmSrc, rmSrc.remoteVT, rmSrc.memValue
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbLong + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 8 Bytes (Currency) from/to memory
'*******************************************************************************
Public Property Get MemCur(ByVal memAddress As LongPtr) As Currency
    #If Mac Or (VBA7 = 0) Then
        CopyMemory MemCur, ByVal memAddress, 8
    #ElseIf TWINBASIC Then
        GetMem8 memAddress, MemCur
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbCurrency + VT_BYREF, MemCur, rm.memValue
    #End If
End Property
Public Property Let MemCur(ByVal memAddress As LongPtr, ByVal newValue As Currency)
    #If Mac Or (VBA7 = 0) Then
        CopyMemory ByVal memAddress, newValue, 8
    #ElseIf TWINBASIC Then
        PutMem8 memAddress, newValue
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbCurrency + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 8 Bytes (Date) from/to memory
'*******************************************************************************
Public Property Get MemDate(ByVal memAddress As LongPtr) As Date
    #If Mac Or TWINBASIC Or (VBA7 = 0) Then
        CopyMemory MemDate, ByVal memAddress, 8
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbDate + VT_BYREF, MemDate, rm.memValue
    #End If
End Property
Public Property Let MemDate(ByVal memAddress As LongPtr, ByVal newValue As Date)
    #If Mac Or TWINBASIC Or (VBA7 = 0) Then
        CopyMemory ByVal memAddress, newValue, 8
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbDate + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 8 Bytes (Double) from/to memory
'*******************************************************************************
Public Property Get MemDbl(ByVal memAddress As LongPtr) As Double
    #If Mac Or TWINBASIC Or (VBA7 = 0) Then
        CopyMemory MemDbl, ByVal memAddress, 8
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbDouble + VT_BYREF, MemDbl, rm.memValue
    #End If
End Property
Public Property Let MemDbl(ByVal memAddress As LongPtr, ByVal newValue As Double)
    #If Mac Or TWINBASIC Or (VBA7 = 0) Then
        CopyMemory ByVal memAddress, newValue, 8
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbDouble + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Dereference an object by it's pointer
'*******************************************************************************
Public Function MemObj(ByVal memAddress As LongPtr) As Object
    If memAddress = 0 Then Exit Function
    '
    #If Mac Or TWINBASIC Or (VBA7 = 0) Then
        Dim obj As Object
        CopyMemory obj, memAddress, PTR_SIZE
        Set MemObj = obj
        memAddress = 0 'We don't just use 0 (below) because we need 0& or 0^
        CopyMemory obj, memAddress, PTR_SIZE
    #Else
        Static rm As REMOTE_MEMORY: rm.memValue = memAddress
        If Not rm.isInitialized Then InitRemoteMemory rm
        Set MemObj = RemObject(rm.remoteVT, rm.memValue)
    #End If
End Function
Private Property Get RemObject(ByRef vt As Variant _
                             , ByRef memValue As Variant) As Object
    vt = vbObject
    Set RemObject = memValue
    vt = vbEmpty
End Property

'Method purpose explanation at:
'https://gist.github.com/cristianbuse/b9cc79164c1d31fdb30465f503ac36a9
'
'Practical note Jan-2021 from Vladimir Vissoultchev (https://github.com/wqweto):
'This is mostly not needed in client application code even for LARGEADDRESSAWARE
'   32-bit processes nowadays as a reliable technique to prevent pointer
'   arithmetic overflows is to VirtualAlloc a 64KB sentinel chunk around 2GB
'   boundary at application start up so that the boundary is never (rarely)
'   crossed in normal pointer operations.
'This same sentinel chunk fixes native PropertyBag as well which has troubles
'   when internal storage crosses 2GB boundary.
Public Function UnsignedAdd(ByVal unsignedPtr As LongPtr, ByVal signedOffset As LongPtr) As LongPtr
    #If Win64 Then
        Const minNegative As LongLong = &H8000000000000000^
    #Else
        Const minNegative As Long = &H80000000
    #End If
    UnsignedAdd = ((unsignedPtr Xor minNegative) + signedOffset) Xor minNegative
End Function

'*******************************************************************************
'Redirects the instance of a class to another instance of the same class within
'   the scope of a private class Function (not  Sub) where the call happens.
'
'Warning! ONLY call this method from a Private Function of a class!
'
'vbArray + vbString Function return type is not supported. It would be
'   possible to find the correct address by reading memory in a loop but there
'   would be no checking available
'*******************************************************************************
Public Sub RedirectInstance(ByVal funcReturnPtr As LongPtr _
                          , ByVal currentInstance As Object _
                          , ByVal targetInstance As Object)
    Const methodName As String = "RedirectInstance"
    Dim originalPtr As LongPtr
    Dim newPtr As LongPtr
    Dim swapAddress As LongPtr
    '
    originalPtr = ObjPtr(GetDefaultInterface(currentInstance))
    newPtr = ObjPtr(GetDefaultInterface(targetInstance))
    '
    'Validate Input
    If currentInstance Is Nothing Or targetInstance Is Nothing Then
        Err.Raise 91, methodName, "Object not set"
    ElseIf MemLongPtr(originalPtr) <> MemLongPtr(newPtr) Then 'Faster to compare vTables than to compare TypeName(s)
        Err.Raise 5, methodName, "Expected same VB class"
    ElseIf funcReturnPtr = 0 Then
        Err.Raise 5, methodName, "Missing Function Return Pointer"
    End If
    '
    'On x64 the shadow stack space is allocated next to the Function Return
    'On x32 the stack space has a fixed offset (found through testing)
    #If Win64 Then
        Const memOffsetNonVariant As LongLong = PTR_SIZE
        Const memOffsetVariant As LongLong = PTR_SIZE * 3
    #Else
        Const memOffsetNonVariant As Long = PTR_SIZE * 28
        Const memOffsetVariant As Long = PTR_SIZE * 31
    #End If
    '
    swapAddress = FindSwapAddress(funcReturnPtr, memOffsetNonVariant, originalPtr)
    If swapAddress = 0 Then
        swapAddress = FindSwapAddress(funcReturnPtr, memOffsetVariant, originalPtr)
        If swapAddress = 0 Then
            Err.Raise 5, methodName, "Invalid input or not called " _
            & "from class Function or vbArray + vbString function return type"
        End If
    End If
    '
    'Redirect Instance
    MemLongPtr(swapAddress) = newPtr
End Sub

'*******************************************************************************
'Finds the swap address (address of the instance pointer on the stack)
'*******************************************************************************
#If Win64 Then
Private Function FindSwapAddress(ByVal funcReturnPtr As LongLong _
                               , ByVal memOffset As LongLong _
                               , ByVal originalPtr As LongLong) As LongLong
    Dim swapAddr As LongLong: swapAddr = funcReturnPtr + memOffset
    '
    'Adjust alignment for Boolean/Byte/Integer/Long function return type
    'Needed on #Mac but not on #Win (at least not found in testing)
    'Safer to have for #Win as well
    swapAddr = swapAddr - (swapAddr Mod PTR_SIZE)
    '
    If MemLongLong(swapAddr) = originalPtr Then
        FindSwapAddress = swapAddr
    End If
End Function
#Else
Private Function FindSwapAddress(ByVal funcReturnPtr As Long _
                               , ByVal memOffset As Long _
                               , ByVal originalPtr As Long) As Long
    Dim startAddr As Long: startAddr = funcReturnPtr + memOffset
    Dim swapAddr As Long
    '
    'Adjust memory alignment for Boolean/Byte/Integer function return type
    startAddr = startAddr - (startAddr Mod PTR_SIZE)
    '
    swapAddr = GetSwapIndirectAddress(startAddr)
    If swapAddr = 0 Then
        'Adjust mem alignment for Currency/Date/Double function return type
        swapAddr = GetSwapIndirectAddress(startAddr + PTR_SIZE)
        If swapAddr = 0 Then Exit Function
    End If
    If MemLongPtr(swapAddr) = originalPtr Then
        FindSwapAddress = swapAddr
    End If
End Function
Private Function GetSwapIndirectAddress(ByVal startAddr As Long) As Long
    Const maxOffset As Long = PTR_SIZE * 100
    Dim swapAddr As Long: swapAddr = MemLong(startAddr) + PTR_SIZE * 2
    '
    'Check if the address is within acceptable limits. The address
    '   of the instance pointer (within the stack frame) cannot be too far
    '   from the function return address (first offsetted to startAddr)
    If startAddr < swapAddr And swapAddr - startAddr < maxOffset * PTR_SIZE Then
        GetSwapIndirectAddress = swapAddr
    End If
End Function
#End If

'*******************************************************************************
'Returns the default interface for an object
'Casting from IUnknown to IDispatch (Object) forces a call to QueryInterface for
'   the IDispatch interface (which knows about the default interface)
'*******************************************************************************
Public Function GetDefaultInterface(ByVal obj As IUnknown) As Object
    Set GetDefaultInterface = obj
End Function

'*******************************************************************************
'Returns the memory address of a variable of array type
'Returns error 5 for a non-array
'*******************************************************************************
Public Function VarPtrArr(ByRef arr As Variant) As LongPtr
    Dim vt As VbVarType: vt = MemInt(VarPtr(arr)) 'VarType(arr) ignores VT_BYREF
    If vt And vbArray Then
        Const pArrayOffset As Long = 8
        VarPtrArr = VarPtr(arr) + pArrayOffset
        If vt And VT_BYREF Then VarPtrArr = MemLongPtr(VarPtrArr)
    Else
        Err.Raise 5, "VarPtrArr", "Array required"
    End If
End Function

'*******************************************************************************
'Returns the pointer to the underlying SAFEARRAY structure of a VB array
'Returns error 5 for a non-array
'*******************************************************************************
Public Function ArrPtr(ByRef arr As Variant) As LongPtr
    Dim vt As VbVarType: vt = MemInt(VarPtr(arr)) 'VarType(arr) ignores VT_BYREF
    If vt And vbArray Then
        Const pArrayOffset As Long = 8
        ArrPtr = MemLongPtr(VarPtr(arr) + pArrayOffset)
        If vt And VT_BYREF Then ArrPtr = MemLongPtr(ArrPtr)
    Else
        Err.Raise 5, "ArrPtr", "Array required"
    End If
End Function

'*******************************************************************************
'Alternative for CopyMemory - not affected by API speed issues on Windows
'--------------------------
'Mac - wrapper around CopyMemory/memmove
'Win - bytesCount 1 to 16777216 - no API calls. Uses a combination of
'      REMOTE_MEMORY/SAFEARRAY_1D structs as well as native Strings and Arrays
'      to manipulate memory.
'      For some smaller sizes (<=5) optimizes via MemLong, MemInt, MemByte etc.
'    - bytesCount < 0 or > 16777216 - wrapper around CopyMemory/RtlMoveMemory
'*******************************************************************************
Public Sub MemCopy(ByVal destinationPtr As LongPtr _
                 , ByVal sourcePtr As LongPtr _
                 , ByVal bytesCount As LongPtr)
    If destinationPtr = sourcePtr Then Exit Sub
#If Mac Or (VBA7 = 0) Then
    CopyMemory ByVal destinationPtr, ByVal sourcePtr, bytesCount
#Else
    Const maxSizeSpeedGain As Long = &H1000000 'Beyond this use API directly
    If bytesCount < 0 Or bytesCount > maxSizeSpeedGain Then
        CopyMemory ByVal destinationPtr, ByVal sourcePtr, bytesCount
        Exit Sub
    End If
    '
    If bytesCount <= 4 Then 'Cannot copy via BSTR as destination
        Select Case bytesCount
            Case 1: MemByte(destinationPtr) = MemByte(sourcePtr)
            Case 2: MemInt(destinationPtr) = MemInt(sourcePtr)
            Case 3: MemInt(destinationPtr) = MemInt(sourcePtr)
                    MemByte(destinationPtr + 2) = MemByte(sourcePtr + 2)
            Case 4: MemLong(destinationPtr) = MemLong(sourcePtr)
        End Select
        Exit Sub
    ElseIf bytesCount = 8 Then 'Optional optimization - small gain
        MemCur(destinationPtr) = MemCur(sourcePtr)
        Exit Sub
    End If
    '
    'Structs used to read/write memory
    Static sArrByte As SAFEARRAY_1D
    Static rmArrSrc As REMOTE_MEMORY
    Static rmSrc As REMOTE_MEMORY
    Static rmDest As REMOTE_MEMORY
    Static rmBSTR As REMOTE_MEMORY
    '
    If Not rmArrSrc.isInitialized Then
        With sArrByte
            .cDims = 1
            .fFeatures = FADF_HAVEVARTYPE
            .cbElements = BYTE_SIZE
        End With
        rmArrSrc.memValue = VarPtr(sArrByte)
        '
        InitRemoteMemory rmArrSrc
        InitRemoteMemory rmSrc
        InitRemoteMemory rmDest
        InitRemoteMemory rmBSTR
    End If
    '
    rmSrc.memValue = sourcePtr
    rmDest.memValue = destinationPtr
    CopyBytes CLng(bytesCount), rmSrc, rmSrc.remoteVT, rmDest, rmDest.remoteVT _
        , rmDest.memValue, sArrByte, rmArrSrc.memValue, rmArrSrc.remoteVT _
        , rmBSTR, rmBSTR.remoteVT, rmBSTR.memValue
#End If
End Sub
'*******************************************************************************
'Utility for 'MemCopy' - avoids extra stack frames
'The 'bytesCount' expected to be larger than 4 because the first 4 bytes are
'   needed for the destination BSTR's length.
'The source can either be a String or an array of bytes depending on the first 4
'   bytes in the source. Choice between the 2 is based on speed considerations
'Note that no byte is changed in source regardless if BSTR or SAFEARRAY is used
'*******************************************************************************
Private Sub CopyBytes(ByVal bytesCount As Long _
                    , ByRef rmSrc As REMOTE_MEMORY, ByRef vtSrc As Variant _
                    , ByRef rmDest As REMOTE_MEMORY, ByRef vtDest As Variant _
                    , ByRef destValue As Variant, ByRef sArr As SAFEARRAY_1D _
                    , ByRef arrBytes As Variant, ByRef vtArr As Variant _
                    , ByRef rmBSTR As REMOTE_MEMORY, ByRef vtBSTR As Variant _
                    , ByRef bstrPtrValue As Variant)
    Const bstrPrefixSize As Long = 4
    Dim bytes As Long: bytes = bytesCount - bstrPrefixSize
    Dim bstrLength As Long
    Dim s As String 'Must not be Variant so that LSet is faster
    Dim destString As String
    Dim useBSTR As Boolean
    Dim hasOverlap As Boolean
    Dim overlapBSTRLen As Long
    Dim overlapOffset As LongPtr
    '
    vtSrc = vbLong + VT_BYREF
    bstrLength = rmSrc.memValue 'Copy first 4 bytes froum source
    vtSrc = vbLongPtr
    '
    Const maxMidBs As Long = 2 ^ 5 'Use SAFEARRAY and MidB below this value
    useBSTR = (bstrLength >= bytes Or bstrLength < 0) And bytes > maxMidBs
    If useBSTR Then 'Prepare source BSTR
        rmBSTR.memValue = VarPtr(s)
        #If Win64 Then
            Const curBSTRPrefixSize As Currency = 0.0004
            vtSrc = vbCurrency
            vtBSTR = vbCurrency + VT_BYREF
            bstrPtrValue = rmSrc.memValue + curBSTRPrefixSize
            vtSrc = vbLongPtr
        #Else
            vtBSTR = vbLong + VT_BYREF
            bstrPtrValue = rmSrc.memValue + bstrPrefixSize
        #End If
    Else 'Prepare source SAFEARRAY
        sArr.pvData = rmSrc.memValue + bstrPrefixSize
        sArr.rgsabound0.cElements = bytes
        vtArr = vbArray + vbByte
    End If
    '
    'Prepare destination BSTR
    If rmDest.memValue > rmSrc.memValue Then
        hasOverlap = UnsignedAdd(rmSrc.memValue, bytes + bstrPrefixSize) > rmDest.memValue
        If hasOverlap Then overlapOffset = rmDest.memValue - rmSrc.memValue
    Else
        hasOverlap = False
    End If
    vtDest = vbLong + VT_BYREF
    If hasOverlap Then overlapBSTRLen = destValue
    destValue = bytes
    vtDest = vbLongPtr
    rmDest.memValue = rmDest.memValue + bstrPrefixSize
    vtDest = vbString
    '
    'Copy and clean
    If useBSTR Then
        #If TWINBASIC Then
            MemLongPtr(VarPtr(destString)) = StrPtr(destValue)
            MemLongPtr(VarPtr(s)) = rmSrc.memValue + bstrPrefixSize
            LSet destString = s
            If bytes Mod 2 = 1 Then
                MidB(destString, bytes, 1) = MidB$(s, bytes, 1)
            End If
            MemLongPtr(VarPtr(destString)) = CLngPtr(0)
            MemLongPtr(VarPtr(s)) = CLngPtr(0)
        #Else
            LSet destValue = s 'LSet cannot copy an odd number of bytes
            If bytes Mod 2 = 1 Then
                MidB(destValue, bytes, 1) = MidB$(s, bytes, 1)
            End If
        #End If
        bstrPtrValue = 0
        vtBSTR = vbEmpty
    Else
        Const maxMidBa As Long = maxMidBs * 2 ^ 3
        If bytes > maxMidBa Then
            #If TWINBASIC Then
                MemLongPtr(VarPtr(destString)) = StrPtr(destValue)
                LSet destString = arrBytes
                If bytes Mod 2 = 1 Then
                    Static lastByte(0 To 0) As Byte
                    lastByte(0) = arrBytes(UBound(arrBytes))
                    MidB(destString, bytes, 1) = lastByte
                End If
                MemLongPtr(VarPtr(destString)) = CLngPtr(0)
            #Else
                LSet destValue = arrBytes
                If bytes Mod 2 = 1 Then
                    Static lastByte(0 To 0) As Byte
                    lastByte(0) = arrBytes(UBound(arrBytes))
                    MidB(destValue, bytes, 1) = lastByte
                End If
            #End If
        Else
            #If TWINBASIC Then
                MemLongPtr(VarPtr(destString)) = StrPtr(destValue)
                MidB(destString, 1) = arrBytes
                MemLongPtr(VarPtr(destString)) = CLngPtr(0)
            #Else
                MidB(destValue, 1) = arrBytes
            #End If
        End If
        vtArr = vbEmpty
    End If
    '
    vtDest = vbLongPtr
    rmDest.memValue = rmDest.memValue - bstrPrefixSize
    vtDest = vbLong + VT_BYREF
    destValue = bstrLength 'Copy the correct 'BSTR length' bytes
    vtDest = vbLongPtr
    If hasOverlap Then
        rmDest.memValue = rmDest.memValue + overlapOffset
        vtDest = vbLong + VT_BYREF
        destValue = overlapBSTRLen
        vtDest = vbLongPtr
    End If
End Sub

'*******************************************************************************
'Copy a param array to another array of Variants while preserving ByRef elements
'*******************************************************************************
Public Sub CloneParamArray(ByRef firstElem As Variant _
                         , ByVal elemCount As Long _
                         , ByRef outArray() As Variant)
    ReDim outArray(0 To elemCount - 1)
    MemCopy VarPtr(outArray(0)), VarPtr(firstElem), VARIANT_SIZE * elemCount
    '
    Static sArr As SAFEARRAY_1D 'Fake array of VarTypes (Integers)
    Static rmArr As REMOTE_MEMORY
    '
    If Not rmArr.isInitialized Then
        With sArr
            .cDims = 1
            .fFeatures = FADF_HAVEVARTYPE
            .cbElements = INT_SIZE
        End With
        InitRemoteMemory rmArr
        rmArr.memValue = VarPtr(sArr)
    End If
    sArr.rgsabound0.cElements = elemCount * VT_SPACING
    sArr.pvData = VarPtr(outArray(0))
    '
    FixByValElements outArray, rmArr, rmArr.remoteVT
End Sub

'*******************************************************************************
'Utility for 'CloneParamArray' - avoid deallocation on elements passed ByVal
'e.g. if original ParamArray has a pointer to a BSTR then safely clear the copy
'*******************************************************************************
Private Sub FixByValElements(ByRef arr() As Variant _
                           , ByRef rmArr As REMOTE_MEMORY _
                           , ByRef vtArr As Variant)
    Dim i As Long
    Dim v As Variant
    Dim vtIndex As Long: vtIndex = 0
    Dim vt As VbVarType
    '
    vtArr = vbArray + vbInteger
    For i = 0 To UBound(arr)
        vt = rmArr.memValue(vtIndex)
        If (vt And VT_BYREF) = 0 Then
            If (vt And vbArray) = vbArray Or vt = vbObject Or vt = vbString _
            Or vt = vbDataObject Or vt = vbUserDefinedType Then
                If vt = vbObject Then Set v = arr(i) Else v = arr(i)
                rmArr.memValue(vtIndex) = vbEmpty 'Avoid deallocation
                If vt = vbObject Then Set arr(i) = v Else arr(i) = v
            End If
        End If
        vtIndex = vtIndex + VT_SPACING
    Next i
    vtArr = vbEmpty
End Sub

'*******************************************************************************
'Returns the input array wrapped in a ByRef Variant without copying the array
'*******************************************************************************
Public Function GetArrayByRef(ByRef arr As Variant) As Variant
    If IsArray(arr) Then
        GetArrayByRef = VarPtrArr(arr)
        MemInt(VarPtr(GetArrayByRef)) = VarType(arr) Or VT_BYREF
    Else
        Err.Raise 5, "GetArrayByRef", "Array required"
    End If
End Function

'*******************************************************************************
'Reads the memory of a String to an Array of Integers
'Notes:
'   - Ignores the last byte if input has an odd number of bytes
'   - If 'outLength' is -1 (default) then the remaining length is returned
'   - Excess length is ignored
'*******************************************************************************
Public Function StringToIntegers(ByRef s As String _
                               , Optional ByVal startIndex As Long = 1 _
                               , Optional ByVal outLength As Long = -1 _
                               , Optional ByVal outLowBound As Long = 0) As Integer()
    Static rm As REMOTE_MEMORY
    Static sArr As SAFEARRAY_1D
    Const methodName As String = "StringToIntegers"
    Dim cLen As Long: cLen = Len(s)

    If startIndex < 1 Then
        Err.Raise 9, methodName, "Invalid Start Index"
    ElseIf outLength < -1 Then
        Err.Raise 5, methodName, "Invalid Length for output"
    ElseIf outLength = -1 Or startIndex + outLength - 1 > cLen Then
        outLength = cLen - startIndex + 1 'Excess length is ignored
        If outLength < 0 Then outLength = 0
    End If
    '
    If Not rm.isInitialized Then
        InitRemoteMemory rm
        With sArr
            .cDims = 1
            .fFeatures = FADF_HAVEVARTYPE
            .cbElements = INT_SIZE
        End With
    End If
    '
    sArr.pvData = StrPtr(s) + (startIndex - 1) * INT_SIZE
    sArr.rgsabound0.lLbound = outLowBound
    sArr.rgsabound0.cElements = outLength
    '
    RemoteAssign rm, VarPtr(sArr), rm.remoteVT, vbArray + vbInteger, StringToIntegers, rm.memValue
End Function

'*******************************************************************************
'Reads the memory of an Array of Integers into a String
'Notes:
'   - If 'outLength' is -1 (default) then the remaining length is returned
'   - Excess length is ignored
'*******************************************************************************
Public Function IntegersToString(ByRef ints() As Integer _
                               , Optional ByVal startIndex As Long = 0 _
                               , Optional ByVal outLength As Long = -1) As String
    Static rm As REMOTE_MEMORY
    Static sArr As SAFEARRAY_1D
    Const methodName As String = "IntegersToString"

    If GetArrayDimsCount(ints) <> 1 Then
        Err.Raise 5, methodName, "Expected 1D Array of Integers"
    ElseIf startIndex < LBound(ints) Then
        Err.Raise 9, methodName, "Invalid Start Index"
    ElseIf outLength < -1 Then
        Err.Raise 5, methodName, "Invalid Length for output"
    ElseIf outLength = -1 Or startIndex + outLength - 1 > UBound(ints) Then
        outLength = UBound(ints) - startIndex + 1
        If outLength < 0 Then Exit Function
    End If
    If Not rm.isInitialized Then
        InitRemoteMemory rm
        With sArr
            .cDims = 1
            .fFeatures = FADF_HAVEVARTYPE
            .cbElements = BYTE_SIZE
            .rgsabound0.lLbound = 0
        End With
    End If
    '
    sArr.pvData = VarPtr(ints(startIndex))
    sArr.rgsabound0.cElements = outLength * INT_SIZE
    '
    RemoteAssign rm, VarPtr(sArr), rm.remoteVT, vbArray + vbByte, IntegersToString, rm.memValue
End Function

'*******************************************************************************
'Returns an empty array of the requested size and type
'*******************************************************************************
Public Function EmptyArray(ByVal numberOfDimensions As Long _
                         , ByVal vType As VbVarType) As Variant
    Const methodName As String = "EmptyArray"
    Const MAX_DIMENSION As Long = 60
    '
    If numberOfDimensions < 1 Or numberOfDimensions > MAX_DIMENSION Then
        Err.Raise 5, methodName, "Invalid number of dimensions"
    End If
    Select Case vType
    Case vbByte, vbInteger, vbLong, vbLongLong 'Integers
    Case vbCurrency, vbDecimal, vbDouble, vbSingle, vbDate 'Decimal-point
    Case vbBoolean, vbString, vbObject, vbDataObject, vbVariant 'Other
    Case Else
        Err.Raise 13, methodName, "Type mismatch"
    End Select
    '
    Static fakeSafeArray() As Long
    Static rm As REMOTE_MEMORY
    #If Win64 Then
        Const safeArraySize = 6
    #Else
        Const safeArraySize = 4
    #End If
    Const FADF_HAVEVARTYPE As Long = &H80
    Const fFeaturesHi As Long = FADF_HAVEVARTYPE * &H10000
    Dim i As Long
    '
    If Not rm.isInitialized Then
        InitRemoteMemory rm
        ReDim fakeSafeArray(0 To safeArraySize + MAX_DIMENSION * 2 - 1)
        fakeSafeArray(1) = 1 'cbElements member - needs to be non-zero
        rm.memValue = VarPtr(fakeSafeArray(0)) 'The fake ArrPtr
        '
        'Set 'cElements' to 1 for each SAFEARRAYBOUND
        For i = safeArraySize To UBound(fakeSafeArray, 1) Step 2
            fakeSafeArray(i) = 1
        Next i
    End If
    fakeSafeArray(0) = fFeaturesHi + numberOfDimensions 'cDims and fFeatures
    i = safeArraySize + (numberOfDimensions - 1) * 2 'Highest dimension position
    '
    fakeSafeArray(i) = 0 'Highest dimension must have 0 'cElements'
    RemoteAssign rm, VarPtr(fakeSafeArray(0)), rm.remoteVT, vbArray + vType, EmptyArray, rm.memValue
    fakeSafeArray(i) = 1
End Function

'*******************************************************************************
'Allows user to update the LBound Index for an array dimension
'*******************************************************************************
Public Sub UpdateLBound(ByRef arr As Variant _
                      , ByVal dimension As Long _
                      , ByVal newLB As Long)
    #If Win64 Then
        Const bOffset As Long = 28
    #Else
        Const bOffset As Long = 20
    #End If
    Const methodName As String = "UpdateLBound"
    Const maxL As Long = &H7FFFFFFF
    Dim dimensionCount As Long: dimensionCount = GetArrayDimsCount(arr)
    '
    If dimension < 1 Or dimension > dimensionCount Then
        Err.Raise 5, methodName, "Invalid dimension or not array"
    ElseIf maxL - UBound(arr, dimension) + LBound(arr, dimension) < newLB Then
        Err.Raise 5, methodName, "New bound overflow"
    End If
    MemLong(ArrPtr(arr) + bOffset + (dimensionCount - dimension) * 8) = newLB
End Sub

'*******************************************************************************
'Returns the Number of dimensions for an input array
'Returns 0 if array is uninitialized or input not an array
'Note that a zero-length array has 1 dimension! Ex. Array() bounds are (0 to -1)
'*******************************************************************************
Private Function GetArrayDimsCount(ByRef arr As Variant) As Long
    Const MAX_DIMENSION As Long = 60 'VB limit
    Dim dimension As Long
    Dim tempBound As Long
    '
    On Error GoTo FinalDimension
    For dimension = 1 To MAX_DIMENSION
        tempBound = LBound(arr, dimension)
    Next dimension
FinalDimension:
    GetArrayDimsCount = dimension - 1
End Function

'*******************************************************************************
'Fills target memory with zero
'*******************************************************************************
Public Sub MemZero(ByVal destinationPtr As LongPtr, ByVal bytesCount As LongPtr)
    MemFill destinationPtr, bytesCount, 0
End Sub

'*******************************************************************************
'Fills target memory with the specified Byte value
'*******************************************************************************
Public Sub MemFill(ByVal destinationPtr As LongPtr _
                 , ByVal bytesCount As LongPtr _
                 , ByVal fillByte As Byte)
#If Mac Then
    FillMemory ByVal destinationPtr, fillByte, bytesCount
#ElseIf (VBA7 = 0) Then
    FillMemory ByVal destinationPtr, bytesCount, fillByte
#Else
    If bytesCount = 0 Then Exit Sub
    Const maxSizeSpeedGain As Long = &H100000 'Beyond this use API directly
    If bytesCount < 0 Or bytesCount > maxSizeSpeedGain Then
        FillMemory ByVal destinationPtr, bytesCount, fillByte
        Exit Sub
    End If
    '
    Const maxSizeMidB As Long = &H2000 'Beyond this use MemCopy
    Dim bytesLeft As Long
    Dim bytes As Long
    Dim chunk As Long
    Static rm As REMOTE_MEMORY
    If Not rm.isInitialized Then InitRemoteMemory rm
    '
    If bytesCount > maxSizeMidB Then
        bytes = maxSizeMidB
        bytesLeft = CLng(bytesCount) - maxSizeMidB
        chunk = maxSizeMidB
    Else
        bytes = CLng(bytesCount)
    End If
    FillBytes destinationPtr, bytes, fillByte, rm.remoteVT, rm.memValue
    '
    Do While bytesLeft > 0
        If chunk > bytesLeft Then bytes = bytesLeft Else bytes = chunk
        MemCopy destinationPtr + chunk, destinationPtr, bytes
        bytesLeft = bytesLeft - bytes
        chunk = chunk * 2
    Loop
#End If
End Sub
Private Sub FillBytes(ByVal destinationPtr As LongPtr _
                    , ByVal bytesCount As Long _
                    , ByVal fillByte As Byte _
                    , ByRef vt As Variant _
                    , ByRef rmValue As Variant)
    rmValue = destinationPtr
    If bytesCount > 5 Then
        Const bstrPrefixSize As Long = 4
        Dim s As String
        '
        vt = vbLong + VT_BYREF
        rmValue = bytesCount - bstrPrefixSize 'Write BSTR LenB
        '
        vt = vbLongPtr
        rmValue = destinationPtr + bstrPrefixSize 'Move address to byte 5
        '
        #If Win64 Then
            Dim bstrPtr As Currency: vt = vbCurrency
        #Else
            Dim bstrPtr As Long:     vt = vbLong
        #End If
        bstrPtr = rmValue
        '
        vt = vbByte + VT_BYREF
        rmValue = fillByte 'Copy fill byte to byte 5
        '
        vt = vbLongPtr
        rmValue = VarPtr(s)
        '
        #If Win64 Then
            vt = vbCurrency + VT_BYREF
        #Else
            vt = vbLong + VT_BYREF
        #End If
        rmValue = bstrPtr 'Redirect to fake BSTR
        '
        MidB$(s, 2) = s 'The actual fill
        rmValue = 0     'Clear BSTR
        '
        vt = vbLongPtr
        rmValue = destinationPtr
    Else
        Dim extraByte As Boolean: extraByte = (bytesCount Mod 2) = 1
    End If
    '
    Select Case bytesCount
    Case 1
    Case 2, 3
        vt = vbInteger + VT_BYREF
        If fillByte And &H80 Then
            rmValue = fillByte Or (fillByte Xor &H80) * &H100 Or &H8000
        Else
            rmValue = fillByte * &H101
        End If
        If extraByte Then
            vt = vbLongPtr
            rmValue = rmValue + 2
        End If
    Case Else
        vt = vbLong + VT_BYREF
        If fillByte And &H80 Then
            rmValue = fillByte * &H10101 Or (fillByte Xor &H80) _
                    * &H1000000 Or &H80000000
        Else
            rmValue = fillByte * &H1010101
        End If
        If extraByte Then
            vt = vbLongPtr
            rmValue = rmValue + 4
        End If
    End Select
    If extraByte Then
        vt = vbByte + VT_BYREF
        rmValue = fillByte
    End If
    vt = vbEmpty
End Sub
