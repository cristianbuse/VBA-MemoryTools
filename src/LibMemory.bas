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
        Public Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
    #Else
        Public Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
    #End If
#Else 'Windows
    'https://msdn.microsoft.com/en-us/library/mt723419(v=vs.85).aspx
    #If VBA7 Then
        Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    #Else
        Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    #End If
#End If

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

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY_1D
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
#If Win64 Then
Private Property Let MemIntAPI(ByVal memAddress As LongLong, ByVal newValue As Integer)
#Else
Private Property Let MemIntAPI(ByVal memAddress As Long, ByVal newValue As Integer)
#End If
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
#If Win64 Then
Private Sub RemoteAssign(ByRef rm As REMOTE_MEMORY _
                       , ByRef memAddress As LongLong _
                       , ByRef remoteVT As Variant _
                       , ByVal newVT As VbVarType _
                       , ByRef targetVariable As Variant _
                       , ByRef newValue As Variant)
#Else
Private Sub RemoteAssign(ByRef rm As REMOTE_MEMORY _
                       , ByRef memAddress As Long _
                       , ByRef remoteVT As Variant _
                       , ByVal newVT As VbVarType _
                       , ByRef targetVariable As Variant _
                       , ByRef newValue As Variant)
#End If
    rm.memValue = memAddress
    If Not rm.isInitialized Then InitRemoteMemory rm
    remoteVT = newVT
    targetVariable = newValue
    remoteVT = vbEmpty 'Stop linking to remote address, for safety
End Sub

'*******************************************************************************
'Read/Write a Byte from/to memory
'*******************************************************************************
#If Win64 Then
Public Property Get MemByte(ByVal memAddress As LongLong) As Byte
#Else
Public Property Get MemByte(ByVal memAddress As Long) As Byte
#End If
    #If Mac Then
        CopyMemory MemByte, ByVal memAddress, 1
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbByte + VT_BYREF, MemByte, rm.memValue
    #End If
End Property
#If Win64 Then
Public Property Let MemByte(ByVal memAddress As LongLong, ByVal newValue As Byte)
#Else
Public Property Let MemByte(ByVal memAddress As Long, ByVal newValue As Byte)
#End If
    #If Mac Then
        CopyMemory ByVal memAddress, newValue, 1
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbByte + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 2 Bytes (Integer) from/to memory
'*******************************************************************************
#If Win64 Then
Public Property Get MemInt(ByVal memAddress As LongLong) As Integer
#Else
Public Property Get MemInt(ByVal memAddress As Long) As Integer
#End If
    #If Mac Then
        CopyMemory MemInt, ByVal memAddress, 2
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbInteger + VT_BYREF, MemInt, rm.memValue
    #End If
End Property

#If Win64 Then
Public Property Let MemInt(ByVal memAddress As LongLong, ByVal newValue As Integer)
#Else
Public Property Let MemInt(ByVal memAddress As Long, ByVal newValue As Integer)
#End If
    #If Mac Then
        CopyMemory ByVal memAddress, newValue, 2
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbInteger + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 2 Bytes (Boolean) from/to memory
'*******************************************************************************
#If Win64 Then
Public Property Get MemBool(ByVal memAddress As LongLong) As Boolean
#Else
Public Property Get MemBool(ByVal memAddress As Long) As Boolean
#End If
    #If Mac Then
        CopyMemory MemBool, ByVal memAddress, 2
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbBoolean + VT_BYREF, MemBool, rm.memValue
    #End If
End Property
#If Win64 Then
Public Property Let MemBool(ByVal memAddress As LongLong, ByVal newValue As Boolean)
#Else
Public Property Let MemBool(ByVal memAddress As Long, ByVal newValue As Boolean)
#End If
    #If Mac Then
        CopyMemory ByVal memAddress, newValue, 2
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbBoolean + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 4 Bytes (Long) from/to memory
'*******************************************************************************
#If Win64 Then
Public Property Get MemLong(ByVal memAddress As LongLong) As Long
#Else
Public Property Get MemLong(ByVal memAddress As Long) As Long
#End If
    #If Mac Then
        CopyMemory MemLong, ByVal memAddress, 4
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbLong + VT_BYREF, MemLong, rm.memValue
    #End If
End Property
#If Win64 Then
Public Property Let MemLong(ByVal memAddress As LongLong, ByVal newValue As Long)
#Else
Public Property Let MemLong(ByVal memAddress As Long, ByVal newValue As Long)
#End If
    #If Mac Then
        CopyMemory ByVal memAddress, newValue, 4
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbLong + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 4 Bytes (Single) from/to memory
'*******************************************************************************
#If Win64 Then
Public Property Get MemSng(ByVal memAddress As LongLong) As Single
#Else
Public Property Get MemSng(ByVal memAddress As Long) As Single
#End If
    #If Mac Then
        CopyMemory MemSng, ByVal memAddress, 4
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbSingle + VT_BYREF, MemSng, rm.memValue
    #End If
End Property
#If Win64 Then
Public Property Let MemSng(ByVal memAddress As LongLong, ByVal newValue As Single)
#Else
Public Property Let MemSng(ByVal memAddress As Long, ByVal newValue As Single)
#End If
    #If Mac Then
        CopyMemory ByVal memAddress, newValue, 4
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbSingle + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 8 Bytes (LongLong) from/to memory
'*******************************************************************************
#If Win64 Then
Public Property Get MemLongLong(ByVal memAddress As LongLong) As LongLong
    #If Mac Then
        CopyMemory MemLongLong, ByVal memAddress, 8
    #Else
        'Cannot set Variant/LongLong ByRef so we cannot use 'RemoteAssign'
        Static rm As REMOTE_MEMORY: rm.memValue = memAddress
        MemLongLong = ByRefLongLong(rm, rm.remoteVT, rm.memValue)
    #End If
End Property
Public Property Let MemLongLong(ByVal memAddress As LongLong, ByVal newValue As LongLong)
    #If Mac Then
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
#If Win64 Then
Public Property Get MemLongPtr(ByVal memAddress As LongLong) As LongLong
#Else
Public Property Get MemLongPtr(ByVal memAddress As Long) As Long
#End If
    #If Mac Then
        CopyMemory MemLongPtr, ByVal memAddress, PTR_SIZE
    #ElseIf Win64 Then
        Static rm As REMOTE_MEMORY: rm.memValue = memAddress
        MemLongPtr = ByRefLongLong(rm, rm.remoteVT, rm.memValue)
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbLong + VT_BYREF, MemLongPtr, rm.memValue
    #End If
End Property
#If Win64 Then
Public Property Let MemLongPtr(ByVal memAddress As LongLong, ByVal newValue As LongLong)
#Else
Public Property Let MemLongPtr(ByVal memAddress As Long, ByVal newValue As Long)
#End If
    #If Mac Then
        CopyMemory ByVal memAddress, newValue, PTR_SIZE
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
#If Win64 Then
Public Property Get MemCur(ByVal memAddress As LongLong) As Currency
#Else
Public Property Get MemCur(ByVal memAddress As Long) As Currency
#End If
    #If Mac Then
        CopyMemory MemCur, ByVal memAddress, 8
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbCurrency + VT_BYREF, MemCur, rm.memValue
    #End If
End Property
#If Win64 Then
Public Property Let MemCur(ByVal memAddress As LongLong, ByVal newValue As Currency)
#Else
Public Property Let MemCur(ByVal memAddress As Long, ByVal newValue As Currency)
#End If
    #If Mac Then
        CopyMemory ByVal memAddress, newValue, 8
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbCurrency + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 8 Bytes (Date) from/to memory
'*******************************************************************************
#If Win64 Then
Public Property Get MemDate(ByVal memAddress As LongLong) As Date
#Else
Public Property Get MemDate(ByVal memAddress As Long) As Date
#End If
    #If Mac Then
        CopyMemory MemDate, ByVal memAddress, 8
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbDate + VT_BYREF, MemDate, rm.memValue
    #End If
End Property
#If Win64 Then
Public Property Let MemDate(ByVal memAddress As LongLong, ByVal newValue As Date)
#Else
Public Property Let MemDate(ByVal memAddress As Long, ByVal newValue As Date)
#End If
    #If Mac Then
        CopyMemory ByVal memAddress, newValue, 8
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbDate + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Read/Write 8 Bytes (Double) from/to memory
'*******************************************************************************
#If Win64 Then
Public Property Get MemDbl(ByVal memAddress As LongLong) As Double
#Else
Public Property Get MemDbl(ByVal memAddress As Long) As Double
#End If
    #If Mac Then
        CopyMemory MemDbl, ByVal memAddress, 8
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbDouble + VT_BYREF, MemDbl, rm.memValue
    #End If
End Property
#If Win64 Then
Public Property Let MemDbl(ByVal memAddress As LongLong, ByVal newValue As Double)
#Else
Public Property Let MemDbl(ByVal memAddress As Long, ByVal newValue As Double)
#End If
    #If Mac Then
        CopyMemory ByVal memAddress, newValue, 8
    #Else
        Static rm As REMOTE_MEMORY
        RemoteAssign rm, memAddress, rm.remoteVT, vbDouble + VT_BYREF, rm.memValue, newValue
    #End If
End Property

'*******************************************************************************
'Dereference an object by it's pointer
'*******************************************************************************
#If Win64 Then
Public Function MemObj(ByVal memAddress As LongLong) As Object
#Else
Public Function MemObj(ByVal memAddress As Long) As Object
#End If
    If memAddress = 0 Then Exit Function
    '
    #If Mac Then
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
#If Win64 Then
Public Function UnsignedAdd(ByVal unsignedPtr As LongLong, ByVal signedOffset As LongLong) As LongLong
    UnsignedAdd = ((unsignedPtr Xor &H8000000000000000^) + signedOffset) Xor &H8000000000000000^
End Function
#Else
Public Function UnsignedAdd(ByVal unsignedPtr As Long, ByVal signedOffset As Long) As Long
    UnsignedAdd = ((unsignedPtr Xor &H80000000) + signedOffset) Xor &H80000000
End Function
#End If

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
#If Win64 Then
Public Sub RedirectInstance(ByVal funcReturnPtr As LongLong _
                          , ByVal currentInstance As Object _
                          , ByVal targetInstance As Object)
#Else
Public Sub RedirectInstance(ByVal funcReturnPtr As Long _
                          , ByVal currentInstance As Object _
                          , ByVal targetInstance As Object)
#End If
    Const methodName As String = "RedirectInstance"
    #If Win64 Then
        Dim originalPtr As LongLong
        Dim newPtr As LongLong
        Dim swapAddress As LongLong
    #Else
        Dim originalPtr As Long
        Dim newPtr As Long
        Dim swapAddress As Long
    #End If
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
'Returns error 5 for a non-array or an array wrapped in a Variant
'*******************************************************************************
#If Win64 Then
Public Function VarPtrArr(ByRef arr As Variant) As LongLong
#Else
Public Function VarPtrArr(ByRef arr As Variant) As Long
#End If
    Const vtArrByRef As Long = vbArray + VT_BYREF
    Dim vt As VbVarType: vt = MemInt(VarPtr(arr)) 'VarType(arr) ignores VT_BYREF
    If (vt And vtArrByRef) = vtArrByRef Then
        Const pArrayOffset As Long = 8
        VarPtrArr = MemLongPtr(VarPtr(arr) + pArrayOffset)
    Else
        Err.Raise 5, "VarPtrArr", "Array required"
    End If
End Function

'*******************************************************************************
'Returns the pointer to the underlying SAFEARRAY structure of a VB array
'Returns error 5 for a non-array
'*******************************************************************************
#If Win64 Then
Public Function ArrPtr(ByRef arr As Variant) As LongLong
#Else
Public Function ArrPtr(ByRef arr As Variant) As Long
#End If
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
'Win - bytesCount 1 to 2147483647 - no API calls. Uses a combination of
'      REMOTE_MEMORY/SAFEARRAY_1D structs as well as native Strings and Arrays
'      to manipulate memory. Works within size limitation of Strings in VBA
'      For some smaller sizes (<=5) optimizes via MemLong, MemInt, MemByte etc.
'    - bytesCount < 0 or > 2147483647 - wrapper around CopyMemory/RtlMoveMemory
'*******************************************************************************
#If Win64 Then
Public Sub MemCopy(ByVal destinationPtr As LongLong _
                 , ByVal sourcePtr As LongLong _
                 , ByVal bytesCount As LongLong)
#Else
Public Sub MemCopy(ByVal destinationPtr As Long _
                 , ByVal sourcePtr As Long _
                 , ByVal bytesCount As Long)
#End If
#If Mac Then
    CopyMemory ByVal destinationPtr, ByVal sourcePtr, bytesCount
#Else
    #If Win64 Then
        Const maxLong As Long = &H7FFFFFFF
        If bytesCount < 0 Or bytesCount > maxLong Then
    #Else
        If bytesCount < 0 Then
    #End If
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
    Dim tempSize As Long
    Dim useBSTR As Boolean
    '
    Do
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
            Const maxStartMidB As Long = 2 ^ 30 'MidB second param limit (bug)
            If bytes > maxStartMidB And bytes Mod 2 = 1 Then
                tempSize = maxStartMidB
                bytes = bytes - maxStartMidB
            Else
                tempSize = bytes
                bytes = 0
            End If
        Else 'Prepare source SAFEARRAY
            'For large amounts it is faster to copy memory in smaller chunks
            Const chunkSize As Long = 2 ^ 16 'Similar performance with 2 ^ 17
            '
            If bytes > chunkSize + bstrPrefixSize + 1 Then
                tempSize = chunkSize
                bytes = bytes - chunkSize - bstrPrefixSize
            Else
                tempSize = bytes
                bytes = 0
            End If
            sArr.pvData = rmSrc.memValue + bstrPrefixSize
            sArr.rgsabound0.cElements = tempSize
            vtArr = vbArray + vbByte
        End If
        '
        'Prepare destination BSTR
        vtDest = vbLong + VT_BYREF
        destValue = tempSize
        vtDest = vbLongPtr
        rmDest.memValue = rmDest.memValue + bstrPrefixSize
        vtDest = vbString
        '
        'Copy and clean
        If useBSTR Then
            LSet destValue = s 'LSet cannot copy an odd number of bytes
            If tempSize Mod 2 = 1 Then
                MidB(destValue, tempSize, 1) = MidB$(s, tempSize, 1)
            End If
            bstrPtrValue = 0
            vtBSTR = vbEmpty
        Else
            Const maxMidBa As Long = maxMidBs * 2 ^ 3
            If tempSize > maxMidBa Then
                LSet destValue = arrBytes
                If tempSize Mod 2 = 1 Then
                    Static lastByte(0 To 0) As Byte
                    lastByte(0) = arrBytes(UBound(arrBytes))
                    MidB(destValue, tempSize, 1) = lastByte
                End If
            Else
                MidB(destValue, 1) = arrBytes
            End If
            vtArr = vbEmpty
        End If
        '
        vtDest = vbLongPtr
        rmDest.memValue = rmDest.memValue - bstrPrefixSize
        vtDest = vbLong + VT_BYREF
        destValue = bstrLength 'Copy the correct 'BSTR length' bytes
        vtDest = vbLongPtr
        '
        If bytes > 0 Then 'Advance address for next chunk
            Dim bytesOffset As Long: bytesOffset = chunkSize + bstrPrefixSize
            rmDest.memValue = UnsignedAdd(rmDest.memValue, bytesOffset)
            rmSrc.memValue = UnsignedAdd(rmSrc.memValue, bytesOffset)
        End If
    Loop Until bytes = 0
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
