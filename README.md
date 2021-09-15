# VBA-MemoryTools
Native memory manipulation in VBA

Using CopyMemory API (RtlMoveMemory on Windows) is quite slow when used many times. Moreover, on some systems this Memory API is even slower due to certain software (e.g. Windows Defender - see [article](https://stackoverflow.com/questions/57885185/windows-defender-extremly-slowing-down-macro-only-on-windows-10)). The API can become so slow that is pretty much unusable (e.g. on my x32 Windows machine it is 600 times slower than it used to be). Using the **LibMemory** module presented here overcomes the speed issues for reading and writing 1, 2, 4 and 8 bytes from and into memory.

Related [Code Review question](https://codereview.stackexchange.com/questions/252659/fast-native-memory-manipulation-in-vba)

## Implementation
Same technique used [here](https://codereview.stackexchange.com/a/249125/227582) was implemented. A remote Variant allows the changing of the VarType on a second Variant which in turn reads memory remotely as well (has VT_BYREF flag set). A single CopyMemory API call is done when initializing the mentioned remote VarType. Subsequent usage relies on native VBA code only.

5 main parametric properties (Get/Let) are exposed:
 1. MemByte
 2. MetInt 
 3. MemLong
 4. MemLongPtr
 5. MemLongLong (x64 only)

A few other utilities:
 - GetDefaultInterface
 - MemObject (dereferences a pointer and returns an Object)
 - UnsignedAddition

See ```DemoMain``` method in the Demo module.

## Class Instance Redirection

Class instance redirection is now supported. This allows Private Class Initializers thus achieving true immutabilty.
Simply call the ```RedirectInstance``` method within a ```Private Function``` of any VB class to gain access to other instances of the same class.
Related [Code Review question](https://codereview.stackexchange.com/questions/253233/private-vba-class-initializer-called-from-factory-2).

See ```DemoInstanceRedirection``` method in the Demo module.

## Installation
Just import the following code modules in your VBA Project:
* [**LibMemory.bas**](https://github.com/cristianbuse/VBA-MemoryTools/blob/master/src/LibMemory.bas)

## Demo
Import the following code modules from the [demo folder](https://github.com/cristianbuse/VBA-MemoryTools/blob/master/src/Demo) in your VBA Project:
* [DemoLibMemory.bas](https://github.com/cristianbuse/VBA-MemoryTools/blob/master/src/Demo/DemoLibMemory.bas) - run ```DemoMain```
* [DemoClass](https://github.com/cristianbuse/VBA-MemoryTools/blob/master/src/Demo/DemoClass.cls)

## Notes
* CopyMemory API is also exposed just in case the main methods are not satisfying the requirement (e.g. copy 50 bytes at once)

## License
MIT License

Copyright (c) 2020 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
