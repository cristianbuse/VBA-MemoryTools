# VBA-MemoryTools
Native memory manipulation in VBA.

There is an issue with the speed of API calls in **VBA7**. This is very well tested and explained in [this Code Review question](https://codereview.stackexchange.com/questions/270258/evaluate-performance-of-dll-calls-from-vba).

This library overcomes the speed issues for reading and writing from and into memory by using a native approach - see related [CR question](https://codereview.stackexchange.com/questions/252659/fast-native-memory-manipulation-in-vba).

Moreover, this library exposes some useful utilities and wrappers to make it easier to manipulate memory. For **Mac**, **TwinBasic** and **VBA6** (and prior) this library simply uses wrapped API calls as there is no speed benefit in using the native approach.

Copying a byte for 10,000 times on Windows with VBA7 x64 using ```RtlMoveMemory``` API takes around 10 seconds while the native By Ref approach takes only around 16 milliseconds for the same number of iterations. So, the speed gain is 600x is some cases.

## Implementation
Same technique used [here](https://codereview.stackexchange.com/a/249125/227582) was implemented. A remote ```Variant``` allows the changing of the ```VarType``` on a second ```Variant``` which in turn reads memory remotely as well (has ```VT_BYREF``` flag set). A single CopyMemory API call is done when initializing the base REMOTE_MEMORY structure (see ```MemIntAPI```). Subsequent usage relies on native VBA code only.

## Use
```MemCopy``` - a faster alternative to ```CopyMemory``` (for VBA7) without API calls on Windows up to sizes of 16777216. Uses a combination of fake BSTR and SAFEARRAY structures to copy memory.

10 parametric properties (Get/Let) are exposed:
 01. ```MemByte```
 02. ```MetInt```
 03. ```MemLong```
 04. ```MemLongPtr```
 05. ```MemLongLong``` (x64 only)
 06. ```MemBool```
 07. ```MemSng``` 
 08. ```MemCur```
 09. ```MemDate```
 10. ```MemDbl```

A few other utilities:
 - ```GetDefaultInterface```
 - ```MemObject``` (dereferences a pointer and returns an Object)
 - ```UnsignedAddition```
 - ```VarPtrArr``` (```VarPtr``` for arrays)
 - ```ArrPtr``` (as ```ObjPtr``` is for objects and ```StrPtr``` is for strings) - returns the pointer to the underlying SAFEARRAY structure
 - ```CloneParamArray``` - copies a param array to another array of Variants while preserving ByRef elements
 - ```GetArrayByRef``` - returns the input array wrapped in a ByRef Variant without copying the array
 - ```StringToIntegers``` - copies the memory of a String to an Array of Integers
 - ```IntegersToString``` - copies the memory of an Array of Integers to a String 
 - ```EmptyArray``` - returns an empty array of the requested size and data type
 - ```UpdateLBound``` - changes the Lower Bound for a given array's dimension

## Class Instance Redirection

Class instance redirection is supported. This allows Private Class Initializers thus achieving true immutabilty.
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

## Testing
Just import [TestLibMemory.bas](https://github.com/cristianbuse/VBA-MemoryTools/blob/master/src/Test/TestLibMemory.bas) module and run method ```RunAllTests```. On failure, execution will stop on the first failed Assert.

Please [raise an issue](https://github.com/cristianbuse/VBA-MemoryTools/issues/new) if any test is failing.

## License
MIT License

Copyright (c) 2020 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
