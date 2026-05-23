# CopyMemorySafe
Intro to Vectored Exception Handling: Crash-proof CopyMemory

It's been a long standing problem that access violations like a bad address for `CopyMemory` and other exceptions can't be handled by `On Error`. One solution to that is Vectored Exception Handling (VEH). You can register a procedure to handle true exceptions like access violations, then set it to skip the offending instruction. 

This is a small module to introduce the concept that allows you to call `CopyMemory` safely, your app will not crash even if you supply an invalid address. If an invalid address is provided, the operation is skipped and it returns False.

This works by modifying the `CONTEXT` structure, which contains among other things the contents of all registers (where things like arguments and return values are actually stored at the assembly code/hardware level), including the instruction pointer register that tells the system exactly what instruction is executing- `Eip` for 32bit, `Rip` for 64bit. If an access violation is encountered, we skip the instruction by adding the instruction size-- this is where it gets the most complicated, and to be honest I used Claude AI for the functions to calculate the length, and don't totally understand it, since it's dynamic at runtime and not just looking at the disassembly on disk.
 

**Requirements:**
This is a standalone module with no dependencies.\
The definitions were copied from Windows Development Library for twinBASIC, so if you use that package, you can remove all the declares/types/enums/consts.

**Usage**
Just add modSafeCopy.bas to your project and use CopyMemorySafe in place of CopyMemory. Note you'll have to use VarPtr/StrPtr/ObjPtr since neither VB6 nor tB supports As Any in local functions. You can also use `CopyMemorySafe` as a function, if the copy is successful without being passed a null pointer and without an exception it returns `True`. 

**NOTE:** In twinBASIC this currently only works in compiled exes.

**Important:** For VBA, the document must be saved in a Trusted Location.

