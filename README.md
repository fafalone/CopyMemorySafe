# CopyMemorySafe
Intro to Vectored Exception Handling: Crash-proof CopyMemory

It's been a long standing problem that access violations like a bad address for `CopyMemory` and other exceptions can't be handled by `On Error`. One solution to that is Vectored Exception Handling (VEH). You can register a procedure to handle true exceptions like access violations, then set it to skip the offending instruction. 

This is a small module to introduce the concept that allows you to call `CopyMemory` safely, your app will not crash even if you supply an invalid address. If an invalid address is provided, the operation is skipped, allowing you to then just check if your destination has changed to determine how to proceed.

This works by modifying the `CONTEXT` structure, which contains among other things the contents of all registers, including the instruction pointer register- `Eip` for 32bit, `Rip` for 64bit. If an access violation is encountered, we skip the instruction by adding the instruction size-- this is where it gets the most complicated, and to be honest I used Claude AI for the functions to calculate the length, and don't totally understand it, since it's dynamic at runtime and not just looking at the disassembly on disk.

`CONTEXT` and many other definitions in this project are entirely different for 32bit and 64bit, so the requirements are as follows:

VB6: Standalone. No dependencies. 32bit definitions are included.\
twinBASIC: Requires Windows Development Library for twinBASIC, added via References->Available packages. This is just code, there's no binary to be distributed.\

For both, just add modSafeCopy.bas to your project and use CopyMemorySafe in place of CopyMemory. Note you'll have to use VarPtr/StrPtr/ObjPtr since neither VB6 nor tB supports As Any in local functions.\

**NOTE:** In twinBASIC the CopyMemorySafe method is marked `[Debuggable(False)]` due to it not currently working in the IDE otherwise.
