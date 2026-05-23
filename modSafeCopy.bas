Attribute VB_Name = "modSafeCopy"
Option Explicit

'******************************************************
'modSafeCopy.bas v1.1
'by Jon Johnson
'(c) 2026
'
'Provides an introduction to Vectored Exception Handling
'via a CopyMemory alternative that won't crash from an
'access violation attempting to read/write a bad address.
'If a bad address is passed, the operation is skipped.
'This is done by creating a handler for exceptions that
'normally crash the app/host. 
'
'Usage:
' Use CopyMemorySafe in place of CopyMemory (RtlMoveMemory)
' You'll need to use VarPtr/StrPtr since local functions
'  can't use As Any arguments.
'
'Updates:
' -v1.1 - Add local copy of 64bit defs to eliminate depends
'         so VBA can also use this.
'
'Definitions taken from Windows Development Library for twinBASIC
'You can remove all these if you have that package present.
'
' More details and up to date version:
'  https://github.com/fafalone/CopyMemorySafe
'
'******************************************************

#If VBA7 = 0 Then
Private Enum LongPtr: [_]: End Enum
#End If
Private Const STATUS_ACCESS_VIOLATION = &HC0000005
Private Const EXCEPTION_EXECUTE_HANDLER = 1
Private Const EXCEPTION_CONTINUE_SEARCH = 0
Private Const EXCEPTION_CONTINUE_EXECUTION = (-1)
Private Type EXCEPTION_POINTERS
    ExceptionRecord As LongPtr 'PEXCEPTION_RECORD
    ContextRecord As LongPtr 'PCONTEXT
End Type


Private Enum EXCEPTION_FLAGS
    EXCEPTION_NONCONTINUABLE = &H1  ' Noncontinuable exception
    EXCEPTION_UNWINDING = &H2  ' Unwind is in progress
    EXCEPTION_EXIT_UNWIND = &H4  ' Exit unwind is in progress
    EXCEPTION_STACK_INVALID = &H8  ' Stack out of limits or unaligned
    EXCEPTION_NESTED_CALL = &H10  ' Nested exception handler call
    EXCEPTION_TARGET_UNWIND = &H20  ' Target unwind in progress
    EXCEPTION_COLLIDED_UNWIND = &H40  ' Collided exception handler call
    EXCEPTION_SOFTWARE_ORIGINATE = &H80  ' Exception originated in software
    EXCEPTION_UNWIND = (EXCEPTION_UNWINDING Or EXCEPTION_EXIT_UNWIND Or EXCEPTION_TARGET_UNWIND Or EXCEPTION_COLLIDED_UNWIND)
End Enum

Private Const EXCEPTION_MAXIMUM_PARAMETERS = 15
Private Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As EXCEPTION_FLAGS
    ExceptionRecord As LongPtr
    ExceptionAddress As LongPtr
    NumberParameters As Long
    ExceptionInformation(0 To (EXCEPTION_MAXIMUM_PARAMETERS - 1)) As LongPtr
End Type


#If Win64 Then
Private Declare PtrSafe Function AddVectoredExceptionHandler Lib "kernel32" (ByVal First As Long, ByVal Handler As LongPtr) As LongPtr
Private Declare PtrSafe Function RemoveVectoredExceptionHandler Lib "kernel32" (ByVal Handle As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Public Enum EXCEPTION_CONTEXT_FLAGS
    CONTEXT_AMD64 = &H00100000
    CONTEXT_CONTROL = (CONTEXT_AMD64 Or &H00000001)
    CONTEXT_INTEGER = (CONTEXT_AMD64 Or &H00000002)
    CONTEXT_SEGMENTS = (CONTEXT_AMD64 Or &H00000004)
    CONTEXT_FLOATING_POINT = (CONTEXT_AMD64 Or &H00000008)
    CONTEXT_DEBUG_REGISTERS = (CONTEXT_AMD64 Or &H00000010)
    CONTEXT_FULL = (CONTEXT_CONTROL Or CONTEXT_INTEGER Or CONTEXT_FLOATING_POINT)
    CONTEXT_ALL = (CONTEXT_CONTROL Or CONTEXT_INTEGER Or CONTEXT_SEGMENTS Or CONTEXT_FLOATING_POINT Or CONTEXT_DEBUG_REGISTERS)
    CONTEXT_XSTATE = (CONTEXT_AMD64 Or &H00000040)
    CONTEXT_KERNEL_CET = (CONTEXT_AMD64 Or &H00000080)
    CONTEXT_EXCEPTION_ACTIVE = &H08000000
    CONTEXT_SERVICE_ACTIVE = &H10000000
    CONTEXT_EXCEPTION_REQUEST = &H40000000
    CONTEXT_EXCEPTION_REPORTING = &H80000000
'  CONTEXT_UNWOUND_TO_CALL flag is set by the unwinder if it
'  has unwound to a call site, and cleared whenever it unwinds
'  through a trap frame.
    CONTEXT_UNWOUND_TO_CALL = &H20000000
End Enum

Public Type M128A
    Low As LongLong
    High As LongLong
End Type
Public Type XSAVE_FORMAT
    ControlWord As Integer
    StatusWord As Integer
    TagWord As Byte
    Reserved1 As Byte
    ErrorOpcode As Integer
    ErrorOffset As Long
    ErrorSelector As Integer
    Reserved2 As Integer
    DataOffset As Long
    DataSelector As Integer
    Reserved3 As Integer
    MxCsr As Long
    MxCsr_Mask As Long
    FloatRegisters(0 To 7) As M128A
    #If Win64 Then
    XmmRegisters(0 To 15) As M128A
    Reserved4(0 To 95) As Byte
    #Else
    XmmRegisters(0 To 7) As M128A
    Reserved4(0 To 223) As Byte
    #End If
End Type

Public Type CONTEXT
    ' Register parameter home addresses.
    ' N.B. These fields are for convience - they could be used to extend the
    '      context record in the future.
    P1Home As LongLong
    P2Home As LongLong
    P3Home As LongLong
    P4Home As LongLong
    P5Home As LongLong
    P6Home As LongLong
    ' Control flags.
    ContextFlags As EXCEPTION_CONTEXT_FLAGS
    MxCsr As Long
    ' Segment Registers and processor flags.
    SegCs As Integer
    SegDs As Integer
    SegEs As Integer
    SegFs As Integer
    SegGs As Integer
    SegSs As Integer
    EFlags As Long
    ' Debug registers
    Dr0 As LongLong
    Dr1 As LongLong
    Dr2 As LongLong
    Dr3 As LongLong
    Dr6 As LongLong
    Dr7 As LongLong
    ' Integer registers.
    Rax As LongLong
    Rcx As LongLong
    Rdx As LongLong
    Rbx As LongLong
    Rsp As LongLong
    Rbp As LongLong
    Rsi As LongLong
    Rdi As LongLong
    R8 As LongLong
    R9 As LongLong
    R10 As LongLong
    R11 As LongLong
    R12 As LongLong
    R13 As LongLong
    R14 As LongLong
    R15 As LongLong
    ' Program counter.
    Rip As LongLong
    ' Floating point state.
    'union {
    FltSave As XSAVE_FORMAT 'XMM_SAVE_AREA32
    '    struct {
    '        M128A Header[2];
    '        M128A Legacy[8];
    '        M128A Xmm0;
    '        M128A Xmm1;
    '        M128A Xmm2;
    '        M128A Xmm3;
    '        M128A Xmm4;
    '        M128A Xmm5;
    '        M128A Xmm6;
    '        M128A Xmm7;
    '        M128A Xmm8;
    '        M128A Xmm9;
    '        M128A Xmm10;
    '        M128A Xmm11;
    '        M128A Xmm12;
    '        M128A Xmm13;
    '        M128A Xmm14;
    '        M128A Xmm15;
    '    } DUMMYSTRUCTNAME;
    '} DUMMYUNIONNAME;
    ' Vector registers.
    VectorRegister(0 To 25) As M128A
    VectorControl As LongLong
    ' Special debug control registers.
    DebugControl As LongLong
    LastBranchToRip As LongLong
    LastBranchFromRip As LongLong
    LastExceptionToRip As LongLong
    LastExceptionFromRip As LongLong
End Type

#Else
Private Declare Function AddVectoredExceptionHandler Lib "kernel32" (ByVal First As Long, ByVal Handler As LongPtr) As LongPtr
Private Declare Function RemoveVectoredExceptionHandler Lib "kernel32" (ByVal Handle As LongPtr) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)


Private Enum EXCEPTION_CONTEXT_FLAGS
    CONTEXT_i386 = &H10000     ' this assumes that i386 and
    CONTEXT_i486 = &H10000     ' i486 have identical context records
'  end_wx86
    CONTEXT_CONTROL = (CONTEXT_i386 Or &H1)         ' SS:SP, CS:IP, FLAGS, BP
    CONTEXT_INTEGER = (CONTEXT_i386 Or &H2)         ' AX, BX, CX, DX, SI, DI
    CONTEXT_SEGMENTS = (CONTEXT_i386 Or &H4)         ' DS, ES, FS, GS
    CONTEXT_FLOATING_POINT = (CONTEXT_i386 Or &H8)         ' 387 state
    CONTEXT_DEBUG_REGISTERS = (CONTEXT_i386 Or &H10)        ' DB 0-3,6,7
    CONTEXT_EXTENDED_REGISTERS = (CONTEXT_i386 Or &H20)        ' cpu specific extensions
    CONTEXT_FULL = (CONTEXT_CONTROL Or CONTEXT_INTEGER Or CONTEXT_SEGMENTS)
    CONTEXT_ALL = (CONTEXT_CONTROL Or CONTEXT_INTEGER Or CONTEXT_SEGMENTS Or CONTEXT_FLOATING_POINT Or CONTEXT_DEBUG_REGISTERS Or CONTEXT_EXTENDED_REGISTERS)
    CONTEXT_XSTATE = (CONTEXT_i386 Or &H40)
    CONTEXT_EXCEPTION_ACTIVE = &H8000000
    CONTEXT_SERVICE_ACTIVE = &H10000000
    CONTEXT_EXCEPTION_REQUEST = &H40000000
    CONTEXT_EXCEPTION_REPORTING = &H80000000
End Enum
Private Const SIZE_OF_80387_REGISTERS = 80
Private Type FLOATING_SAVE_AREA
    ControlWord As Long
    StatusWord As Long
    TagWord As Long
    ErrorOffset As Long
    ErrorSelector As Long
    DataOffset As Long
    DataSelector As Long
    RegisterArea(0 To (SIZE_OF_80387_REGISTERS - 1)) As Byte
    Spare0 As Long
End Type
Private Const MAXIMUM_SUPPORTED_EXTENSION = 512

Private Type CONTEXT
    ' The flags values within this flag control the contents of
    ' a CONTEXT record.
    ' If the context record is used as an input parameter, then
    ' for each portion of the context record controlled by a flag
    ' whose value is set, it is assumed that that portion of the
    ' context record contains valid context. If the context record
    ' is being used to modify a threads context, then only that
    ' portion of the threads context will be modified.
    ' If the context record is used as an IN OUT parameter to capture
    ' the context of a thread, then only those portions of the thread's
    ' context corresponding to set flags will be returned.
    ' The context record is never used as an OUT only parameter.
    ContextFlags As EXCEPTION_CONTEXT_FLAGS
    ' This section is specified/returned if CONTEXT_DEBUG_REGISTERS is
    ' set in ContextFlags.  Note that CONTEXT_DEBUG_REGISTERS is NOT
    ' included in CONTEXT_FULL.
    Dr0 As Long
    Dr1 As Long
    Dr2 As Long
    Dr3 As Long
    Dr6 As Long
    Dr7 As Long
    ' This section is specified/returned if the
    ' ContextFlags word contians the flag CONTEXT_FLOATING_POINT.
    FloatSave As FLOATING_SAVE_AREA
    ' This section is specified/returned if the
    ' ContextFlags word contians the flag CONTEXT_SEGMENTS.
    SegGs As Long
    SegFs As Long
    SegEs As Long
    SegDs As Long
    ' This section is specified/returned if the
    ' ContextFlags word contians the flag CONTEXT_INTEGER.
    Edi As Long
    Esi As Long
    Ebx As Long
    Edx As Long
    Ecx As Long
    Eax As Long
    ' This section is specified/returned if the
    ' ContextFlags word contians the flag CONTEXT_CONTROL.
    Ebp As Long
    Eip As Long
    SegCs As Long ' MUST BE SANITIZED
    EFlags As Long ' MUST BE SANITIZED
    Esp As Long
    SegSs As Long
    ' This section is specified/returned if the ContextFlags word
    ' contains the flag CONTEXT_EXTENDED_REGISTERS.
    ' The format and contexts are processor specific
    ExtendedRegisters(0 To (MAXIMUM_SUPPORTED_EXTENSION - 1)) As Byte
End Type
#End If


Private Function VectoredHandler(ExceptionInfo As EXCEPTION_POINTERS) As Long
     Dim pRecord As LongPtr 'EXCEPTION_RECORD PTR
     Dim pContext As LongPtr 'CONTEXT PTR
     Dim tRecord As EXCEPTION_RECORD
     Dim tContext As CONTEXT
     pRecord = ExceptionInfo.ExceptionRecord
     pContext = ExceptionInfo.ContextRecord
     CopyMemory tRecord, ByVal pRecord, LenB(tRecord)
     CopyMemory tContext, ByVal pContext, LenB(tContext)
     If tRecord.ExceptionCode = STATUS_ACCESS_VIOLATION Then
        Dim cbInstr As Long
        #If Win64 Then
        cbInstr = InstructionLength(tContext.Rip)
        If cbInstr = 0 Then
            VectoredHandler = EXCEPTION_CONTINUE_SEARCH
            Exit Function
        End If
        tContext.Rip = tContext.Rip + cbInstr
        #Else
        cbInstr = InstructionLength(tContext.Eip)
        If cbInstr = 0 Then
            VectoredHandler = EXCEPTION_CONTINUE_SEARCH
            Exit Function
        End If
        tContext.Eip = tContext.Eip + cbInstr
        #End If
        CopyMemory ByVal pContext, tContext, LenB(tContext)
        VectoredHandler = EXCEPTION_CONTINUE_EXECUTION
        Exit Function
     End If
     VectoredHandler = EXCEPTION_CONTINUE_SEARCH
 End Function
 
 #If Win64 Then
 Private Function InstructionLength(ByVal address As LongPtr) As Long
     Dim b As Byte
     Dim p As LongPtr
     Dim start As LongPtr
     Dim hasREX As Boolean
     Dim rexW As Boolean
     Dim op As Byte
     Dim modRM As Byte
     Dim mod_ As Byte
     Dim rm As Byte
     Dim hasSIB As Boolean
     Dim dispSize As Long
     Dim immSize As Long
     Dim opSize As Long ' operand size (default 32, REX.W = 64, 66h prefix = 16)
    
     start = address
     p = address
     opSize = 32
    
     ' 1. Consume prefixes
     Dim consuming As Boolean
     consuming = True
     Do While consuming
         CopyMemory b, ByVal p, 1
         Select Case b
             Case &HF0, &HF2, &HF3  ' LOCK, REPNE, REP
                 p = p + 1
             Case &H66               ' operand size override
                 opSize = 16
                 p = p + 1
             Case &H67               ' address size override
                 p = p + 1
             Case &H2E, &H3E, &H26, &H64, &H65  ' segment overrides
                 p = p + 1
             Case &H40 To &H4F       ' REX
                 hasREX = True
                 rexW = (b And &H8) <> 0
                 If rexW Then opSize = 64
                 p = p + 1
             Case &HC4, &HC5         ' VEX 3/2 byte - punt to 3rd party
                 InstructionLength = 0
                 Exit Function
             Case &H62               ' EVEX - punt to 3rd party
                 InstructionLength = 0
                 Exit Function
             Case Else
                 consuming = False
         End Select
     Loop
    
     ' 2. Read opcode
     CopyMemory op, ByVal p, 1
     p = p + 1
    
     Dim twoByteOp As Boolean
     If op = &HF Then
         twoByteOp = True
         CopyMemory op, ByVal p, 1
         p = p + 1
     End If
    
     ' 3. Determine ModRM, displacement, immediate sizes
     dispSize = 0
     immSize = 0
    
     If twoByteOp Then
         ' Common 0F opcodes
         Select Case op
             Case &H80 To &H8F   ' Jcc rel32
                 immSize = 4
             Case &HB6, &HB7, &HBE, &HBF  ' MOVZX, MOVSX - has ModRM
                 GoTo ReadModRM
             Case &H10, &H11, &H12, &H13  ' SSE MOV - has ModRM
                 GoTo ReadModRM
             Case Else
                 GoTo ReadModRM  ' assume ModRM for unknown 0F opcodes
         End Select
     Else
         Select Case op
             ' No ModRM, no immediate
             Case &H90                           ' NOP
             Case &H50 To &H5F                   ' PUSH/POP reg
             Case &H70 To &H7F                   ' Jcc rel8
                 immSize = 1
             Case &HEB                           ' JMP rel8
                 immSize = 1
             Case &HE9                           ' JMP rel32
                 immSize = 4
             Case &HE8                           ' CALL rel32
                 immSize = 4
             Case &HB0 To &HB7                   ' MOV r8, imm8
                 immSize = 1
             Case &HB8 To &HBF                   ' MOV r64/32, imm
                 If opSize = 64 Then
                     immSize = 8
                 Else
                     immSize = 4
                 End If
             Case &H68                           ' PUSH imm32
                 immSize = 4
             Case &H6A                           ' PUSH imm8
                 immSize = 1
             Case &HA0, &HA1, &HA2, &HA3        ' MOV moffs
                 immSize = 8  ' 64-bit address
             Case &HC3, &HCB                     ' RET
             Case &HCC, &HCD                     ' INT
                 If op = &HCD Then immSize = 1
             Case &H40 To &H48                   ' INC/DEC (32-bit mode only, REX in 64)
             Case &HF1                           ' INT1
             Case &H9C, &H9D                     ' PUSHF/POPF
             ' Has ModRM, no immediate
             Case &H0 To &H3, &H8 To &HB        ' ADD, OR
                 GoTo ReadModRM
             Case &H10 To &H13, &H18 To &H1B    ' ADC, SBB
                 GoTo ReadModRM
             Case &H20 To &H23, &H28 To &H2B    ' AND, SUB
                 GoTo ReadModRM
             Case &H30 To &H33, &H38 To &H3B    ' XOR, CMP
                 GoTo ReadModRM
             Case &H84 To &H8E                   ' TEST, XCHG, MOV, LEA etc
                 GoTo ReadModRM
             Case &HF6, &HF7                     ' TEST/NOT/NEG/MUL/DIV
                 GoTo ReadModRM
             Case &HFE, &HFF                     ' INC/DEC/CALL/JMP/PUSH mem
                 GoTo ReadModRM
             Case &HD0 To &HD3                   ' Shift/rotate
                 GoTo ReadModRM
             Case &HC0, &HC1                     ' Shift imm8
                 GoTo ReadModRM
                 ' (immSize set after ModRM)
             Case &H80, &H81, &H83               ' ADD/OR/ADC... imm
                 GoTo ReadModRM
                 ' (immSize set after ModRM)
             Case Else
                 ' Unknown - return 0 to signal failure
                 InstructionLength = 0
                 Exit Function
         End Select
         GoTo Done
     End If
    
ReadModRM:
     CopyMemory modRM, ByVal p, 1
     p = p + 1
     mod_ = (modRM And &HC0) \ &H40
     rm = modRM And &H7
    
     ' SIB byte
     hasSIB = (mod_ <> 3) And (rm = 4)
     If hasSIB Then p = p + 1
    
     ' Displacement
     If mod_ = 1 Then
         dispSize = 1
     ElseIf mod_ = 2 Then
         dispSize = 4
     ElseIf mod_ = 0 And rm = 5 Then
         dispSize = 4  ' RIP-relative
     End If
     p = p + dispSize
    
     ' Immediates that follow ModRM
     Select Case op
         Case &H80, &HC0, &HC1  ' imm8
             immSize = 1
         Case &H81               ' imm16/32
             If opSize = 16 Then immSize = 2 Else immSize = 4
         Case &H83               ' sign-extended imm8
             immSize = 1
     End Select

Done:
     p = p + immSize
     InstructionLength = CLng(p - start)
 End Function
 #Else
 Private Function InstructionLength(ByVal address As LongPtr) As Long
     Dim b As Byte
     Dim p As LongPtr
     Dim start As LongPtr
     Dim op As Byte
     Dim modRM As Byte
     Dim mod_ As Byte
     Dim rm As Byte
     Dim dispSize As Long
     Dim immSize As Long
     Dim opSize As Long      ' default 32
     Dim addrSize As Long    ' default 32
     Dim hasSIB As Boolean

     start = address
     p = address
     opSize = 32
     addrSize = 32

     ' 1. Consume prefixes
     Dim consuming As Boolean
     consuming = True
     Do While consuming
         CopyMemory b, ByVal p, 1
         Select Case b
             Case &HF0, &HF2, &HF3
                 p = p + 1
             Case &H66
                 opSize = 16
                 p = p + 1
             Case &H67
                 addrSize = 16
                 p = p + 1
             Case &H2E, &H3E, &H26, &H36, &H64, &H65
                 p = p + 1
             Case Else
                 consuming = False
         End Select
     Loop

     ' 2. Read opcode
     CopyMemory op, ByVal p, 1
     p = p + 1

     Dim twoByteOp As Boolean
     If op = &HF Then
         twoByteOp = True
         CopyMemory op, ByVal p, 1
         p = p + 1
     End If

     dispSize = 0
     immSize = 0

     If twoByteOp Then
         Select Case op
             Case &H80 To &H8F       ' Jcc rel32
                 immSize = 4
             Case &HB6, &HB7         ' MOVZX r, r/m8/16
                 GoTo ReadModRM
             Case &HBE, &HBF         ' MOVSX r, r/m8/16
                 GoTo ReadModRM
             Case &HA0 To &HA3       ' SHLD/SHRD imm8
                 GoTo ReadModRM
                 immSize = 1
             Case &HBA               ' BT/BTS/BTR/BTC r/m, imm8
                 GoTo ReadModRM
                 immSize = 1
             Case &H0 To &H3         ' SLDT/STR/LLDT etc, LAR/LSL
                 GoTo ReadModRM
             Case &H10 To &H17       ' SSE MOV
                 GoTo ReadModRM
             Case &H18               ' PREFETCH
                 GoTo ReadModRM
             Case &H1F               ' NOP r/m
                 GoTo ReadModRM
             Case &H28 To &H2F       ' SSE
                 GoTo ReadModRM
             Case &H40 To &H4F       ' CMOVcc
                 GoTo ReadModRM
             Case &H51 To &H5F       ' SSE
                 GoTo ReadModRM
             Case &H60 To &H76       ' MMX
                 GoTo ReadModRM
             Case &H90 To &H9F       ' SETcc r/m8
                 GoTo ReadModRM
             Case &HAF               ' IMUL r, r/m
                 GoTo ReadModRM
             Case &HB0, &HB1         ' CMPXCHG
                 GoTo ReadModRM
             Case &HB2, &HB4, &HB5   ' LSS/LFS/LGS
                 GoTo ReadModRM
             Case &HC0, &HC1         ' XADD
                 GoTo ReadModRM
             Case &HC2               ' CMPSS/SD imm8
                 GoTo ReadModRM
                 immSize = 1
             Case &HC4 To &HC6       ' PINSRW etc imm8
                 GoTo ReadModRM
                 immSize = 1
             Case &HD0 To &HDF       ' MMX/SSE
                 GoTo ReadModRM
             Case &HE0 To &HEF       ' MMX/SSE
                 GoTo ReadModRM
             Case &HF0 To &HFE       ' MMX/SSE
                 GoTo ReadModRM
             Case Else
                 InstructionLength = 0
                 Exit Function
         End Select
     Else
         Select Case op
             Case &H90               ' NOP
             Case &H50 To &H5F       ' PUSH/POP reg
             Case &H60, &H61         ' PUSHA/POPA
             Case &H9C, &H9D         ' PUSHF/POPF
             Case &H9E, &H9F         ' SAHF/LAHF
             Case &HA4 To &HA7       ' MOVS/CMPS
             Case &HAA To &HAF       ' STOS/LODS/SCAS
             Case &HC3, &HCB         ' RET
             Case &HF1               ' INT1
             Case &HF4               ' HLT
             Case &HF5               ' CMC
             Case &HF8 To &HFC       ' CLC/STC/CLI/STI/CLD/STD
             Case &H98, &H99         ' CBW/CWD
             Case &H9B               ' WAIT
             Case &HD7               ' XLAT

             Case &H70 To &H7F       ' Jcc rel8
                 immSize = 1
             Case &HEB               ' JMP rel8
                 immSize = 1
             Case &HE0 To &HE3       ' LOOPNE/LOOPE/LOOP/JCXZ rel8
                 immSize = 1
             Case &HE9               ' JMP rel32
                 immSize = 4
             Case &HE8               ' CALL rel32
                 immSize = 4
             Case &H9A               ' CALL far ptr
                 immSize = 6
             Case &HEA               ' JMP far ptr
                 immSize = 6

             Case &HCC               ' INT3
             Case &HCD               ' INT imm8
                 immSize = 1
             Case &HCE               ' INTO

             Case &HC2               ' RET imm16
                 immSize = 2
             Case &HCA               ' RETF imm16
                 immSize = 2

             Case &HD4               ' AAM imm8
                 immSize = 1
             Case &HD5               ' AAD imm8
                 immSize = 1

             Case &HE4, &HE5         ' IN eAX, imm8
                 immSize = 1
             Case &HE6, &HE7         ' OUT imm8, eAX
                 immSize = 1

             Case &HA0, &HA1         ' MOV AL/eAX, moffs
                 immSize = 4
             Case &HA2, &HA3         ' MOV moffs, AL/eAX
                 immSize = 4

             Case &HB0 To &HB7       ' MOV r8, imm8
                 immSize = 1
             Case &HB8 To &HBF       ' MOV r32, imm32 (or r16 with 66h)
                 If opSize = 16 Then immSize = 2 Else immSize = 4

             Case &H68               ' PUSH imm32
                 If opSize = 16 Then immSize = 2 Else immSize = 4
             Case &H6A               ' PUSH imm8
                 immSize = 1

             Case &H6B               ' IMUL r, r/m, imm8
                 GoTo ReadModRM
                 immSize = 1
             Case &H69               ' IMUL r, r/m, imm32
                 GoTo ReadModRM
                 If opSize = 16 Then immSize = 2 Else immSize = 4

             ' ModRM, no immediate
             Case &H0 To &H3, &H8 To &HB
                 GoTo ReadModRM
             Case &H10 To &H13, &H18 To &H1B
                 GoTo ReadModRM
             Case &H20 To &H23, &H28 To &H2B
                 GoTo ReadModRM
             Case &H30 To &H33, &H38 To &H3B
                 GoTo ReadModRM
             Case &H84 To &H8E
                 GoTo ReadModRM
             Case &HF6, &HF7
                 GoTo ReadModRM
             Case &HFE, &HFF
                 GoTo ReadModRM
             Case &HD0 To &HD3
                 GoTo ReadModRM
             Case &H62, &H63         ' BOUND, ARPL
                 GoTo ReadModRM
             Case &H8F               ' POP r/m
                 GoTo ReadModRM

             ' ModRM + imm8
             Case &HC0, &HC1         ' Shift r/m, imm8
                 GoTo ReadModRM
                 immSize = 1
             Case &H80               ' ADD/OR/ADC... r/m, imm8
                 GoTo ReadModRM
                 immSize = 1
             Case &H83               ' ADD/OR/ADC... r/m, sign-ext imm8
                 GoTo ReadModRM
                 immSize = 1

             ' ModRM + imm16/32
             Case &H81               ' ADD/OR/ADC... r/m, imm
                 GoTo ReadModRM
                 If opSize = 16 Then immSize = 2 Else immSize = 4

             Case &HC4, &HC5         ' LES/LDS
                 GoTo ReadModRM

             Case &HD8 To &HDF       ' FPU ESC
                 GoTo ReadModRM

             Case Else
                 InstructionLength = 0
                 Exit Function
         End Select
         GoTo Done
     End If

ReadModRM:
     CopyMemory modRM, ByVal p, 1
     p = p + 1
     mod_ = (modRM And &HC0) \ &H40
     rm = modRM And &H7

     If addrSize = 16 Then
         ' 16-bit addressing modes - no SIB
         If mod_ = 1 Then
             dispSize = 1
         ElseIf mod_ = 2 Then
             dispSize = 2
         ElseIf mod_ = 0 And rm = 6 Then
             dispSize = 2  ' direct address
         End If
     Else
         ' 32-bit addressing modes
         hasSIB = (mod_ <> 3) And (rm = 4)
         If hasSIB Then p = p + 1

         If mod_ = 1 Then
             dispSize = 1
         ElseIf mod_ = 2 Then
             dispSize = 4
         ElseIf mod_ = 0 And rm = 5 Then
             dispSize = 4  ' disp32 direct
         End If
     End If
     p = p + dispSize

Done:
     p = p + immSize
     InstructionLength = CLng(p - start)
 End Function
 #End If
 
 Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
     MakeTrue = True: bValue = True
 End Function
 
 #If TWINBASIC Then
[Debuggable(False)]
 #End If
 #If VBA7 Then
 Public Sub CopyMemorySafe(ByVal pDest As LongPtr, ByVal pSrc As LongPtr, ByVal Length As LongPtr)
 Attribute CopyMemorySafe.VB_Description = "A crash-proof CopyMemory wrapper. If an invalid address is passed, the operation is skipped. "
  #Else
 Public Sub CopyMemorySafe(ByVal pDest As Long, ByVal pSrc As Long, ByVal Length As Long)
 Attribute CopyMemorySafe.VB_Description = "A crash-proof CopyMemory wrapper. If an invalid address is passed, the operation is skipped. "
 #End If
 If pDest = 0 Or pSrc = 0 Then Exit Sub
 Dim hVeh As LongPtr
 hVeh = AddVectoredExceptionHandler(1, AddressOf VectoredHandler)
 CopyMemory ByVal pDest, ByVal pSrc, Length
 If hVeh <> 0 Then
     RemoveVectoredExceptionHandler hVeh
 End If
 End Sub


