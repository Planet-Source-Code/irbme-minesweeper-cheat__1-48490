Attribute VB_Name = "modMem"
Option Explicit

Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal ByteLen As Long)

Private Const PROCESS_QUERY_INFORMATION  As Long = 1024
Private Const PROCESS_VM_READ            As Long = 16
Private Const MAX_PATH                   As Long = 260


Public Function GetProcessByName(ByVal EXEName As String) As Long

  Dim cb                   As Long
  Dim cbNeeded             As Long
  Dim NumElements          As Long
  Dim ProcessIDs()         As Long
  Dim cbNeeded2            As Long
  Dim NumElements2         As Long
  Dim Modules(1 To 200)    As Long
  Dim ModuleName           As String
  Dim hProcess             As Long
  Dim i                    As Long
       
    cb = 8
    cbNeeded = 96
       
    Do While cb <= cbNeeded
        cb = cb * 2
        ReDim ProcessIDs(cb / 4) As Long
        EnumProcesses ProcessIDs(1), cb, cbNeeded
    Loop
       
    NumElements = cbNeeded / 4

    For i = 1 To NumElements
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))
           
        If hProcess <> 0 Then
            If EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2) <> 0 Then
                ModuleName = Space(MAX_PATH)

                If (InStr(1, Left$(ModuleName, GetModuleFileNameExA(hProcess, Modules(1), ModuleName, 500)), EXEName) > 0) Then
                    GetProcessByName = hProcess
                    Exit Function
                End If
                   
            End If
        End If

        CloseHandle hProcess
    Next
       
End Function


Public Function ReadMemory(hProcess As Long, lpAddress As Long, ReturnBuffer() As Byte, BytesToRead As Long)
  
  Dim lpBuffer As String, BytesRead As Long, rBytes As Long

    ReDim ReturnBuffer(BytesToRead)
    BytesRead = ReadProcessMemory(hProcess, lpAddress, ByVal VarPtr(ReturnBuffer(0)), BytesToRead, rBytes)
    ReDim Preserve ReturnBuffer(BytesRead)
    
End Function
