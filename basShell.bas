Attribute VB_Name = "basShell"
Option Explicit
'System Info
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Public Type OSVERSIONINFO
   dwOSVersionInfoSize  As Long 'Specifies the length, in bytes, of the structure.
   dwMajorVersion       As Long 'Major Version Number
   dwMinorVersion       As Long 'Minor Version Number
   dwBuildNumber        As Long 'Build Version Number
   dwPlatformId         As Long 'Operating System Running, see below
   szCSDVersion As String * 128 'Windows NT: Contains a null-terminated string,
                                'such as "Service Pack 3", that indicates the latest
                                'Service Pack installed on the system.
                                'If no Service Pack has been installed, the string is empty.
                                'Windows 95: Contains a null-terminated string that provides
                                'arbitrary additional information about the operating system
End Type

Public Const hNull = 0

'  dwPlatformId defines:
Public Const VER_PLATFORM_WIN32s = 0            'Win32s on Windows 3.1.
Public Const VER_PLATFORM_WIN32_WINDOWS = 1     'Win32 on Windows 95 or Windows 98.
                                                'For Windows 95, dwMinorVersion is 0.
                                                'For Windows 98, dwMinorVersion is 1.
Public Const VER_PLATFORM_WIN32_NT = 2          'Win32 on Windows NT.


'==========================================
'WINDOWS 95/98 ONLY, ToolHelp32 APIs don't
'exist on Windows NT, use PSAPI.DLL instead
'==========================================

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, lpme As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, lpme As MODULEENTRY32) As Long

Private Const MAX_MODULE_NAME32 As Integer = 255
Private Const MAX_MODULE_NAME32plus As Integer = MAX_MODULE_NAME32 + 1
Private Const MAX_PATH = 260

Private Const TH32CS_SNAPPROCESS = &H2&
Private Const TH32CS_SNAPMODULE = &H8&

Public Type PROCESSENTRY32
   dwSize               As Long 'Specifies the length, in bytes, of the structure.
   cntUsage             As Long 'Number of references to the process.
   th32ProcessID        As Long 'Identifier of the process.
   th32DefaultHeapID    As Long 'Identifier of the default heap for the process.
   th32ModuleID         As Long 'Module identifier of the process. (Associated exe)
   cntThreads           As Long 'Number of execution threads started by the process.
   th32ParentProcessID  As Long 'Identifier of the process that created the process being examined.
   pcPriClassBase       As Long 'Base priority of any threads created by this process.
   dwFlags              As Long 'Reserved; do not use.
   szExeFile            As String * MAX_PATH 'Path and filename of the executable file for the process.
End Type

Public Type MODULEENTRY32
    dwSize          As Long 'Specifies the length, in bytes, of the structure.
    th32ModuleID    As Long 'Module identifier in the context of the owning process.
    th32ProcessID   As Long 'Identifier of the process being examined.
    GlblcntUsage    As Long 'Global usage count on the module.
    ProccntUsage    As Long 'Module usage count in the context of the owning process.
    modBaseAddr     As Long 'Base address of the module in the context of the owning process.
    modBaseSize     As Long 'Size, in bytes, of the module.
    hModule         As Long 'Handle to the module in the context of the owning process.
    szModule        As String * MAX_MODULE_NAME32plus 'String containing the module name.
    szExePath       As String * MAX_PATH 'String containing the location (path) of the module.
End Type

'NT Stuff
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDeskTopWindow Lib "User32" Alias "GetDesktopWindow" () As Long
Private Declare Function ShowWindow Lib "User32" (ByVal hwnd As Long, ByVal nCmdShow As Integer) As Long
Private Declare Function GetWindow Lib "User32" (ByVal hwnd As Long, ByVal wCmd As Integer) As Long
Private Declare Function GetTopWindow Lib "User32" (ByVal hwnd As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindow Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function IsChild Lib "User32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cdReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1& ' API
Private Const GW_HWNDNEXT = 2

Public Function GetShortName(ByVal sLongFileName As String) As String
Dim lRetVal As Long, sShortPathName As String, iLen As Integer
       'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)

    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    'Strip away unwanted characters.
    GetShortName = left(sShortPathName, lRetVal)
End Function


Public Function WindowsShell(fileName As String, Optional WindowStyle As VbAppWinStyle = vbNormalFocus) As Long
Dim Scr_hDC As Long
    Scr_hDC = GetDeskTopWindow()
    WindowsShell = ShellExecute(Scr_hDC, "Open", fileName, "", "C:\", WindowStyle)
End Function
Public Sub HideForm(MainForm As Form)
Dim OwnerhWnd As Long
Dim ret As Long
Const SW_HIDE = 0
Const GW_OWNER = 4

    OwnerhWnd = GetWindow(MainForm.hwnd, GW_OWNER)
    ret = ShowWindow(OwnerhWnd, SW_HIDE)
End Sub

Public Sub GetWindowList(Optional ByRef windowName As Collection, Optional ByRef windowHandle As Collection, Optional ByRef parentHandle As Collection)
Dim TotalProcs As Integer
Dim procName As String
Dim winTitle As String * 256
Dim appHandle As Long


    TotalProcs = 0
    appHandle = GetTopWindow(0)
    Do Until appHandle = 0
        winTitle = Space(256)
        Call GetWindowText(appHandle, winTitle, 255)
        procName = Mid(winTitle, 1, InStr(1, winTitle, Chr$(0), vbTextCompare) - 1)
        appHandle = GetWindow(appHandle, GW_HWNDNEXT)
        If Len(procName) > 0 Then
            TotalProcs = TotalProcs + 1
        End If
    Loop
    
    Set windowName = New Collection
    Set windowHandle = New Collection
    Set parentHandle = New Collection
    
    TotalProcs = 0
    appHandle = GetTopWindow(0)
    Do Until appHandle = 0
        winTitle = Space(256)
        Call GetWindowText(appHandle, winTitle, 255)
        procName = Mid(winTitle, 1, InStr(1, winTitle, Chr$(0), vbTextCompare) - 1)
        If Len(procName) > 0 Then
            TotalProcs = TotalProcs + 1
            Call windowName.Add(CStr(procName))
            Call windowHandle.Add(CLng(appHandle))
            Call parentHandle.Add(CStr(GetParent(appHandle)))
        End If
        appHandle = GetWindow(appHandle, GW_HWNDNEXT)
    Loop
            
End Sub

Public Function KillWindow(Optional windowName As String, Optional windowHandle As Long) As Boolean

Const WM_CLOSE = &H10
Const INFINITE = &HFFFFFFFF
Dim hWindow As Long
Dim lngReturnValue As Long
Dim ret As Long

    If windowHandle = 0 Then
        hWindow = FindWindow(vbNullString, windowName)
        If hWindow = 0 Then
            KillWindow = False
            Exit Function
        End If
    Else
        hWindow = windowHandle
    End If
    lngReturnValue = PostMessage(hWindow, WM_CLOSE, vbNull, vbNull)
    If lngReturnValue = 0 Then
        KillWindow = False
        Exit Function
    Else
        ret = WaitForSingleObject(hWindow, INFINITE)
        DoEvents
        KillWindow = True
    End If
    Call Sleep(1000)
End Function

Public Sub ShellAndWait(cmdline As String)
Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
Dim ret As Long
    start.cb = Len(start)
    ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    ret& = WaitForSingleObject(proc.hProcess, INFINITE)
    ret& = CloseHandle(proc.hProcess)
End Sub

Public Sub GetProcessList(Optional ByRef processName As Collection, Optional ByRef processNumber As Collection, Optional ByVal getDlls As Boolean = True)
    If GetVersion = 1 Then
        Set processName = FillProcessList95(processNumber, getDlls)
    Else
        Set processName = FillProcessListNT(processNumber, getDlls)
    End If
End Sub


Private Function StrZToStr(s As String) As String
   StrZToStr = left$(s, Len(s) - 1)
End Function

Private Function FillProcessList95(Optional ByRef processNumber As Collection, Optional ByVal getDlls As Boolean = True) As Collection
    '=========================================================
    'Clears the listbox specified by the DestListBox parameter
    'and then fills the list with the processes and the
    'modules used by each process
    '=========================================================
    Dim lReturnID       As Long
    Dim hSnapProcess    As Long
    Dim hSnapModule     As Long
    Dim sName           As String
    Dim proc            As PROCESSENTRY32
    Dim module          As MODULEENTRY32
    Dim iProcesses      As Integer
    Dim iModules        As Integer
    
    'Clear the collection
    Set FillProcessList95 = New Collection
    
    'Get a 'at this moment' snapshot of all the processes
    hSnapProcess = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    
    If hSnapProcess = hNull Then
        'If the snapshot is empty, then exit
        'Return the number of Processes found
    Else
        'Initialize the processentry structure
        proc.dwSize = Len(proc)
        
        'Get first process
        lReturnID = Process32First(hSnapProcess, proc)
                
        'Iterate through each process with an ID that <> 0
        Do While lReturnID
            
            'Add the process to the listbox
            FillProcessList95.Add StrZToStr(proc.szExeFile)
            
            processNumber.Add lReturnID
            'Increment the count of processes we've added
            iProcesses = iProcesses + 1
            
            If getDlls = True Then
                'Get a 'at this moment' snapshot of all the modules in this process
                hSnapModule = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, proc.th32ProcessID)
            
                'If the process has modules loaded, iterate through them
                If Not hSnapModule = hNull Then
                
                    'Initialize the moduleentry structure
                    module.dwSize = LenB(module) - 1
                    
                    'Get first module
                    lReturnID = Module32First(hSnapModule, module)
                    
                    ' Iterate through the modules with an ID that <> 0
                    Do While lReturnID
                        
                        'If there is a module, add it to the list
                        FillProcessList95.Add StrZToStr(module.szModule)
                        processNumber.Add lReturnID
                        
                        'Get next module
                        lReturnID = Module32Next(hSnapModule, module)
                    Loop
                End If
            End If
            
            'Close the module snapshot handle
            CloseHandle hSnapModule
            
            'Get next process
            lReturnID = Process32Next(hSnapProcess, proc)
        Loop
        
        'Close the Process snapshot handle
        CloseHandle hSnapProcess
        
    End If
End Function

Private Function FillProcessListNT(Optional ByRef processNumber As Collection, Optional ByVal getDlls As Boolean = True) As Collection
    '=========================================================
    'Clears the listbox specified by the DestListBox parameter
    'and then fills the list with the processes and the
    'modules used by each process
    '=========================================================

    Dim cb                  As Long
    Dim cbNeeded            As Long
    Dim NumElements         As Long
    Dim ProcessIDs()        As Long
    Dim cbNeeded2           As Long
    Dim NumElements2        As Long
    Dim Modules(1 To 200)   As Long
    Dim lRet                As Long
    Dim ModuleName          As String
    Dim nSize               As Long
    Dim hProcess            As Long
    Dim i                   As Long
    Dim sModName            As String
    Dim sChildModName       As String
    Dim iModDlls            As Long
    Dim iProcesses          As Integer
    
    Set FillProcessListNT = New Collection
    Set processNumber = New Collection
    
    'Get the array containing the process id's for each process object
    cb = 8
    cbNeeded = 96
    
    'One important note should be made. Although the documentation
    'names the returned DWORD "cbNeeded", there is actually no way
    'to find out how big the passed in array must be. EnumProcesses()
    'will never return a value in cbNeeded that is larger than the
    'size of array value that you passed in the cb parameter.
    
    'if cbNeeded == cb upon return, allocate a larger array
    'and try again until cbNeeded is smaller than cb.
    Do While cb <= cbNeeded
       cb = cb * 2
       ReDim ProcessIDs(cb / 4) As Long
       lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
    Loop
    
    'calculate how many process IDs were returned
    NumElements = cbNeeded / 4
    
    For i = 1 To NumElements
    
        'Get a handle to the Process
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))
        
        ' Iterate through each process with an ID that <> 0
        If hProcess Then
            
            'Get an array of the module handles for the specified process
            lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
            
            'If the Module Array is retrieved, Get the ModuleFileName
            If lRet <> 0 Then
                
                'Fill the ModuleName buffer with spaces
                ModuleName = Space(MAX_PATH)
                
                'Preset buffer size
                nSize = 500
                
                'Get the module file name
                lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
                
                'Get the module file name out of the buffer, lRet is how
                'many characters the string is, the rest of the buffer is spaces
                sModName = left$(ModuleName, lRet)
                
                'Add the process to the listbox
                processNumber.Add ProcessIDs(i)
                FillProcessListNT.Add sModName
                
                'Increment the count of processes we've added
                iProcesses = iProcesses + 1
                
                If getDlls = True Then
                    iModDlls = 1
                    Do
                        iModDlls = iModDlls + 1
    
                        'Fill the ModuleName buffer with spaces
                        ModuleName = Space(MAX_PATH)
    
                        'Preset buffer size
                        nSize = 500
    
                        'Get the module file name out of the buffer, lRet is how
                        'many characters the string is, the rest of the buffer is spaces
                        lRet = GetModuleFileNameExA(hProcess, Modules(iModDlls), ModuleName, nSize)
                        sChildModName = left$(ModuleName, lRet)
    
                        If sChildModName = sModName Then Exit Do
                        If Trim(sChildModName) <> "" Then
                            processNumber.Add Modules(iModDlls)
                            FillProcessListNT.Add sChildModName
                        End If
                    Loop
                End If
            End If
        End If
        
        'Close the handle to the process
        lRet = CloseHandle(hProcess)
    Next
    
End Function

Private Function GetVersion() As Long
    '=======================================
    'Returns the Operating System being used
    '1 = Windows 95 / Windows 98
    '2 = Windows NT
    '=======================================
    Dim osinfo   As OSVERSIONINFO
    Dim retvalue As Integer
    
    With osinfo
        .dwOSVersionInfoSize = 148
        .szCSDVersion = Space$(128)
        retvalue = GetVersionExA(osinfo)
        GetVersion = .dwPlatformId
    End With
End Function

Public Function KillProcess(Optional processName As String, Optional processHandle As Long) As Boolean

Const WM_CLOSE = &H10
Const INFINITE = &HFFFFFFFF
Dim hWindow As Long
Dim lngReturnValue As Long
Dim ProcessNameColl As Collection
Dim ProcessNumberColl As Collection
Dim nCount As Integer
Dim ret As Long
Dim hProcess As Long

    If processHandle = 0 Then
        processName = Trim(processName)
        Call GetProcessList(ProcessNameColl, ProcessNumberColl, False)
        If ProcessNameColl.Count > 0 Then
            For nCount = 1 To ProcessNameColl.Count
                If InStr(1, UCase(processName), UCase(ProcessNameColl(nCount))) <> 0 Then
                    hProcess = OpenProcess(PROCESS_ALL_ACCESS, True, CLng(ProcessNumberColl(nCount)))
                    ret = TerminateProcess(hProcess, 0&)
                    If ret = 0 Then
                        KillProcess = False
                    End If
                End If
            Next nCount
        End If
        KillProcess = True
        Exit Function
    Else
        hProcess = OpenProcess(PROCESS_ALL_ACCESS, True, processHandle)
        ret = TerminateProcess(hProcess, 0&)
        If ret = 0 Then
            KillProcess = False
        End If
    End If
    
End Function


