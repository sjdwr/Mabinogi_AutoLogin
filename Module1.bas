Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) As Long
        
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szEXEFile As String * 260
End Type

Dim Pid As Long
Dim WindowHandle As Long

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim Temp As Long
    GetWindowThreadProcessId hwnd, Temp
    
    If Temp = Pid Then
        WindowHandle = hwnd
        EnumWindowsProc = 0
    Else
        EnumWindowsProc = 1
    End If
End Function

Private Sub GetHandleFromPid(ProcessName As String)
    Dim uProcess As PROCESSENTRY32, hSnapShot As Long, Result As Long
    
    uProcess.dwSize = Len(uProcess)
    hSnapShot = CreateToolhelpSnapshot(2&, 0&)
    Result = ProcessFirst(hSnapShot, uProcess)
    
    Do
        If Split(uProcess.szEXEFile, vbNullChar)(0) = ProcessName Then
            Pid = uProcess.th32ProcessID
            Call EnumWindows(AddressOf EnumWindowsProc, 0)
        End If
        
        Result = ProcessNext(hSnapShot, uProcess)
    Loop While Result
End Sub

Public Function GetHandleFromPrcName(ProcessName As String) As Long
    Pid = 0
    WindowHandle = 0
    
    Call GetHandleFromPid(ProcessName)
    GetHandleFromPrcName = WindowHandle
End Function


