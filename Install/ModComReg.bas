Attribute VB_Name = "ModComReg"
Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "Kernel32" (ByVal hLibModule As Long) As Long
Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Declare Function CreateThread Lib "Kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function GetExitCodeThread Lib "Kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long

Declare Sub ExitThread Lib "Kernel32" (ByVal dwExitCode As Long)

Enum RegOp
    Register = 1
    UnRegister
End Enum

Public Function FixPath(lzPath As String) As String
    If Right(FixPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Public Function RegisterActiveX(lzAxDll As String, mRegOption As RegOp) As Boolean
Dim mLib As Long, DllProcAddress As Long
Dim mThread
Dim sWait As Long
Dim mExitCode As Long
Dim lpThreadID As Long

Dim slib As String

    slib = lzAxDll
    mLib = LoadLibrary(slib)
    
    If mLib <= 0 Then
        RegisterActiveX = False
        Exit Function
    End If
    
    If mRegOption = Register Then
        DllProcAddress = GetProcAddress(mLib, "DllRegisterServer")
    Else
        DllProcAddress = GetProcAddress(mLib, "DllUnregisterServer")
    End If
    
    If DllProcAddress = 0 Then
        RegisterActiveX = True
        Exit Function
    Else
        mThread = CreateThread(ByVal 0, 0, ByVal DllProcAddress, ByVal 0, 0, lpThreadID)
        
        If mThread = 0 Then
            FreeLibrary mLib
            RegisterActiveX = False
            Exit Function
        Else
            sWait = WaitForSingleObject(mThread, 10000)
            If sWait <> 0 Then
                FreeLibrary lLib
                mExitCode = GetExitCodeThread(mThread, mExitCode)
                ExitThread mExitCode
                Exit Function
            Else
                FreeLibrary mLib
                CloseHandle mThread
            End If
        End If
    End If
    slib = ""
    RegisterActiveX = True
    
End Function

