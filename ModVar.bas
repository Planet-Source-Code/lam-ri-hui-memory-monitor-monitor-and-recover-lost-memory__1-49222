Attribute VB_Name = "Variables"

Public Type MemoryStatus
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

    Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MemoryStatus)
    Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter_ As Long, ByVal X As Long, ByVal y_ As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags_ As Long) As Long
    Const conHwndTopmost = -1
    Const conHwndNoTopmost = -2
    Const conSwpNoActivate = &H10
    Const conSwpShowWindow = &H40
Public Function Always_On_Top(ByVal H, FrmX As Long, FrmY As Long, Hght As Long, Wdth As Long, YesAOT As Boolean)

    If YesAOT = True Then
        SetWindowPos H, conHwndTopmost, FrmX, FrmY, Wdth, Hght, conSwpNoActivate
        ElseIf YesAOT = False Then
        SetWindowPos H, conHwndNoTopmost, FrmX, FrmY, Wdth, Hght, conSwpShowWindow
    End If

End Function
