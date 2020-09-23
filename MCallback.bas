Attribute VB_Name = "MCallback"
Option Explicit
Private Declare Function SetHook Lib "kbhookdll.dll" (ByVal hwnd As Long) As Long
Private Declare Function RemoveHook Lib "kbhookdll.dll" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long
Private Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Const WM_USER = &H400
Private Const WM_COPYDATA = &H4A
Private Const GWL_WNDPROC = (-4)
Private PrevFuncPointer As Long
Private m_hookactive As Boolean
Public m_GlobalHook As CGlobalHook

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Function HookActive() As Boolean
    HookActive = m_hookactive
End Function

Public Function WndProc(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'Here wParam - Virtual KeyCode, lParam - Keyboard ScanCode
    Dim RetVal As Long
    Dim KeyAscii As Long
    Dim KeyName As String
    
    'Is this the message from the dll
    If MSG = WM_USER Then
        ' Yes, log it
        Debug.Print "KbdCallBack: VKCode=" & wParam & ", Flags=" & lParam
        m_GlobalHook.ProcessKeyStroke wParam, lParam
    End If
    
    'Pass the procedure to the default handler
    WndProc = CallWindowProc(PrevFuncPointer, hwnd, MSG, wParam, lParam)
End Function

Public Sub UnSetGlbHook()
  If Not HookActive Then Exit Sub
  
  Debug.Print "Removing global keyboard hook"
  
  'Removes the hook
  m_hookactive = Not RemoveHook
  'removes the text box's input from the WindowProc function
  If Not HookActive Then SetWindowLong Fmain.hwnd, GWL_WNDPROC, PrevFuncPointer
End Sub

Public Function SetGlbHook()
  If HookActive Then Exit Function
  
  Debug.Print "Setting global keyboard hook"
  
  'Sets the text box's input to the WindowProc function
  PrevFuncPointer = SetWindowLong(Fmain.hwnd, GWL_WNDPROC, AddressOf MCallback.WndProc)
  'Sets the hook from the dll and passes to the textbox which passes to WindowProc
  m_hookactive = SetHook(Fmain.hwnd)
  SetGlbHook = HookActive
End Function


