VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGlobalHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event KeyPressed(VKCode As Long, _
                 Scancode As Long, _
                 Repeatcount As Long, _
                 ExtendedKey As Boolean, _
                 AltDown As Boolean, _
                 PrevState As Boolean, _
                 State As Boolean)

'---------------------------------------------------------------------------------------
' Procedure : ProcessKeyStroke
' DateTime  : 11-7-2002 10:40
' Purpose   : Ontleden van een keystroke en event fireen
'---------------------------------------------------------------------------------------
'
Friend Sub ProcessKeyStroke(VKCode As Long, Flags As Long)
    Dim Scancode As Long
    Dim Repeatcount As Long
    Dim ExtendedKey As Boolean
    Dim AltDown As Boolean
    Dim PrevState As Boolean
    Dim State As Boolean
    
    Scancode = Flags And &HFFFF          ' bits 0 .. 15
    Repeatcount = Flags And &HFF0000     ' bits 16 .. 23
    ExtendedKey = Flags And &H1000000    ' bit 24
    AltDown = Flags And &H20000000       ' bit 29
    PrevState = Flags And &H40000000     ' bit 30
    State = Flags And &H80000000         ' bit 31
    
    RaiseEvent KeyPressed(VKCode, Scancode, Repeatcount, ExtendedKey, AltDown, PrevState, State)
End Sub

Private Sub Class_Initialize()
    Set MCallback.m_GlobalHook = Me
    SetGlbHook
End Sub

Private Sub Class_Terminate()
    UnSetGlbHook
    Set MCallback.m_GlobalHook = Nothing
End Sub
