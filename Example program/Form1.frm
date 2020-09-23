VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1140
   ClientLeft      =   2490
   ClientTop       =   4590
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   5175
   Begin VB.Timer Timer 
      Interval        =   50
      Left            =   960
      Top             =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bShow As Boolean
Private Const VK_F2 = &H71
Private Const VK_SHIFT = &H10
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private WithEvents hook As CGlobalHook
Attribute hook.VB_VarHelpID = -1

Private Sub Form_Load()
    Set hook = New CGlobalHook
    Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set hook = Nothing
End Sub

Private Sub hook_KeyPressed(VKCode As Long, Scancode As Long, Repeatcount As Long, ExtendedKey As Boolean, AltDown As Boolean, PrevState As Boolean, State As Boolean)
    Debug.Print "VKCode = " & VKCode & ", shift is " & GetAsyncKeyState(VK_SHIFT)
    If VKCode = VK_F2 And (GetAsyncKeyState(VK_SHIFT) <> 0) Then
        If AltDown Then
            Unload Me
        Else
            m_bShow = True
        End If
    End If
End Sub

Private Sub Timer_Timer()
    If m_bShow Then
        Call MsgBox("You pressed Shift-F2!", vbInformation + vbSystemModal, "Information")
        m_bShow = False
    End If
End Sub
