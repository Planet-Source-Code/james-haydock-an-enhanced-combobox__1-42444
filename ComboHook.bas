Attribute VB_Name = "modHook"
Option Explicit
'This module is credited to Matt Norris

'Window Proc Hooking API's
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Window Proc Hooking Constant's
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_CUT = &H300
Public Const WM_PASTE = &H302

'Getting the hWnd of the Edit portion of a ComboBox
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5

'Getting the hWnd of the DropDown portion of a ComboBox
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Enum STATE_SYSTEM
  NOTPRESSED = 0
  INVISIBLE = &H8000
  PRESSED = &H8
End Enum

Private Type COMBOBOXINFO
  cbSize As Long
  rcItem As RECT
  rcButton As RECT
  stateButton As Long
  hwndCombo As Long
  hwndItem As Long
  hwndList As Long
End Type

Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hWnd As Long, pcbi As COMBOBOXINFO) As Boolean

'Local variables
'Private pcolHooked As Collection
Private mlngHWndOld1 As Long
Private mlngHWndNew1 As Long
Private mlngHWndOld2 As Long
Private mlngHWndNew2 As Long

'Private Function Hook(hWnd As Long) As Long
'
'  Dim lngOldProc As Long
'
'  If pcolHooked Is Nothing Then Set pcolHooked = New Collection
'  On Error Resume Next
'  lngOldProc = pcolHooked("@" & hWnd)
'  If Err.Number <> 0 Then
'    Err.Clear
'    lngOldProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
'    pcolHooked.Add lngOldProc, "@" & hWnd
'  End If
'  'Debug.Print "Hooked " & hWnd & " - " & pcolHooked.Count & " Hooked"
'  Hook = lngOldProc
'
'End Function

'Public Sub UnHook(hWnd As Long) ', OldProc As Long)
'
'  Dim lngOldProc As Long
'
'  On Error Resume Next
'  lngOldProc = pcolHooked("@" & hWnd)
'  If Err.Number <> 0 Then Err.Clear: Exit Sub
'  SetWindowLong hWnd, GWL_WNDPROC, lngOldProc
'  pcolHooked.Remove "@" & hWnd
'  'Debug.Print "Unhooked " & hWnd & " - " & pcolHooked.Count & " Hooked"
'
'End Sub

'Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
'  Dim lngOldProc As Long
'
'  On Error Resume Next
'  lngOldProc = pcolHooked("@" & hWnd)
'  If Err.Number <> 0 Then
'    Err.Clear
'    Exit Function
'  End If
'  'Debug.Print uMsg, wParam, lParam
'  WindowProc = 0
'  Select Case uMsg
'    Case WM_CONTEXTMENU, WM_PASTE, WM_CUT
'      '# EAT THESE MESSAGES YUM YUM ;) #
'      '# - Stops the DEFAULT Context menu popping up in VB Text Boxes #
'      '# - Stops CUT and PASTE operations in VB Text Boxes #
'    Case Else
'      Debug.Print uMsg
'      WindowProc = CallWindowProc(lngOldProc, hWnd, uMsg, wParam, lParam)
'  End Select
'
'End Function

Public Sub HookCombo(cbo As ComboBox)

  Dim pcbi As COMBOBOXINFO
  
  On Error Resume Next
  pcbi.cbSize = Len(pcbi)
  If mlngHWndOld1 = 0 And mlngHWndOld2 = 0 Then
    If GetComboBoxInfo(cbo.hWnd, pcbi) Then
      mlngHWndOld1 = GetWindow(cbo.hWnd, GW_CHILD)
      mlngHWndOld2 = pcbi.hwndList
      If mlngHWndOld1 > 0 And mlngHWndOld2 > 0 Then
        mlngHWndNew1 = SetWindowLong(mlngHWndOld1, GWL_WNDPROC, AddressOf WindowProc1)
        mlngHWndNew2 = SetWindowLong(mlngHWndOld2, GWL_WNDPROC, AddressOf WindowProc2)
      End If
    End If
  End If

End Sub

Public Sub UnHookAll()

  On Error Resume Next
  If mlngHWndOld1 > 0 And mlngHWndOld2 > 0 Then
    SetWindowLong mlngHWndOld1, GWL_WNDPROC, mlngHWndNew1
    SetWindowLong mlngHWndOld2, GWL_WNDPROC, mlngHWndNew2
    mlngHWndOld1 = 0
    mlngHWndOld2 = 0
  End If

End Sub

Public Function WindowProc1(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Dim lngOldProc As Long
  
  On Error Resume Next
  If Err.Number <> 0 Then
    Err.Clear
    Exit Function
  End If
  'Debug.Print uMsg, wParam, lParam
  WindowProc1 = 0
  Select Case uMsg
    Case WM_CONTEXTMENU, WM_PASTE, WM_CUT
      '# EAT THESE MESSAGES YUM YUM ;) #
      '# - Stops the DEFAULT Context menu popping up in VB Text Boxes #
      '# - Stops CUT and PASTE operations in VB Text Boxes #
    Case Else
      'Debug.Print "WindowProc1 ", uMsg, wParam, lParam
      WindowProc1 = CallWindowProc(mlngHWndNew1, hWnd, uMsg, wParam, lParam)
  End Select

End Function

Public Function WindowProc2(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Dim lngOldProc As Long
  
  On Error Resume Next
  If Err.Number <> 0 Then
    Err.Clear
    Exit Function
  End If
  WindowProc2 = 0
  Select Case uMsg
    Case WM_CONTEXTMENU, WM_PASTE, WM_CUT
      '# EAT THESE MESSAGES YUM YUM ;) #
      '# - Stops the DEFAULT Context menu popping up in VB Text Boxes #
      '# - Stops CUT and PASTE operations in VB Text Boxes #
'    Case 418
'      Debug.Print "WindowProc2 Test " & uMsg, wParam, lParam
'      WindowProc2 = CallWindowProc(mlngHWndNew2, hWnd, uMsg, wParam, lParam)
'      CallWindowProc mlngHWndNew1, mlngHWndOld1, 256, 13, 1835009 'uMsg, wParam, lParam '257, 70, -1071579135 '256, 13, 1835009 '176, 0, 0
    Case Else
      'Debug.Print "WindowProc2 ", uMsg, wParam, lParam
      WindowProc2 = CallWindowProc(mlngHWndNew2, hWnd, uMsg, wParam, lParam)
  End Select

End Function
