Attribute VB_Name = "modComboSelect"
Option Explicit
'Original created by James Haydock - Code@JamesHaydock.com

'ComboFindText API's & Constants
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_LIMITTEXT = &H141

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Function ComboFindText(hWnd As Long, Text As String, WholeString As Boolean) As Long

  Dim lngRtn As Long
  
  If WholeString Then
    lngRtn = SendMessage(hWnd, CB_FINDSTRINGEXACT, -1, ByVal Text & vbNullChar)
  Else
    lngRtn = SendMessage(hWnd, CB_FINDSTRING, -1, ByVal Text & vbNullChar)
  End If
  If lngRtn > -1 Then
    ComboFindText = lngRtn
  Else
    ComboFindText = -1
  End If

End Function

Public Property Get ComboDropDown(hWnd As Long) As Boolean
Attribute ComboDropDown.VB_MemberFlags = "400"

  ComboDropDown = SendMessage(hWnd, CB_GETDROPPEDSTATE, 0, 0)

End Property

Public Property Let ComboDropDown(hWnd As Long, ByVal ShowDropDown As Boolean)

  SendMessage hWnd, CB_SHOWDROPDOWN, ShowDropDown, 0

End Property

Public Property Get ComboDropDownWidth(hWnd As Long) As Integer

  ComboDropDownWidth = SendMessage(hWnd, CB_GETDROPPEDWIDTH, 0, 0)

End Property

Public Property Let ComboDropDownWidth(hWnd As Long, ByVal Width As Integer)

  If Width > 0 Then SendMessage hWnd, CB_SETDROPPEDWIDTH, Width, 0

End Property

Public Function MouseX(ByVal hWnd As Long) As Long

  Dim lpPoint As POINTAPI
  
  GetCursorPos lpPoint
  ScreenToClient hWnd, lpPoint
  MouseX = lpPoint.X

End Function

Public Function MouseY(ByVal hWnd As Long) As Long

  Dim lpPoint As POINTAPI
  
  GetCursorPos lpPoint
  ScreenToClient hWnd, lpPoint
  MouseY = lpPoint.Y

End Function
