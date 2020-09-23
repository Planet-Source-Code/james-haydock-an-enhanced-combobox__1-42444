VERSION 5.00
Begin VB.UserControl ComboSelect 
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   ScaleHeight     =   840
   ScaleWidth      =   1935
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Left            =   60
      Top             =   360
   End
   Begin VB.ComboBox cbo 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "ComboSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Original created by James Haydock - Code@JamesHaydock.com

'Please note the usage of Property Let DropDownWidth (See property)

'Enum specifying what is permitted to be typed into the combobox
Public Enum cbsLockLevel
  cbsNone
  cbsOffice
  cbsFull
End Enum

'Default Property Values:
Private Const mCenmLockLevel As Integer = cbsFull
Private Const mClngChangedInterval As Long = 0
Private Const mCblnDateComplete As Boolean = False

'Property Variables:
Private menmLockLevel As cbsLockLevel
Private mblnDateComplete As Boolean

'Private Variables
Private mintKeyValue As Integer
Private mblnChangedEnabled As Boolean
Private mblnChangedBefore As Boolean
Private mblnIgnore As Boolean

'Public Events
Public Event Change()
Public Event Changed()
Public Event Click()
Public Event Clicked()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Public Sub AddItem(Text As String, Optional Value As Long = -1)

  cbo.AddItem Text
  If Value <> -1 Then cbo.ItemData(cbo.NewIndex) = Value

End Sub

Public Property Get BackColor() As OLE_COLOR

  BackColor = cbo.BackColor

End Property

Public Property Let BackColor(ByVal Color As OLE_COLOR)

  cbo.BackColor() = Color
  PropertyChanged "BackColor"

End Property

Private Sub cbo_Change()

  RaiseEvent Change

End Sub

Private Sub cbo_Click()
If mblnIgnore Then Exit Sub

  If mblnChangedEnabled Then
    tmr.Enabled = False
    tmr.Enabled = True
    mblnChangedBefore = True
  End If
  RaiseEvent Click

End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)

  RaiseEvent KeyDown(KeyCode, Shift)
  mintKeyValue = KeyCode
  Select Case menmLockLevel
    Case cbsFull
      KeyCode = KeyDownLockFull(KeyCode)
  End Select

End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)

  RaiseEvent KeyPress(KeyAscii)
  mintKeyValue = KeyAscii
  Select Case menmLockLevel
    Case cbsOffice
      KeyAscii = KeyPressLockOffice(KeyAscii)
    Case cbsFull
      KeyAscii = KeyPressLockFull(KeyAscii)
  End Select

End Sub

Private Sub cbo_KeyUp(KeyCode As Integer, Shift As Integer)

  RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Public Property Get ChangedInterval() As Long

  ChangedInterval = tmr.Interval

End Property

Public Property Let ChangedInterval(v As Long)

  Select Case v
    Case Is < 0
      ChangedInterval = 0
    Case 0
      mblnChangedEnabled = False
      tmr.Interval = v
    Case Else
      mblnChangedEnabled = True
      tmr.Interval = v
  End Select

End Property

Public Sub Clear()

  cbo.Clear

End Sub

Public Property Get DateComplete() As Boolean

  DateComplete = mblnDateComplete

End Property

Public Property Let DateComplete(v As Boolean)

  mblnDateComplete = v

End Property

Public Property Get DropDown() As Boolean

  DropDown = ComboDropDown(cbo.hWnd)

End Property

Public Property Let DropDown(ShowDropDown As Boolean)

  Dim mpcMousePointer As MousePointerConstants
  
  mpcMousePointer = MousePointer
  ComboDropDown(cbo.hWnd) = ShowDropDown
  MousePointer = mpcMousePointer

End Property

Public Property Get DropDownWidth() As Integer
Attribute DropDownWidth.VB_MemberFlags = "400"

  DropDownWidth = ComboDropDownWidth(cbo.hWnd)

End Property

Public Property Let DropDownWidth(ByVal Width As Integer)
'This property must be set early on e.g. at the start of
'a form load event otherwise it won't work properly
'Width is in pixels

  If Width < ScaleX(cbo.Width, UserControl.ScaleMode, vbPixels) Then
    ComboDropDownWidth(cbo.hWnd) = ScaleX(cbo.Width, UserControl.ScaleMode, vbPixels)
  Else
    ComboDropDownWidth(cbo.hWnd) = Width
  End If

End Property

Public Property Get Enabled() As Boolean

  Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(v As Boolean)

  UserControl.Enabled = v
  cbo.Enabled = v

End Property

Public Property Get Font() As Font

  Set Font = cbo.Font
  UserControl_Resize

End Property

Public Property Set Font(ByVal New_Font As Font)

  Set cbo.Font = New_Font
  PropertyChanged "Font"

End Property

Public Property Get ForeColor() As OLE_COLOR

  ForeColor = cbo.ForeColor

End Property

Public Property Let ForeColor(ByVal Color As OLE_COLOR)

  cbo.ForeColor() = Color
  PropertyChanged "ForeColor"

End Property

Private Function KeyDownLockFull(KeyCode As Integer) As Integer

  Dim lngCount As Long, lngPos As Long
  Dim strText As String

  KeyDownLockFull = KeyCode
  Select Case KeyCode
    Case 46 'Delete
      KeyDownLockFull = 0
      lngPos = cbo.SelStart
      If cbo.SelLength > 0 Then
        If cbo.SelLength = Len(cbo) Then
          KeyDownLockFull = 46
          Exit Function
        End If
        strText = Left(cbo, lngPos) & Right(cbo, Len(cbo) - lngPos - cbo.SelLength)
        lngCount = ComboFindText(cbo.hWnd, strText, True)
        If lngCount > -1 Then
          cbo.ListIndex = lngCount
          cbo.SelStart = lngPos
          cbo.SelLength = 0
          Exit Function
        End If
        strText = Left(cbo, lngPos)
      Else
        If lngPos = Len(cbo) Then Exit Function
        strText = Left(cbo, lngPos) & Right(cbo, Len(cbo) - lngPos - 1)
        lngCount = ComboFindText(cbo.hWnd, strText, True)
        If lngCount > -1 Then
          cbo.ListIndex = lngCount
          cbo.SelStart = lngPos
          cbo.SelLength = 0
          Exit Function
        End If
        strText = Left(cbo, lngPos)
        If strText = "" Then
          cbo = ""
          Exit Function
        End If
      End If
      lngCount = ComboFindText(cbo.hWnd, strText, False)
      If lngCount > -1 Then
        cbo.ListIndex = lngCount
        cbo.SelStart = Len(strText)
        cbo.SelLength = Len(cbo) - cbo.SelStart
      End If
  End Select

End Function

Private Function KeyPressLockFull(KeyAscii As Integer) As Integer

  Dim lngCount As Long, lngPos As Long
  Dim strText As String

  KeyPressLockFull = KeyAscii
  Select Case KeyAscii
    Case 13 'Enter
      If cbo.SelLength > 0 Then
        cbo.SelStart = Len(cbo)
        cbo.SelLength = 0
      End If
      If mblnChangedBefore Then
        mblnChangedBefore = False
        RaiseEvent Changed
      End If
      RaiseEvent Clicked
      Exit Function
    Case 47 'Slash /
      lngPos = cbo.SelStart
      If mblnDateComplete And Left(Right(cbo, Len(cbo) - lngPos + 1), 1) = "/" Then
        KeyPressLockFull = 0
        Exit Function
      End If
    Case 8 'Backspace
      KeyPressLockFull = 0
      lngPos = cbo.SelStart
      If cbo.SelLength > 0 Then
        If cbo.SelLength = Len(cbo) Then
          KeyPressLockFull = 8
          Exit Function
        End If
        strText = Left(cbo, lngPos) & Right(cbo, Len(cbo) - lngPos - cbo.SelLength)
        lngCount = ComboFindText(cbo.hWnd, strText, True)
        If lngCount > -1 Then
          cbo.ListIndex = lngCount
          cbo.SelStart = lngPos
          cbo.SelLength = 0
          Exit Function
        End If
        strText = Left(cbo, lngPos)
        lngCount = ComboFindText(cbo.hWnd, strText, True)
        If lngCount > -1 Then
          cbo.ListIndex = lngCount
          cbo.SelStart = Len(strText)
          cbo.SelLength = 0
          Exit Function
        End If
        lngCount = ComboFindText(cbo.hWnd, strText, False)
        If lngCount > -1 Then
          cbo.ListIndex = lngCount
          cbo.SelStart = Len(strText) - 1
          cbo.SelLength = Len(cbo) - cbo.SelStart
          Exit Function
        End If
      Else
        If lngPos > 0 Then
          strText = Left(cbo, lngPos - 1) & Right(cbo, Len(cbo) - lngPos)
          lngCount = ComboFindText(cbo.hWnd, strText, True)
          If lngCount > -1 Then
            cbo.ListIndex = lngCount
            cbo.SelStart = lngPos - 1
            cbo.SelLength = 0
            Exit Function
          End If
          strText = Left(cbo, lngPos - 1)
          If strText = "" Then
            cbo = ""
            Exit Function
          End If
          lngCount = ComboFindText(cbo.hWnd, strText, True)
          If lngCount > -1 Then
            cbo.ListIndex = lngCount
            cbo.SelStart = Len(strText)
            cbo.SelLength = 0
            Exit Function
          End If
          lngCount = ComboFindText(cbo.hWnd, strText, False)
          If lngCount > -1 Then
            cbo.ListIndex = lngCount
            cbo.SelStart = Len(strText)
            cbo.SelLength = Len(cbo) - cbo.SelStart
            Exit Function
          End If
        Else
          Exit Function
        End If
      End If
  End Select
  
  lngPos = cbo.SelStart
  strText = Left(cbo, lngPos) & Chr(KeyAscii) & Right(cbo, Len(cbo) - lngPos - cbo.SelLength)
  lngCount = ComboFindText(cbo.hWnd, strText, True)
  If lngCount > -1 Then
    KeyPressLockFull = 0
    cbo.ListIndex = lngCount
    If mblnDateComplete And Left(Right(cbo, Len(cbo) - lngPos - 1), 1) = "/" Then
      cbo.SelStart = lngPos + 2
    Else
      cbo.SelStart = lngPos + 1
    End If
    cbo.SelLength = Len(cbo) - cbo.SelStart
    Exit Function
  End If
  strText = Left(cbo, lngPos) & Chr(KeyAscii)
  lngCount = ComboFindText(cbo.hWnd, strText, False)
  If lngCount > -1 Then
    KeyPressLockFull = 0
    cbo.ListIndex = lngCount
    If mblnDateComplete And Left(Right(cbo, Len(cbo) - lngPos - 1), 1) = "/" Then
      cbo.SelStart = Len(strText) + 1
    Else
      cbo.SelStart = Len(strText)
    End If
    cbo.SelLength = Len(cbo) - cbo.SelStart
    Exit Function
  End If
  KeyPressLockFull = 0

End Function

Private Function KeyPressLockOffice(KeyAscii As Integer) As Integer

  Dim lngCount As Long, lngPos As Long
  Dim strText As String

  KeyPressLockOffice = KeyAscii
  Select Case KeyAscii
    Case 13 'Enter
      If cbo.SelLength > 0 Then
        cbo.SelStart = Len(cbo)
        cbo.SelLength = 0
      End If
      If mblnChangedBefore Then
        mblnChangedBefore = False
        RaiseEvent Changed
      End If
      RaiseEvent Clicked
      Exit Function
    Case 32 'Space
      'Don't allow accidental spaces at the start of a combo entry
      If Len(cbo) = 0 Then
        KeyPressLockOffice = 0
        Exit Function
      End If
    Case 47 'Slash /
      lngPos = cbo.SelStart
      If mblnDateComplete And Left(Right(cbo, Len(cbo) - lngPos + 1), 1) = "/" Then
        KeyPressLockOffice = 0
        Exit Function
      End If
    Case 8 'Backspace
      Exit Function
  End Select
  
  lngPos = cbo.SelStart
  strText = Left(cbo, lngPos) & Chr(KeyAscii) & Right(cbo, Len(cbo) - lngPos - cbo.SelLength)
  lngCount = ComboFindText(cbo.hWnd, strText, True)
  If lngCount > -1 Then
    KeyPressLockOffice = 0
    cbo.ListIndex = lngCount
    If mblnDateComplete And Left(Right(cbo, Len(cbo) - lngPos - 1), 1) = "/" Then
      cbo.SelStart = lngPos + 2
    Else
      cbo.SelStart = lngPos + 1
    End If
    cbo.SelLength = 0
    Exit Function
  End If
  strText = Left(cbo, lngPos) & Chr(KeyAscii)
  lngCount = ComboFindText(cbo.hWnd, strText, False)
  If lngCount > -1 Then
    KeyPressLockOffice = 0
    cbo.ListIndex = lngCount
    If mblnDateComplete And Left(Right(cbo, Len(cbo) - lngPos - 1), 1) = "/" Then
      cbo.SelStart = Len(strText) + 1
    Else
      cbo.SelStart = Len(strText)
    End If
    cbo.SelLength = Len(cbo) - cbo.SelStart
    Exit Function
  End If

End Function

Public Property Get ListCount() As Integer
Attribute ListCount.VB_MemberFlags = "400"

  ListCount = cbo.ListCount

End Property

Public Property Get LockLevel() As cbsLockLevel

  LockLevel = menmLockLevel

End Property

Public Property Let LockLevel(v As cbsLockLevel)

  Select Case v
    Case Is < 0
      v = 0
    Case Is > 3
      v = 3
  End Select
  menmLockLevel = v

End Property

Public Sub RemoveItem()

  If cbo.ListIndex > -1 Then
    cbo.RemoveItem cbo.ListIndex
  End If

End Sub

Public Sub SelectCurrent()

  Dim lngRow As Long
  
  lngRow = ComboFindText(cbo.hWnd, cbo, True)
  If lngRow > -1 Then cbo.ListIndex = lngRow

End Sub

Public Sub SelectItem(Index As Variant)

  Dim lngRow As Long
  
  Select Case VarType(Index)
    Case vbInteger, vbLong
      If Index > -1 And Index < cbo.ListCount - 1 Then cbo.ListIndex = Index
    Case vbString
      lngRow = ComboFindText(cbo.hWnd, CStr(Index), True)
      If lngRow > -1 Then cbo.ListIndex = lngRow
  End Select

End Sub

Public Sub SelectValue(Index As Long)

  Dim lngRow As Long
  
  For lngRow = 0 To cbo.ListCount - 1
    If cbo.ItemData(lngRow) = Index Then
      cbo.ListIndex = lngRow
      Exit For
    End If
  Next lngRow

End Sub

Public Property Get Text() As String

  Text = cbo

End Property

Public Sub TextClear()

  cbo = ""
  RaiseEvent KeyUp(0, 0)
  RaiseEvent KeyDown(0, 0)
  RaiseEvent KeyPress(0)

End Sub

Private Sub tmr_Timer()

  tmr.Enabled = False
  If mblnChangedBefore Then
    mblnChangedBefore = False
    RaiseEvent Changed
  End If

End Sub

Private Sub UserControl_EnterFocus()

  If Ambient.UserMode Then
    HookCombo cbo
  End If

End Sub

Private Sub UserControl_ExitFocus()

  Dim lngCount As Long

  Select Case menmLockLevel
    Case cbsOffice
      cbo = Trim$(cbo)
      lngCount = ComboFindText(cbo.hWnd, cbo, True)
      If lngCount > -1 Then cbo.ListIndex = lngCount
  End Select
  If mblnChangedBefore Then
    mblnChangedBefore = False
    RaiseEvent Changed
  End If
  If Ambient.UserMode Then
    UnHookAll
  End If
  DoEvents

End Sub

Private Sub UserControl_InitProperties()

  LockLevel = mCenmLockLevel
  ChangedInterval = mClngChangedInterval
  DateComplete = mCblnDateComplete

End Sub

Private Sub UserControl_Initialize()

  cbo.Font = UserControl.Font

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  LockLevel = PropBag.ReadProperty("LockLevel", mCenmLockLevel)
  ChangedInterval = PropBag.ReadProperty("ChangedInterval", mClngChangedInterval)
  DateComplete = PropBag.ReadProperty("DateComplete", mCblnDateComplete)
  Set cbo.Font = PropBag.ReadProperty("Font", Ambient.Font)
  cbo.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
  cbo.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)

End Sub

Private Sub UserControl_Resize()

  cbo.Move 0, 0, Width
  UserControl.Height = cbo.Height

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  PropBag.WriteProperty "LockLevel", LockLevel, mCenmLockLevel
  PropBag.WriteProperty "ChangedInterval", ChangedInterval, mClngChangedInterval
  PropBag.WriteProperty "DateComplete", DateComplete, mCblnDateComplete
  PropBag.WriteProperty "Font", cbo.Font, Ambient.Font
  PropBag.WriteProperty "BackColor", cbo.BackColor, &H80000005
  PropBag.WriteProperty "ForeColor", cbo.ForeColor, &H80000008

End Sub

Public Property Get Value() As Variant
Attribute Value.VB_MemberFlags = "400"

  If cbo.ListIndex = -1 Then
    Value = Null
  Else
    Value = cbo.ItemData(cbo.ListIndex)
  End If

End Property

Public Property Let Value(v As Variant)

  Dim lngCount As Long
  
  If IsNull(v) Then
    mblnIgnore = True
    cbo = ""
    cbo.ListIndex = -1
    mblnIgnore = False
  Else
    For lngCount = 0 To cbo.ListCount - 1
      If cbo.ItemData(lngCount) = v Then
        cbo.ListIndex = lngCount
        Exit Property
      End If
    Next
    cbo.ListIndex = -1
  End If

End Property
