VERSION 5.00
Object = "{B7EFA9D7-1570-4277-81F4-82A58D828123}#1.0#0"; "ComboSelectCtrl.ocx"
Begin VB.Form TestForm 
   Caption         =   "ComboSelect Test Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDropDown 
      Caption         =   "Drop Down"
      Height          =   435
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin ComboSelectCtrl.ComboSelect cbsLockLevel 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComboSelectCtrl.ComboSelect cbs 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1260
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      ChangedInterval =   1000
      DateComplete    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   1260
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Written by James Haydock - please vote!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2820
      Width           =   4335
   End
   Begin VB.Label Label4 
      Caption         =   "Test form to demonstrate the use of the ComboSelect control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Lock Level"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Normal Combo"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ComboSelect"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   1215
   End
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbs_Change()

  With cbs
    Debug.Print "Change " & .Value & " " & .Text
  End With

End Sub

Private Sub cbs_Changed()

  With cbs
    Debug.Print "Changed " & .Value & " " & .Text
  End With

End Sub

Private Sub cbs_Click()

  With cbs
    Debug.Print "Click " & .Value & " " & .Text
  End With

End Sub

Private Sub cbsLockLevel_Change()

  cbs.Enabled = False

End Sub

Private Sub cbsLockLevel_Click()

  With cbs
    .Enabled = True
    .LockLevel = cbsLockLevel.Value
  End With

End Sub

Private Sub cmdDropDown_Click()

  cbs.DropDown = True

End Sub

Private Sub Form_Load()

  With cbs
    .DropDownWidth = 400
    .AddItem "a", 1
    .AddItem "Ab", 2
    .AddItem "abcde", 3
    .AddItem "Abcdefghi", 4
    .AddItem "abcee", 5
    .AddItem "Aaa", 6
    .AddItem "aaaa", 7
    .AddItem "bcde", 8
    .AddItem "bde", 9
    .AddItem "01/01/1999", 101
    .AddItem "02/01/1999", 102
    .AddItem "03/01/1999", 103
    .AddItem "04/01/1999", 104
    .AddItem "01/02/1999", 105
    .AddItem "01/03/1999", 106
    .AddItem "02/02/1999", 107
    .AddItem "01/01/1998", 108
    .AddItem "02/01/1998", 109
    .AddItem "//", 110
  End With
  
  With cbsLockLevel
    .AddItem "cbsNone", cbsNone
    .AddItem "cbsOffice", cbsOffice
    .AddItem "cbsFull", cbsFull
    .SelectValue cbs.LockLevel
  End With
  
  With cbo
    .AddItem "a"
    .AddItem "Ab"
    .AddItem "abcde"
    .AddItem "Abcdefghi"
    .AddItem "abcee"
    .AddItem "Aaa"
    .AddItem "aaaa"
    .AddItem "bcde"
    .AddItem "bde"
  End With

End Sub
