VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "List all Programs"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close All Windows"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim strin As String, i As Boolean, j As Integer
For j = 1 To List1.ListCount
List1.ListIndex = List1.ListIndex + 1
strin = List1.Text
If ((strin <> "Form1") And (strin <> "Program Manager") And (strin <> "Project1")) Then
i = CloseApplication(strin)
End If
Next j
End Sub

Private Sub Command2_Click()
Dim blnRtn As Boolean
List1.Clear
blnRtn = EnumWindows(AddressOf EnumCallBack, 0)


End Sub



Private Sub Form_Load()
Dim blnRtn As Boolean, c As Integer
ReDim holder(0 To 1) As Variant
List1.Clear
blnRtn = EnumWindows(AddressOf EnumCallBack, 0)
ReDim Preserve holder(0 To List1.ListCount) As Variant
For c = 1 To List1.ListCount
    List1.ListIndex = List1.ListIndex + 1
    holder(c - 1) = List1.Text
Next c

End Sub

