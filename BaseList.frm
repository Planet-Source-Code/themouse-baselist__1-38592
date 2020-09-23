VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Digits"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8670
   Icon            =   "BaseList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Convert"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8160
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IntegralHeight  =   0   'False
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   8160
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   6960
      Width           =   8415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   120
      TabIndex        =   19
      Top             =   6720
      Width           =   8415
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00EFEFEF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   6360
      Width           =   3615
   End
   Begin VB.ListBox List5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      IntegralHeight  =   0   'False
      Left            =   6960
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      IntegralHeight  =   0   'False
      Left            =   5280
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      IntegralHeight  =   0   'False
      Left            =   3600
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      IntegralHeight  =   0   'False
      Left            =   1920
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   -120
      ScaleHeight     =   615
      ScaleWidth      =   9255
      TabIndex        =   10
      Top             =   0
      Width           =   9255
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Octal:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   15
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Hexadecimal:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   14
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Decimal:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "ASCII:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Binary:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do It"
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "01100001"
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2002 Binary-Craft"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   18
      Top             =   6360
      Width           =   3135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Display in Font:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Char: "
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Decimal:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Binary digit:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hh As Integer

Private Sub Combo1_Change()
List1.FontName = Combo1.List(Combo1.ListIndex)
List2.FontName = Combo1.List(Combo1.ListIndex)
List3.FontName = Combo1.List(Combo1.ListIndex)
List4.FontName = Combo1.List(Combo1.ListIndex)
List5.FontName = Combo1.List(Combo1.ListIndex)
End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub

Private Sub Combo1_DropDown()
Combo1_Change
End Sub

Private Sub Combo1_GotFocus()
Combo1_Change
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Combo1_Change
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
Combo1_Change
End Sub

Private Sub Combo1_LostFocus()
Combo1_Change
End Sub

Private Sub Combo1_Scroll()
Combo1_Change
End Sub

Private Sub Command1_Click()
Dim t As Variant, a As Double
s = ""
For x = 1 To Len(Text1)
s = s & Mid(Text1.Text, x, 1) & ","
Next

t = Split(Left(s, Len(s) - 1), ",")

a = 0
a = a + t(7) * 1
a = a + t(6) * 2
a = a + t(5) * 4
a = a + t(4) * 8
a = a + t(3) * 16
a = a + t(2) * 32
a = a + t(1) * 64
a = a + t(0) * 128

List1.AddItem Text1.Text
List2.AddItem Chr(a)
List3.AddItem a
List4.AddItem Module2.ConvertBase(CInt(a), 10, 16)
List5.AddItem Module2.ConvertBase(CInt(a), 10, 8)
End Sub

Private Sub Command2_Click()
'ascii - bin
'bin - ascii
'dec oct
'dec hex
'oct dec
'oct hex
'hex oct
'hex dec
Select Case Combo2.ListIndex
Case 0
    Text2.Text = Module1.TextToBinary(Text2.Text)
Case 1
    Text2.Text = Module1.BinaryToText(Text2.Text)
Case 2
        Text2.Text = Module2.ConvertBase(Text2.Text, 10, 8)
Case 3
        Text2.Text = Module2.ConvertBase(Text2.Text, 10, 16)
Case 4
        Text2.Text = Module2.ConvertBase(Text2.Text, 8, 10)
Case 5
        Text2.Text = Module2.ConvertBase(Text2.Text, 8, 16)
Case 6
        Text2.Text = Module2.ConvertBase(Text2.Text, 16, 8)
Case 7
        Text2.Text = Module2.ConvertBase(Text2.Text, 16, 10)
End Select
End Sub

Private Sub Command3_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
hh = List1.Height
'MsgBox Module1.TextToBinary("a")
'End

Me.Show
DoEvents
For x = 0 To 255
Text1.Text = Module1.TextToBinary(Chr(x))
Command1_Click
DoEvents
Me.Refresh
List1.ListIndex = List1.ListCount - 1
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
List4.ListIndex = List1.ListIndex
List5.ListIndex = List1.ListIndex
Next


For x = 0 To Screen.FontCount
Combo1.AddItem Screen.Fonts(x)
Next
Combo1.ListIndex = 0

With Combo2
    .AddItem "ASCII - Binary"
    .AddItem "Binary - ASCII"
    
    .AddItem "Decimal - Octal"
    .AddItem "Decimal - Hexadecimal"
    .AddItem "Octal - Decimal"
    .AddItem "Octal - Hexidecimal"
    .AddItem "Hexadecimal - Octal"
    .AddItem "Hexadecimal - Decimal"
.ListIndex = 0
End With

End Sub

Private Sub List1_Click()
List1.ListIndex = List1.ListIndex
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
List4.ListIndex = List1.ListIndex
List5.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
List2.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
List4.ListIndex = List2.ListIndex
List5.ListIndex = List2.ListIndex
End Sub

Private Sub List3_Click()
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
List3.ListIndex = List3.ListIndex
List4.ListIndex = List3.ListIndex
List5.ListIndex = List3.ListIndex
End Sub

Private Sub List4_Click()
List1.ListIndex = List4.ListIndex
List2.ListIndex = List4.ListIndex
List3.ListIndex = List4.ListIndex
List4.ListIndex = List4.ListIndex
List5.ListIndex = List4.ListIndex
End Sub

Private Sub List5_Click()
List1.ListIndex = List5.ListIndex
List2.ListIndex = List5.ListIndex
List3.ListIndex = List5.ListIndex
List4.ListIndex = List5.ListIndex
List5.ListIndex = List5.ListIndex
End Sub
