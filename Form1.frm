VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Age"
      Height          =   3255
      Left            =   7920
      TabIndex        =   8
      Top             =   2760
      Width           =   5895
      Begin VB.Label Label7 
         Height          =   375
         Left            =   2400
         TabIndex        =   14
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label6 
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label5 
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Days"
         Height          =   495
         Left            =   480
         TabIndex        =   11
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Years"
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   615
      Left            =   5520
      TabIndex        =   7
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   615
      Left            =   3480
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   1440
      TabIndex        =   5
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter DOB"
      Height          =   1575
      Left            =   1440
      TabIndex        =   1
      Top             =   2760
      Width           =   5655
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   4320
         TabIndex        =   4
         Text            =   "Year"
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Text            =   "Month"
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Text            =   "Date"
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Age calculator"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4560
      TabIndex        =   0
      Top             =   840
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Integer
For i = 1 To 31 Step 1
Combo1.AddItem i
Next i

For i = 1 To 12 Step 1
Combo2.AddItem i
Next i

For i = 1900 To 2019 Step 1
Combo3.AddItem i
Next i

Form1.WindowState = vbMaximized
End Sub

Private Sub Command2_Click()
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
End Sub

Private Sub Command3_Click()
Dim reply As Integer
reply = MsgBox("Are you sure to exit?", vbYesNo + vbQuestion, "Closing...!")
If reply = vbYes Then
End
End If
End Sub
Private Sub Command1_Click()
Dim cyr As Double
cyr = Year(Now)
Dim cmo As Double
cmo = Month(Now)
Dim cd As Double
cd = Day(Now)

Dim byr As Double
byr = Combo3.Text
Dim bmo As Double
bmo = Combo2.Text
Dim bd As Double
bd = Combo1.Text

If cyr > byr Then
    Label5.Caption = cyr - byr
Else
    Label5.Caption = byr - cyr
End If

If cmo > bmo Then
    Label6.Caption = cmo - bmo
Else
    Label6.Caption = bmo - cmo
End If

If cd > bd Then
    Label7.Caption = cd - bd
Else
    Label7.Caption = bd - cd
End If
End Sub
