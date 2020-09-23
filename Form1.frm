VERSION 5.00
Object = "*\AHotTextBoxControl.vbp"
Begin VB.Form Form1 
   Caption         =   "Hot Textbox Demonstration"
   ClientHeight    =   3405
   ClientLeft      =   1740
   ClientTop       =   1470
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   227
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
   Begin VB.HScrollBar SC 
      Height          =   195
      Index           =   2
      Left            =   60
      Max             =   255
      TabIndex        =   8
      Top             =   3120
      Value           =   255
      Width           =   6135
   End
   Begin VB.HScrollBar SC 
      Height          =   195
      Index           =   1
      Left            =   60
      Max             =   255
      TabIndex        =   7
      Top             =   2940
      Width           =   6135
   End
   Begin VB.HScrollBar SC 
      Height          =   195
      Index           =   0
      Left            =   60
      Max             =   255
      TabIndex        =   6
      Top             =   2760
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Text            =   "This is a ~Hot Spot Text Box~ which allows ~Multiple~ hot spots, and word wrap. All done in Visual Basic! ~Want to know more?~ *"
      Top             =   300
      Width           =   6375
   End
   Begin HotTextBoxControl.HotTextBox HotTextBox1 
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   960
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   1296
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ControlString   =   "This is a ~Hot Spot Text Box~ which allows ~Multiple~ hot spots, and word wrap. All done in Visual Basic! ~Want to know more?~ *"
      Picture         =   "Form1.frx":0000
   End
   Begin VB.Label Label4 
      Caption         =   "Colour:"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Choose a Picture:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   1800
      Width           =   2835
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   1740
      Picture         =   "Form1.frx":1D0A
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   3
      Left            =   1380
      Picture         =   "Form1.frx":44AC
      Top             =   2040
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   1
      Left            =   420
      Picture         =   "Form1.frx":61A6
      Top             =   2040
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   960
      Picture         =   "Form1.frx":6268
      Top             =   2040
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":7F62
      Top             =   2040
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Hot Text Box:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   720
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Control String:"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   2835
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HotTextBox1_HotSpotClick(Index As Long)
MsgBox "This Hot Spot " & Index
End Sub

Private Sub Image1_Click(Index As Integer)
Set HotTextBox1.Picture = Image1(Index).Picture
HotTextBox1.Refresh
End Sub

Private Sub SC_Change(Index As Integer)
SC_Scroll (Index)
End Sub

Private Sub SC_Scroll(Index As Integer)
HotTextBox1.HotspotColor = RGB(SC(0).Value, SC(1).Value, SC(2).Value)
HotTextBox1.Refresh
End Sub

Private Sub Text1_Change()
HotTextBox1.ControlString = Text1.Text
End Sub
