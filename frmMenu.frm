VERSION 5.00
Begin VB.Form frmMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   3225
   ClientTop       =   3285
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   900
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
frmMenu.Hide
End Sub
