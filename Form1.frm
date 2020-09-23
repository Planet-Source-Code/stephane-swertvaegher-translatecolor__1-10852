VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Show RealColor of the form"
      Height          =   645
      Left            =   2475
      TabIndex        =   3
      Top             =   225
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Color of the form"
      Height          =   645
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   1815
   End
   Begin VB.Label Label1 
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1935
      Width           =   3885
   End
   Begin VB.Label Label1 
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   1305
      Width           =   3885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'API for translating system colors to 'normal' colors
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Dim SC, NC&
  
  'This is the function that translates system colors to normal (RGB) colors
  Private Function RealColor(ByVal Color As OLE_COLOR) As Long
     Dim Col As Long
     Col = TranslateColor(Color, 0, RealColor)
  End Function

'Change the color of the form (system color)
Private Sub Command1_Click()
SC = SC + 1
If SC = 25 Then SC = 0
Form1.BackColor = (&H80000000 Or SC)
Label1(0).Caption = "System color: " & Hex(Form1.BackColor)
End Sub

Private Sub Command2_Click()
'translate to normal (RGB) color
NC = Form1.BackColor
NC = RealColor(NC)
Label1(1).Caption = "Normal RGB Color: " & Hex(NC)
End Sub

Private Sub Form_Load()
SC = 15
Form1.BackColor = &H8000000F 'force default system color
Label1(0).Caption = "System color: " & Hex(Form1.BackColor)
Label1(1).Caption = ""
End Sub
