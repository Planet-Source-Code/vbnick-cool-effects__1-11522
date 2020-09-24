VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PLEASE READ TO AVOID UNWANTED SIDE EFFECTS!!!!!!!!!"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Default         =   -1  'True
      Height          =   195
      Left            =   540
      TabIndex        =   5
      Top             =   5940
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   300
      Left            =   2205
      TabIndex        =   3
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Timer Timer2 
      Left            =   1005
      Top             =   5250
   End
   Begin VB.Timer Timer1 
      Left            =   1860
      Top             =   5205
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Please Read!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   2025
      TabIndex        =   4
      Top             =   5040
      Width           =   3075
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   $"plsread.frx":0000
      Height          =   2430
      Left            =   255
      TabIndex        =   2
      Top             =   2595
      Width           =   6705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE READ FIRST"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   48
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2355
      Left            =   300
      TabIndex        =   0
      Top             =   210
      Width           =   6540
   End
   Begin VB.Label Label2 
      Height          =   2385
      Left            =   240
      TabIndex        =   1
      Top             =   225
      Width           =   6675
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Load Form1
'Form1.Show
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.Visible = False
Label4.Visible = True
End Sub

Private Sub Form_Load()
Timer1.Interval = 10
Timer1.Enabled = True
Timer2.Interval = 20
Label4.Visible = False
End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Load Form1
Form1.Show

End Sub

Private Sub Form_Terminate()
Load Form1
Form1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Load Form1
Form1.Show
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.Visible = True
End Sub

Private Sub Timer1_Timer()
Static Col1, Col2, Col3 As Integer
Static C1, C2, C3 As Integer
If (Col1 = 0 Or Col1 = 250) And (Col2 = 0 Or Col2 = 250) And (Col3 = 0 Or Col3 = 250) Then
C1 = Int(Rnd * 3)
C2 = Int(Rnd * 3)
C3 = Int(Rnd * 3)
End If
If C1 = 1 And Col1 <> 0 Then Col1 = Col1 - 10
If C2 = 1 And Col2 <> 0 Then Col2 = Col2 - 10
If C3 = 1 And Col3 <> 0 Then Col3 = Col3 - 10
If C1 = 2 And Col1 <> 250 Then Col1 = Col1 + 10
If C2 = 2 And Col2 <> 250 Then Col2 = Col2 + 10
If C3 = 2 And Col3 <> 250 Then Col3 = Col3 + 10
Label1.ForeColor = RGB(Col1, Col2, Col3)

End Sub


Private Sub Timer2_Timer()
Static Col1, Col2, Col3 As Integer
Static C1, C2, C3 As Integer
If (Col1 = 0 Or Col1 = 250) And (Col2 = 0 Or Col2 = 250) And (Col3 = 0 Or Col3 = 250) Then
C1 = Int(Rnd * 3)
C2 = Int(Rnd * 3)
C3 = Int(Rnd * 3)
End If
If C1 = 1 And Col1 <> 0 Then Col1 = Col1 - 10
If C2 = 1 And Col2 <> 0 Then Col2 = Col2 - 10
If C3 = 1 And Col3 <> 0 Then Col3 = Col3 - 10
If C1 = 2 And Col1 <> 250 Then Col1 = Col1 + 10
If C2 = 2 And Col2 <> 250 Then Col2 = Col2 + 10
If C3 = 2 And Col3 <> 250 Then Col3 = Col3 + 10
Label2.BackColor = RGB(Col1, Col2, Col3)
End Sub
