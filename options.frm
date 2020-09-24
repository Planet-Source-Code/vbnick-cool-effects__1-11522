VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Options"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   LinkTopic       =   "Form2"
   ScaleHeight     =   4620
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   1260
      Left            =   2520
      TabIndex        =   10
      Top             =   2025
      Width           =   2490
   End
   Begin VB.TextBox Text3 
      Height          =   435
      Left            =   2520
      TabIndex        =   8
      Top             =   1530
      Width           =   2460
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   30
      TabIndex        =   5
      Top             =   1560
      Width           =   2445
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   30
      TabIndex        =   4
      Top             =   1215
      Width           =   2460
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2505
      TabIndex        =   2
      Top             =   480
      Width           =   2460
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2505
      TabIndex        =   1
      Top             =   60
      Width           =   2475
   End
   Begin VB.Label Label7 
      Caption         =   "Record Path"
      Height          =   270
      Left            =   195
      TabIndex        =   12
      Top             =   975
      Width           =   1920
   End
   Begin VB.Label Label6 
      Caption         =   "Number of Square per Second"
      Height          =   405
      Left            =   60
      TabIndex        =   11
      Top             =   150
      Width           =   2355
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "NAME"
      Height          =   315
      Left            =   2595
      TabIndex        =   9
      Top             =   1230
      Width           =   2205
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   285
      Left            =   675
      TabIndex        =   7
      Top             =   3015
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "Size:"
      Height          =   360
      Left            =   75
      TabIndex        =   6
      Top             =   3015
      Width           =   1155
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   15
      X2              =   5040
      Y1              =   885
      Y2              =   885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   -15
      X2              =   5010
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Label Label2 
      Caption         =   "Recording Time"
      Height          =   255
      Left            =   75
      TabIndex        =   3
      Top             =   555
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   945
      Left            =   15
      TabIndex        =   0
      Top             =   3645
      Width           =   4965
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Sub Command1_Click()
On Local Error GoTo hell
ressay = Val(Text1) * Val(Text2)
Form1.Timer13.Interval = 1000 / Val(Text1)
Form1.Timer12.Interval = 1000 / Val(Text1)
Form1.Visible = False
dosya = Text3
On Local Error GoTo yanlisad
MkDir dosya & "1111" ' test the file name , dosya = file
MkDir dosya & "1111"
Open dosya & ".kaya" For Output As #1
 Dim x
 x = Text1.Text
 Write #1, x
 x = Text2.Text
 Write #1, x
 Close #1
 Unload Me
 Exit Sub
hell:
 Exit Sub
yanlisad:
 MsgBox ("Wrong File Name")
 Exit Sub
End Sub

Private Sub Dir1_Change()
ChDir Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
ChDrive Dir1.Path
End Sub

Private Sub form_load()
 Const PLANES = 14
 Const BITSPIXEL = 12
 Const colorres = 108
 Label4 = (Screen.Width / 2 / Screen.TwipsPerPixelX) * (Screen.Height / 2 / Screen.TwipsPerPixelY) * Val(Text1) * Val(Text2) * Log((GetDeviceCaps(hdc, PLANES) * 2 ^ GetDeviceCaps(hdc, BITSPIXEL))) / Log(2) / 8
 Label1.Caption = "Now, everything that you do will be recorded in the time you have chosen. To start hit the Start button. After program ends you are going to able to watch what you have done.."
End Sub
Private Sub text1_change()
 Const PLANES = 14
 Const BITSPIXEL = 12
 Const colorres = 108
  Label4 = (Screen.Width / 2 / Screen.TwipsPerPixelX) * (Screen.Height / 2 / Screen.TwipsPerPixelY) * Val(Text1) * Val(Text2) * Log((GetDeviceCaps(hdc, PLANES) * 2 ^ GetDeviceCaps(hdc, BITSPIXEL))) / Log(2) / 8
End Sub

Private Sub Text2_Change()
Const PLANES = 14
 Const BITSPIXEL = 12
 Const colorres = 108
  Label4 = (Screen.Width / 2 / Screen.TwipsPerPixelX) * (Screen.Height / 2 / Screen.TwipsPerPixelY) * Val(Text1) * Val(Text2) * Log((GetDeviceCaps(hdc, PLANES) * 2 ^ GetDeviceCaps(hdc, BITSPIXEL))) / Log(2) / 8
End Sub
