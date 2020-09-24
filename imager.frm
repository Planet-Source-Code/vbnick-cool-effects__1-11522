VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "COOLLL EFFECTS AND SCREEN CAM BY KAYHAN TANRISEVEN"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check12 
      Caption         =   "Camera"
      Height          =   345
      Left            =   4815
      TabIndex        =   21
      Top             =   3330
      Width           =   1770
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   45
      TabIndex        =   20
      Top             =   3945
      Width           =   6585
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   22
         Text            =   "imager.frx":0000
         Top             =   60
         Width           =   6480
      End
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Shaded Screen"
      Height          =   345
      Left            =   4815
      TabIndex        =   19
      Top             =   2985
      Width           =   1770
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Flag"
      Height          =   345
      Left            =   4815
      TabIndex        =   18
      Top             =   2670
      Width           =   1770
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Frame"
      Height          =   195
      Left            =   3195
      TabIndex        =   17
      Top             =   3720
      Width           =   825
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Wolf"
      Height          =   195
      Left            =   3195
      TabIndex        =   16
      Top             =   3390
      Width           =   1350
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Black Hole"
      Height          =   195
      Left            =   3195
      TabIndex        =   15
      Top             =   3120
      Width           =   1350
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Pattern"
      Height          =   195
      Left            =   3195
      TabIndex        =   14
      Top             =   2850
      Width           =   1350
   End
   Begin VB.CheckBox Check5 
      Caption         =   "black - white"
      Height          =   195
      Left            =   3195
      TabIndex        =   13
      Top             =   2520
      Width           =   1350
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Cross Change"
      Height          =   195
      Left            =   3195
      TabIndex        =   12
      Top             =   2220
      Width           =   1350
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Drawing"
      Height          =   195
      Left            =   3195
      TabIndex        =   11
      Top             =   1905
      Width           =   1350
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Walking Screen"
      Height          =   195
      Left            =   3195
      TabIndex        =   10
      Top             =   1605
      Width           =   1515
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Walking Frame"
      Height          =   195
      Left            =   3195
      TabIndex        =   9
      Top             =   1335
      Width           =   1605
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3570
      Top             =   6540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture6 
      Height          =   570
      Left            =   4875
      ScaleHeight     =   510
      ScaleWidth      =   1575
      TabIndex        =   8
      Top             =   705
      Width           =   1635
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      Height          =   630
      Index           =   0
      Left            =   4860
      ScaleHeight     =   570
      ScaleWidth      =   1605
      TabIndex        =   7
      Top             =   1935
      Width           =   1665
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      Height          =   615
      Left            =   4860
      ScaleHeight     =   555
      ScaleWidth      =   1590
      TabIndex        =   6
      Top             =   1290
      Width           =   1650
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   1230
      Left            =   3195
      Picture         =   "imager.frx":0146
      ScaleHeight     =   1170
      ScaleWidth      =   1575
      TabIndex        =   5
      Top             =   60
      Width           =   1635
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   30
      Picture         =   "imager.frx":0588
      ScaleHeight     =   2115
      ScaleWidth      =   6570
      TabIndex        =   4
      Top             =   5190
      Width           =   6600
   End
   Begin VB.Timer Timer13 
      Left            =   3750
      Top             =   6735
   End
   Begin VB.Timer Timer12 
      Left            =   5640
      Top             =   6765
   End
   Begin VB.Timer Timer11 
      Left            =   4965
      Top             =   6585
   End
   Begin VB.Timer Timer10 
      Left            =   2685
      Top             =   6495
   End
   Begin VB.Timer Timer9 
      Left            =   1260
      Top             =   6810
   End
   Begin VB.Timer Timer8 
      Left            =   3165
      Top             =   6630
   End
   Begin VB.Timer Timer7 
      Left            =   2085
      Top             =   6690
   End
   Begin VB.Timer Timer6 
      Left            =   315
      Top             =   6690
   End
   Begin VB.Timer Timer5 
      Left            =   4260
      Top             =   6510
   End
   Begin VB.Timer Timer4 
      Left            =   3135
      Top             =   6855
   End
   Begin VB.Timer Timer3 
      Left            =   2325
      Top             =   6825
   End
   Begin VB.Timer Timer2 
      Left            =   1470
      Top             =   6390
   End
   Begin VB.Timer Timer1 
      Left            =   630
      Top             =   6570
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   645
      Left            =   4890
      ScaleHeight     =   585
      ScaleWidth      =   1560
      TabIndex        =   3
      Top             =   45
      Width           =   1620
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OPEN THE MOVIE( YOU MUST HAVE SAVED BEFORE)"
      Height          =   600
      Left            =   30
      TabIndex        =   1
      Top             =   3045
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2955
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   2865
   End
   Begin VB.Label Label2 
      Caption         =   "CLICK CAMERA FOR OPTIONS"
      Height          =   315
      Left            =   4305
      TabIndex        =   23
      Top             =   3660
      Width           =   2340
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   585
      Left            =   15
      TabIndex        =   2
      Top             =   3675
      Width           =   2910
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'COPRYRIGHT BY KAYHAN TANRISEVEN 2000
'THANX TO MR. KARAGULLE AND MY SISTER FOR THEIR BIG HELP
'THAX TO MY GF TO MAKE ME HAVE SOME BRAKE.>>
'ANY QUESTIONS OR COMMENTS TO kayahan@presidency.com
' OR JUST COMMENT ME....DONT FORGET TO VOTE.......
Option Explicit

Private Sub Check1_Click()
Picture2.Refresh
Timer1.Enabled = Check1.Value
Picture2.SetFocus
End Sub

Private Sub Check10_Click()
List1.Refresh
Timer10.Enabled = Check10.Value
List1.SetFocus
End Sub

Private Sub Check11_Click()
List1.Refresh
Timer11.Enabled = Check11.Value
List1.SetFocus
End Sub

Private Sub Check12_Click()
If Check12.Value Then
Form2.Show 1
Timer12.Enabled = Check12.Value
Label1.Visible = True
Timer13.Enabled = False
End If
End Sub

Private Sub Check2_Click()
Picture2.Refresh
Timer2.Enabled = Check2.Value
Picture2.SetFocus
End Sub

Private Sub Check3_Click()
Picture2.Refresh
Timer3.Enabled = Check3.Value
Picture2.SetFocus
End Sub

Private Sub Check4_Click()
Picture2.Refresh
Timer4.Enabled = Check4.Value
Picture2.SetFocus
End Sub

Private Sub Check5_Click()
Picture2.Refresh
Timer5.Enabled = Check5.Value
Picture2.SetFocus
End Sub

Private Sub Check6_Click()
Picture2.Refresh
Timer6.Enabled = Check6.Value
Picture2.SetFocus
End Sub
Private Sub Check7_Click()
Picture2.Refresh
Timer7.Enabled = Check7.Value
Picture2.SetFocus
End Sub
Private Sub Check8_Click()
Picture2.Refresh
Timer8.Enabled = Check8.Value
Picture2.SetFocus
End Sub
Private Sub Check9_Click()
Picture2.Refresh
Timer9.Enabled = Check9.Value
Picture2.SetFocus
End Sub

Private Sub Command1_Click()
On Local Error Resume Next
CommonDialog1.FileName = "*.kaya"
CommonDialog1.Action = 1
If CommonDialog1.FileName = "" Then Exit Sub
dosya = Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4)
Open CommonDialog1.FileName For Input As #1
Dim x
Input #1, x
Timer13.Interval = 1000 / Val(x)
Input #1, x
ressay = Val(x)
Close #1
Timer13.Enabled = True
Picture6.Visible = True
End Sub


Private Sub form_load()
Timer1.Interval = 10
Timer1.Enabled = False
Timer10.Interval = 500
Timer10.Enabled = False
Timer11.Interval = 500
Timer11.Enabled = False
Timer12.Interval = 1000
Timer12.Enabled = False
Timer13.Interval = 1000
Timer13.Enabled = False
Timer2.Interval = 10
Timer3.Interval = 10
Timer3.Enabled = False
Timer4.Interval = 10
Timer4.Enabled = False
Timer5.Interval = 10
Timer5.Enabled = False
Timer6.Interval = 10
Timer6.Enabled = False
Timer7.Interval = 10
Timer7.Enabled = False
Timer8.Interval = 10
Timer8.Enabled = False
Timer9.Interval = 10
Timer9.Enabled = False
Picture1.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5(0).Visible = False
Picture6.Visible = False
Picture3.ScaleMode = 3
Picture5(0).ScaleMode = 3
Picture6.ScaleMode = 3

Dim i
ressay = 50
For i = 1 To 15
List1.AddItem Space(5) & i
Next
Picture5(0).Width = Screen.Width / 2
Picture5(0).Height = Screen.Height / 2
Picture6.Width = Picture5(0).Width
Picture6.Height = Picture5(0).Height
Picture1.Width = Screen.Width
Picture1.Height = Screen.Height
End Sub

Private Sub Picture6_Click()
Picture6.Visible = False
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub Timer1_Timer()
Dim hd, t, hw, w, h
Dim koor As RECT
Static x, y, maxx, maxy, dx, dy
w = 10
h = 10
If IsEmpty(dx) Then
ScaleMode = 3
dx = w
dy = 0 'h
maxx = 640
maxy = 480
End If
hw = GetFocus() 'active window handle
GetClientRect hw, koor
maxx = koor.Right - koor.Left
maxy = koor.Bottom - koor.Top
x = x + dx
y = y + dy
If x > (maxx - w) Then x = maxx - w: dx = -dx: dy = h: dx = 0
If x < 0 Then x = 0: dx = -dx: dy = -h: dx = 0
If y > maxy - h Then y = maxy - h: dy = -dy: dy = 0: dx = -w
If y < 0 Then y = 0: dy = -dy: dy = 0: dx = w
hd = GetDC(hw)
t = BitBlt(hd, x, y, w, h, hd, x, y, &H550009) 'destinvert
t = ReleaseDC(GetDesktopWindow(), hd)
DoEvents
End Sub
Sub timer10_timer()
Dim hd, t, hw, w, h
Dim koor As RECT
Static x, y, maxx, maxy, dx, dy
w = Picture3.ScaleWidth
h = Picture3.ScaleHeight
hw = GetFocus() 'active window handle
GetClientRect hw, koor
maxx = koor.Right - koor.Left
maxy = koor.Bottom - koor.Top
hd = GetDC(hw)
t = BitBlt(hd, 0, 0, w, h, Picture3.hdc, 0, 0, &H8800C6) 'srcand
t = ReleaseDC(GetDesktopWindow(), hd)
DoEvents
End Sub
Sub timer11_timer()
Dim hd, t, hw, w, h
Dim koor As RECT
Static x, y, maxx, maxy, dx, dy
w = Picture3.ScaleWidth
h = Picture3.ScaleHeight
hw = GetDesktopWindow() 'Desktop Handle No
GetClientRect hw, koor
maxx = koor.Right - koor.Left
maxy = koor.Bottom - koor.Top
hd = GetDC(hw)
t = StretchBlt(Picture4.hdc, 0, 0, maxx / 5, maxy / 5, hd, 0, 0, maxx, maxy, &HCC0020) 'srccopy
t = ReleaseDC(GetDesktopWindow(), hd)
hw = GetFocus() 'desktop handle
hd = GetDC(hw)
t = BitBlt(hd, 0, 0, maxx / 5, maxy / 5, Picture4.hdc, 0, 0, &HCC0020) 'srccopy
t = ReleaseDC(GetDesktopWindow(), hd)
DoEvents
End Sub
Sub timer12_timer() 'camera
Dim d, t, hw, w, h, hd
Dim koor As RECT
Static i
If IsEmpty(i) Then i = 0
Label1 = i
Static x, y, maxx, maxy, dx, dy
w = Picture5(0).ScaleWidth
h = Picture5(0).ScaleHeight
hw = GetDesktopWindow() 'desktop handle
GetClientRect hw, koor
maxx = koor.Right - koor.Left
maxy = koor.Bottom - koor.Top
hd = GetDC(hw)
't = StretchBlt(Picture5(i).hdc, 0, 0, maxx / 2, maxy / 2, hd, 0, 0, maxx, maxy, &HCC0020)
t = StretchBlt(Picture5(0).hdc, 0, 0, maxx / 2, maxy / 2, hd, 0, 0, maxx, maxy, &HCC0020)
t = ReleaseDC(GetDesktopWindow(), hd)
SavePicture Picture5(0).Image, dosya & i & ".bmp"
DoEvents
i = i + 1
If i = (ressay + 1) Then
Timer12.Enabled = False
i = 0
Check12.Value = False
Label1.Visible = False
Form1.Visible = True
MsgBox ("Screen Camera Has ended!")
End If
End Sub
Sub timer13_timer()
On Local Error Resume Next
Static i
Dim x
i = (i + 1) Mod (ressay + 1)
Dim maxx, maxy
maxx = Picture5(0).ScaleWidth
maxy = Picture5(0).ScaleHeight
Picture6.Picture = LoadPicture(dosya & i & ".bmp")
End Sub
Sub Timer2_Timer()
Dim t, hw, w, h
Dim koor As RECT
Static x, y, maxx, maxy, dx, dy, hd
w = 25
h = 25
If IsEmpty(dx) Then
ScaleMode = 3
dx = w
dy = 10
maxx = 640
maxy = 480
End If
If hw <> GetFocus() Then
hw = GetFocus() 'active control handle
y = 0: x = 0
End If
GetClientRect hw, koor
maxx = koor.Right - koor.Left
maxy = koor.Bottom - koor.Top
h = maxy
x = x + dx
If x >= maxx Then
x = 0: dx = 0: y = y + dy: 'dx = -dx
End If
If x < 0 Then
x = 0: dx = 0: dx = 0: y = y + dy
End If
If y > maxy - h Then
y = maxy - h: dy = -dy: dx = -w
End If
If y < 0 Then
y = 0: dy = -dy: dx = w
End If
hd = GetDC(hw)
t = BitBlt(Picture1.hdc, 0, 0, w, h, hd, maxx - w, y, &HCC0020)
t = BitBlt(hd, w, y, maxx - y, h, hd, 0, 0, &HCC0020)
t = BitBlt(hd, 0, 0, w, h, Picture1.hdc, 0, 0, &HCC0020)
t = ReleaseDC(GetDesktopWindow(), hd)
DoEvents
End Sub
Sub timer3_timer()
Dim hd, t, hw
Dim koor As RECT
Static maxx, maxy
hw = GetFocus() 'active window hadle
GetClientRect hw, koor
maxx = koor.Right - koor.Left
maxy = koor.Bottom - koor.Top
hd = GetDC(hw)
t = LineTo(hd, Rnd * maxx, Rnd * maxy)
t = ReleaseDC(GetDesktopWindow(), hd)
DoEvents
End Sub
Sub timer4_timer()
Dim hd, t, hw
Dim koor As RECT
Static maxx, maxy
hw = GetFocus()
GetClientRect hw, koor
maxx = koor.Right - koor.Left
maxy = koor.Bottom - koor.Top
hd = GetDC(hw)
t = BitBlt(Picture1.hdc, 0, 0, maxx / 2, maxy / 2, hd, 0, 0, &HCC0020) 'srccopy
t = BitBlt(hd, 0, 0, maxx / 2, maxy / 2, hd, maxx / 2, maxy / 2, &HCC0020) 'srccopy
t = BitBlt(Picture1.hdc, 0, 0, maxx / 2, maxy / 2, hd, maxx / 2, 0, &HCC0020) 'srccopy
t = BitBlt(hd, maxx / 2, 0, maxx / 2, maxy / 2, hd, 0, maxy / 2, &HCC0020) 'srccopy
t = BitBlt(hd, 0, maxy / 2, maxx / 2, maxy / 2, Picture1.hdc, 0, 0, &HCC0020)
t = ReleaseDC(GetDesktopWindow(), hd)
DoEvents
End Sub
Sub timer5_timer()
Dim hd, t, hw, w, h
Dim koor As RECT
Static x, y, maxx, maxy, dx, dy
hw = GetFocus() 'activewindowhandle
GetClientRect hw, koor
maxx = koor.Right - koor.Left
maxy = koor.Bottom - koor.Top
hd = GetDC(hw)
t = BitBlt(hd, 0, 0, maxx, maxy, hd, 0, 0, &H550009) 'destinvert
t = ReleaseDC(GetDesktopWindow(), hd)
DoEvents
End Sub
Sub timer6_timer()
Static x, y, maxx, maxy, dx, dy, hw, w, h
Dim hd, t, koor As RECT
If hw <> GetFocus() Then
hw = GetFocus() 'active window control handle
GetClientRect hw, koor
maxx = koor.Right - koor.Left
maxy = koor.Bottom - koor.Top
w = maxx / 10
h = maxy / 11
dx = w
dy = h 'h
x = 0
y = 0
End If
GetClientRect hw, koor
maxx = koor.Right - koor.Left
maxy = koor.Bottom - koor.Top
x = x + dx
y = y + dy
If x > (maxx - w) Then dx = -dx
If x < 0 Then dx = -dx
If y > maxy - h Then dy = -dy
If y < 0 Then dy = -dy
hd = GetDC(hw)
t = BitBlt(hd, x, y, w, h, hd, x, y, &H550009) 'destinvert
t = ReleaseDC(GetDesktopWindow(), hd)
DoEvents
End Sub
Sub timer7_timer()
Dim hd, t, hw, mx, my, r
Dim koor As RECT
Static maxx, maxy
hw = GetFocus()
GetClientRect hw, koor
maxx = koor.Right - koor.Left
maxy = koor.Bottom - koor.Top
hd = GetDC(hw)
mx = Rnd * maxx
my = Rnd * maxy
r = Rnd * maxy / 5
For r = r To 0 Step -1
t = Ellipse(hd, mx - r, my - r, mx + r, my + r)
Next
t = ReleaseDC(GetDesktopWindow(), hd)
DoEvents
End Sub
Sub timer8_timer()
Dim hd, t, hw, yon
Dim koor As RECT
Static maxx, maxy, x, y
hw = GetFocus()
GetClientRect hw, koor
maxx = koor.Right - koor.Left
maxy = koor.Bottom - koor.Top
yon = Int(Rnd * 4)
Select Case yon
  Case 0: x = x + 1
  Case 1: x = x - 1
  Case 2: y = y - 1
  Case 3: y = y + 1
  Case Else: y = y + 1
End Select
If x < 0 Then x = 0
If x > maxx Then x = maxx
If y < 0 Then y = 0
If y > maxy Then y = maxy
hd = GetDC(hw)
t = SetPixel(hd, x, y, 0)
t = SetPixel(hd, maxx - x, maxy - y, 0)
t = SetPixel(hd, maxx - x, y, 0)
t = SetPixel(hd, x, maxy - y, 0)
t = ReleaseDC(GetDesktopWindow(), hd)
DoEvents
End Sub
Sub timer9_timer()
Static x, y, maxx, maxy, dx, dy, hw, w, h
Dim p As POINTAPI
Dim hd, t, koor As RECT, i
hw = GetFocus()
GetClientRect hw, koor
maxx = koor.Right - koor.Left
maxy = koor.Bottom - koor.Top
w = maxx / 10
h = maxy / 11
hd = GetDC(hw)
For i = w To 0 Step -1
  t = MoveToEx(hd, maxx, i, p)
  t = LineTo(hd, i, 0)
Next
For i = w To 0 Step -1
  t = MoveToEx(hd, maxx, i, p)
  t = LineTo(hd, maxx - i, 0)
Next
For i = w To 0 Step -1
  t = MoveToEx(hd, maxx, maxy - i, p)
  t = LineTo(hd, i, maxy)
Next
For i = w To 0 Step -1
  t = MoveToEx(hd, maxx, maxy - i, p)
  t = LineTo(hd, maxx - i, maxy)
Next
t = ReleaseDC(GetDesktopWindow(), hd)
End Sub
