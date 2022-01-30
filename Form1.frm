VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Duyu - 软件推广"
   ClientHeight    =   6570
   ClientLeft      =   24870
   ClientTop       =   10365
   ClientWidth     =   4275
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":048A
   ScaleHeight     =   6570
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   360
      Top             =   3720
   End
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   360
      Top             =   4200
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   360
      Top             =   4680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   360
      Top             =   5160
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "天行健,君子以自强不息"
      BeginProperty Font 
         Name            =   "华文行楷"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   2775
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "本次开机时间："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0秒"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   5520
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   3960
      Picture         =   "Form1.frx":C55E
      ToolTipText     =   "关闭"
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long

Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const SWP_NOSIZE& = &H1
' 保持窗口大小
Private Const SWP_NOMOVE& = &H2
' 保持窗口位置


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_DblClick()
Image2_Click
End Sub

Private Sub Form_Load()
Dim str(1 To 5) As String
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' 将窗口设为总在最前
Me.Top = Screen.Height
Me.Left = (Screen.Width - Me.Width) * 0.995
If Dir(App.Path & "\DyText.txt") <> "" Then
Open App.Path & "\DyText.txt" For Input As #1
Line Input #1, str(1)
Line Input #1, str(2)
Line Input #1, str(3)
Line Input #1, str(4)
Line Input #1, str(5)
Label3.Caption = str(CInt(1 + 4 * Rnd()))
End If
If Dir("dypic.bmp") <> "" Then
Me.Picture = LoadPicture("dypic.bmp")
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static ox As Integer, oy As Integer
  If Button = 1 Then
    Me.Left = Me.Left + X - ox
    Me.Top = Me.Top + Y - oy
  Else
    ox = X
    oy = Y
  End If
End Sub

Private Sub Image1_Click()
End
End Sub

Private Sub Image2_Click()
Dim es As Integer
On Error Resume Next
es = Shell(App.Path & "\" & "DyEsp.exe", vbNormalFocus)
End Sub

Private Sub Label3_dblClick()

On Error Resume Next
OpenUrl "https://www.baidu.com/s?ie=utf-8&f=8&rsv_bp=1&rsv_idx=1&tn=62095104_19_oem_dg&wd=" & Replace(Label3.Caption, " ", "") & "&fenlei=256&rsv_pq=a5da6ea60001fc91&rsv_t=15f30%2BEA52%2BwMtEOBnCloowP4s5gyRk8Ys3PxgOyCIxTVkWBnr4ALojfJYu0Y1spuEyZaT30obHp&rqlang=cn&rsv_enter=0&rsv_dl=tb&rsv_sug3=1&rsv_sug1=1&rsv_sug7=100&rsv_btype=i"
End Sub
 Private Sub OpenUrl(tUrl As String)
 ShellExecute Me.hwnd, "Open", tUrl, 0, 0, 0
 End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static ox As Integer, oy As Integer
  If Button = 1 Then
    Me.Left = Me.Left + X - ox
    Me.Top = Me.Top + Y - oy
  Else
    ox = X
    oy = Y
  End If
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Timer2.Enabled = False
If CLng(GetTickCount / 1000) > 59 Then
Label1.Caption = Replace(Format(DateAdd("s", CDec(CLng(GetTickCount / 1000)), "00:00"), "mm:ss"), ":", "分") & "秒"
Else
Label1.Caption = CLng(GetTickCount / 1000) & "秒"
End If
If Left(Label1.Caption, 3) = "12分" Then
Label1.Caption = "40秒"
End If
End Sub

Private Sub Timer2_Timer()
Label1.Caption = CStr(Val(Label1.Caption) + 1) & "秒"
If Val(Label1.Caption) = 40 Then
Label1.Caption = "10"
End If
End Sub

Private Sub Timer3_Timer()
End
End Sub

Private Sub Timer4_Timer()
Me.Top = Me.Top - 100
If Me.Top <= (Screen.Height - Me.Height) * 0.9 Then
Timer4.Enabled = False
Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub
