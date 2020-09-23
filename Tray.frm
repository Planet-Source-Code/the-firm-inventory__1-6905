VERSION 5.00
Begin VB.Form Tray 
   Caption         =   "Tray"
   ClientHeight    =   1185
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1185
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3960
      Top             =   240
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   3120
      Picture         =   "Tray.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   2520
      Picture         =   "Tray.frx":030A
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   1320
      Picture         =   "Tray.frx":0614
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   1920
      Picture         =   "Tray.frx":091E
      Top             =   240
      Width           =   480
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu mnuPop 
         Caption         =   "Maximize"
         Index           =   1
      End
      Begin VB.Menu mnuPop 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Exit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim Tray As NOTIFYICONDATA
Private Sub Form_Load()


    Tray.cbSize = Len(Tray)
    Tray.hwnd = Picture1.hwnd
    Tray.uId = 1&
    Tray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    Tray.ucallbackMessage = WM_LBUTTONDOWN
    Tray.hIcon = Image1(1).Picture

    'Create the icon
    Shell_NotifyIcon NIM_ADD, Tray
    Me.Hide
End Sub

Private Sub mnuPop_Click(Index As Integer)
    Select Case Index
        Case 1 'Maximize
        Invent.Visible = True
    
        Case 3  'End
            Tray.cbSize = Len(Tray)
            Tray.hwnd = Picture1.hwnd
            Tray.uId = 1&
            Shell_NotifyIcon NIM_DELETE, Tray
            End
    End Select
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msg = X / Screen.TwipsPerPixelX
    If msg = WM_LBUTTONDBLCLK Then
        mnuPop_Click 0
    ElseIf msg = WM_RBUTTONUP Then
        Me.PopupMenu Menu
    End If
End Sub


Private Sub Timer1_Timer()
    'Animate icon
    Static mPic As Integer
    Me.Icon = Image1(mPic).Picture
    Tray.hIcon = Image1(mPic).Picture
    mPic = mPic + 1
    If mPic = 4 Then mPic = 0
    Shell_NotifyIcon NIM_MODIFY, Tray
End Sub
