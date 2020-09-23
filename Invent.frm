VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Invent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8520
   Icon            =   "Invent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8520
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons(1)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageKey        =   "Back"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Map Network Drive"
            Object.ToolTipText     =   "Map Network Drive"
            ImageKey        =   "Map Network Drive"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Disconnect Net Drive"
            Object.ToolTipText     =   "Disconnect Net Drive"
            ImageKey        =   "Disconnect Net Drive"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Users"
            Object.ToolTipText     =   "Administrator?"
            ImageKey        =   "Users"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "Help"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Arw08up"
            Object.ToolTipText     =   "Stay on top"
            ImageKey        =   "Arw08up"
            Style           =   1
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tray"
            Object.ToolTipText     =   "Tray"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "W95mbx01"
            Object.ToolTipText     =   "Exit"
            ImageKey        =   "W95mbx01"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7800
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   529
      TabCaption(0)   =   "Overview"
      TabPicture(0)   =   "Invent.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " NetworkInfo"
      TabPicture(1)   =   "Invent.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Network"
      Tab(1).Control(1)=   "ImageList4"
      Tab(1).Control(2)=   "Picture2(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   " DirectoryLocation"
      TabPicture(2)   =   "Invent.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView2"
      Tab(2).Control(1)=   "Picture2(2)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Software"
      TabPicture(3)   =   "Invent.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture2(3)"
      Tab(3).Control(1)=   "lsprograms"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   " HardwareInfo"
      TabPicture(4)   =   "Invent.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Hardware"
      Tab(4).Control(1)=   "ImageList3"
      Tab(4).Control(2)=   "Picture2(4)"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "System DSN"
      TabPicture(5)   =   "Invent.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "ListView1"
      Tab(5).Control(1)=   "Picture2(5)"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Network "
      TabPicture(6)   =   "Invent.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "TreeView1"
      Tab(6).Control(1)=   "Picture2(6)"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Config "
      TabPicture(7)   =   "Invent.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "StatusBar1"
      Tab(7).Control(1)=   "Frame1"
      Tab(7).Control(2)=   "Picture2(7)"
      Tab(7).ControlCount=   3
      Begin VB.PictureBox Picture2 
         Height          =   4575
         Index           =   7
         Left            =   -74880
         ScaleHeight     =   4515
         ScaleWidth      =   1875
         TabIndex        =   80
         Top             =   480
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         Height          =   4575
         Index           =   6
         Left            =   -74880
         ScaleHeight     =   4515
         ScaleWidth      =   1875
         TabIndex        =   79
         Top             =   480
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         Height          =   4575
         Index           =   5
         Left            =   -74880
         ScaleHeight     =   4515
         ScaleWidth      =   1875
         TabIndex        =   78
         Top             =   480
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         Height          =   4575
         Index           =   4
         Left            =   -74880
         ScaleHeight     =   4515
         ScaleWidth      =   1875
         TabIndex        =   77
         Top             =   480
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         Height          =   4575
         Index           =   3
         Left            =   -74880
         ScaleHeight     =   4515
         ScaleWidth      =   1875
         TabIndex        =   76
         Top             =   480
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         Height          =   4575
         Index           =   2
         Left            =   -74880
         ScaleHeight     =   4515
         ScaleWidth      =   1875
         TabIndex        =   75
         Top             =   480
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         Height          =   4575
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   4515
         ScaleWidth      =   1875
         TabIndex        =   74
         Top             =   480
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         Height          =   4575
         Index           =   0
         Left            =   120
         Picture         =   "Invent.frx":03EA
         ScaleHeight     =   4515
         ScaleWidth      =   1875
         TabIndex        =   73
         Top             =   480
         Width           =   1935
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4575
         Left            =   -72840
         TabIndex        =   68
         Top             =   480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   8070
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Frame Frame1 
         Height          =   4215
         Left            =   -72840
         TabIndex        =   34
         Top             =   360
         Width           =   6015
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   0
            Left            =   600
            MouseIcon       =   "Invent.frx":1AF94
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":1B29E
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   49
            Top             =   600
            Width           =   495
         End
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   1
            Left            =   1680
            MouseIcon       =   "Invent.frx":1B6E0
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":1B9EA
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   48
            Top             =   600
            Width           =   495
         End
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   2
            Left            =   2760
            MouseIcon       =   "Invent.frx":1BE2C
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":1C136
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   47
            Top             =   600
            Width           =   495
         End
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   3
            Left            =   3840
            MouseIcon       =   "Invent.frx":1C578
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":1C882
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   46
            Top             =   600
            Width           =   495
         End
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   4
            Left            =   4920
            MouseIcon       =   "Invent.frx":1CCC4
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":1CFCE
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   45
            Top             =   600
            Width           =   495
         End
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   5
            Left            =   600
            MouseIcon       =   "Invent.frx":1F770
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":1FA7A
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   44
            Top             =   1800
            Width           =   495
         End
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   6
            Left            =   1680
            MouseIcon       =   "Invent.frx":1FEBC
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":201C6
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   43
            Top             =   1800
            Width           =   495
         End
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   7
            Left            =   2760
            MouseIcon       =   "Invent.frx":20610
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":2091A
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   42
            Top             =   1800
            Width           =   495
         End
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   8
            Left            =   3840
            MouseIcon       =   "Invent.frx":20D5C
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":21066
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   41
            Top             =   1800
            Width           =   495
         End
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   9
            Left            =   4920
            MouseIcon       =   "Invent.frx":21370
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":2167A
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   40
            Top             =   1800
            Width           =   495
         End
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   10
            Left            =   600
            MouseIcon       =   "Invent.frx":21ABC
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":21DC6
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   39
            Top             =   2880
            Width           =   495
         End
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   11
            Left            =   1680
            MouseIcon       =   "Invent.frx":22208
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":22512
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   38
            Top             =   2880
            Width           =   495
         End
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   12
            Left            =   2760
            MouseIcon       =   "Invent.frx":2281C
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":22B26
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   37
            Top             =   2880
            Width           =   495
         End
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   13
            Left            =   3840
            MouseIcon       =   "Invent.frx":22F68
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":23272
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   36
            Top             =   2880
            Width           =   495
         End
         Begin VB.PictureBox Icon 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   14
            Left            =   4920
            MouseIcon       =   "Invent.frx":236B4
            MousePointer    =   99  'Custom
            Picture         =   "Invent.frx":239BE
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   35
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "Add New Hardware"
            Height          =   375
            Index           =   0
            Left            =   360
            MouseIcon       =   "Invent.frx":23E00
            MousePointer    =   99  'Custom
            TabIndex        =   64
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "Add/Remove Programs"
            Height          =   495
            Index           =   1
            Left            =   1440
            MouseIcon       =   "Invent.frx":2410A
            MousePointer    =   99  'Custom
            TabIndex        =   63
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "Date/Time"
            Height          =   255
            Index           =   2
            Left            =   2520
            MouseIcon       =   "Invent.frx":24414
            MousePointer    =   99  'Custom
            TabIndex        =   62
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "Display"
            Height          =   255
            Index           =   3
            Left            =   3600
            MouseIcon       =   "Invent.frx":2471E
            MousePointer    =   99  'Custom
            TabIndex        =   61
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "Internet"
            Height          =   255
            Index           =   4
            Left            =   4680
            MouseIcon       =   "Invent.frx":24A28
            MousePointer    =   99  'Custom
            TabIndex        =   60
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "Joystick"
            Height          =   255
            Index           =   5
            Left            =   360
            MouseIcon       =   "Invent.frx":24D32
            MousePointer    =   99  'Custom
            TabIndex        =   59
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "Keyboard"
            Height          =   255
            Index           =   6
            Left            =   1440
            MouseIcon       =   "Invent.frx":2503C
            MousePointer    =   99  'Custom
            TabIndex        =   58
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "Modems"
            Height          =   255
            Index           =   7
            Left            =   2520
            MouseIcon       =   "Invent.frx":25346
            MousePointer    =   99  'Custom
            TabIndex        =   57
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "Mouse"
            Height          =   255
            Index           =   8
            Left            =   3600
            MouseIcon       =   "Invent.frx":25650
            MousePointer    =   99  'Custom
            TabIndex        =   56
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "Multimedia"
            Height          =   255
            Index           =   9
            Left            =   4680
            MouseIcon       =   "Invent.frx":2595A
            MousePointer    =   99  'Custom
            TabIndex        =   55
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "Network"
            Height          =   255
            Index           =   10
            Left            =   360
            MouseIcon       =   "Invent.frx":25C64
            MousePointer    =   99  'Custom
            TabIndex        =   54
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "Passwords"
            Height          =   255
            Index           =   11
            Left            =   1440
            MouseIcon       =   "Invent.frx":25F6E
            MousePointer    =   99  'Custom
            TabIndex        =   53
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "Regional Settings"
            Height          =   495
            Index           =   12
            Left            =   2520
            MouseIcon       =   "Invent.frx":26278
            MousePointer    =   99  'Custom
            TabIndex        =   52
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "Sounds"
            Height          =   255
            Index           =   13
            Left            =   3600
            MouseIcon       =   "Invent.frx":26582
            MousePointer    =   99  'Custom
            TabIndex        =   51
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label Iconname 
            Alignment       =   2  'Center
            Caption         =   "System"
            Height          =   255
            Index           =   14
            Left            =   4680
            MouseIcon       =   "Invent.frx":2688C
            MousePointer    =   99  'Custom
            TabIndex        =   50
            Top             =   3360
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4575
         Left            =   -72840
         TabIndex        =   25
         Top             =   480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   8070
         View            =   3
         Arrange         =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Name"
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Key             =   "Driver"
            Text            =   "Driver"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.Frame Frame2 
         Height          =   4695
         Left            =   2160
         TabIndex        =   6
         Top             =   360
         Width           =   6015
         Begin VB.Timer Timer1 
            Left            =   5520
            Top             =   3120
         End
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000009&
            Height          =   135
            Left            =   120
            ScaleHeight     =   75
            ScaleWidth      =   5715
            TabIndex        =   26
            Top             =   4080
            Width           =   5775
            Begin VB.PictureBox Picture6 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   135
               Left            =   0
               ScaleHeight     =   135
               ScaleWidth      =   5715
               TabIndex        =   27
               Top             =   0
               Width           =   5715
            End
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            X1              =   120
            X2              =   5880
            Y1              =   3720
            Y2              =   3720
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   11
            Left            =   1440
            TabIndex        =   71
            Top             =   4320
            Width           =   4455
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   70
            Top             =   2280
            Width           =   1155
         End
         Begin VB.Label Label7 
            Caption         =   "Label7"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   840
            TabIndex        =   69
            Top             =   4320
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   10
            Left            =   1680
            TabIndex        =   66
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   9
            Left            =   1680
            TabIndex        =   33
            Top             =   2760
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   32
            Top             =   2760
            Width           =   1755
         End
         Begin VB.Label Label6 
            Caption         =   "Label6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   4320
            Width           =   2535
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Label5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   30
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Label4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   29
            Top             =   3840
            Width           =   2175
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   3840
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   23
            Top             =   600
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   22
            Top             =   2520
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   21
            Top             =   1080
            Width           =   1635
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   20
            Top             =   1320
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   19
            Top             =   1560
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   18
            Top             =   1800
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   7
            Left            =   3480
            TabIndex        =   17
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   16
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   15
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   14
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   3
            Left            =   1920
            TabIndex        =   13
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   4
            Left            =   1920
            TabIndex        =   12
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   5
            Left            =   1920
            TabIndex        =   11
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   6
            Left            =   1920
            TabIndex        =   10
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   7
            Left            =   4440
            TabIndex        =   9
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   8
            Left            =   3480
            TabIndex        =   8
            Top             =   600
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   8
            Left            =   4440
            TabIndex        =   7
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture5 
         Height          =   4575
         Index           =   0
         Left            =   -74880
         ScaleHeight     =   4515
         ScaleWidth      =   1875
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.PictureBox Picture1 
         Height          =   4575
         Index           =   0
         Left            =   -74880
         ScaleHeight     =   4515
         ScaleWidth      =   1875
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin MSComctlLib.ImageList ImageList4 
         Left            =   -72480
         Top             =   4080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":26B96
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":26EB2
               Key             =   "NetworkPrinters"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":271CC
               Key             =   "Network"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":29980
               Key             =   "AdapterName"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":29C9C
               Key             =   "MacAddress"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":29FB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2A40C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2A860
               Key             =   "WINS"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2ACB4
               Key             =   "DNS"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2B108
               Key             =   "Default Gateway"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2B55C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2B9B0
               Key             =   "Subnetmask"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2BE04
               Key             =   "NetworkComment"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2C258
               Key             =   "IPAddress"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2C574
               Key             =   "LocalHostname"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   -72720
         Top             =   4320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2C9C8
               Key             =   "Keyboard"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2CE1C
               Key             =   "cdrom"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2D270
               Key             =   "Diskettestation"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2D6C4
               Key             =   "Disks"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2DB18
               Key             =   "Networkdrive"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2DF6C
               Key             =   "Networkprinters"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2E288
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2E5A4
               Key             =   "hardware"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2E8C0
               Key             =   "Soundcard"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2EBDC
               Key             =   "Graphic card"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2EEF8
               Key             =   "Modem"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2F214
               Key             =   "Printers"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2F530
               Key             =   "Mouse"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Invent.frx":2F984
               Key             =   "Monitor"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView Hardware 
         Height          =   4575
         Left            =   -72840
         TabIndex        =   2
         Top             =   480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   8070
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin VB.ListBox lsprograms 
         Height          =   4545
         Left            =   -72840
         TabIndex        =   1
         Top             =   480
         Width           =   6015
      End
      Begin MSComctlLib.TreeView Network 
         Height          =   4575
         Left            =   -72840
         TabIndex        =   3
         Top             =   480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   8070
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   -72840
         TabIndex        =   65
         Top             =   4680
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   4
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               Alignment       =   1
               AutoSize        =   1
               Object.Width           =   2699
               TextSave        =   "16:05"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               Alignment       =   1
               AutoSize        =   1
               Object.Width           =   2699
               TextSave        =   "31/03/2000"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   1
               Alignment       =   1
               Enabled         =   0   'False
               TextSave        =   "CAPS"
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   2
               Alignment       =   1
               TextSave        =   "NUM"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4575
         Left            =   -72840
         TabIndex        =   67
         Top             =   480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   8070
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Name"
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Location"
            Text            =   "Location"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   0
      Left            =   3540
      Top             =   2700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":2FDD8
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":2FEEA
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":2FFFC
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":3010E
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":30220
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":30332
            Key             =   "Map Network Drive"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":30444
            Key             =   "Disconnect Net Drive"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":30556
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":30668
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   1
      Left            =   3660
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":3077A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":3088C
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":3099E
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":30AB0
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":30BC2
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":30CD4
            Key             =   "Map Network Drive"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":30DE6
            Key             =   "Disconnect Net Drive"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":30EF8
            Key             =   "Users"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":31212
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":31324
            Key             =   "Arw08up"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Invent.frx":3163E
            Key             =   "W95mbx01"
         EndProperty
      EndProperty
   End
   Begin VB.Menu FIle 
      Caption         =   "&File"
      Begin VB.Menu save 
         Caption         =   "&Save"
      End
      Begin VB.Menu Streepken 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu sys 
      Caption         =   "&System"
      Begin VB.Menu system 
         Caption         =   "Logoff"
         Index           =   1
      End
      Begin VB.Menu system 
         Caption         =   "Reboot"
         Index           =   2
      End
      Begin VB.Menu system 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu system 
         Caption         =   "Shutdown"
         Index           =   4
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu helps 
         Caption         =   "&Help"
      End
      Begin VB.Menu Streep 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Invent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Z As Integer
Dim X As Integer
Dim D As Integer
Dim s As Integer

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Tray.cbSize = Len(Tray)
            Tray.hwnd = Picture1.hwnd
            Tray.uId = 1&
            Shell_NotifyIcon NIM_DELETE, Tray
    
End
End Sub

Private Sub system_Click(Index As Integer)
    Select Case Index
        Case 1 'Logoff
            Call LogOff
        Case 2  'Reboot
            Call Reboots
        Case 4 ' Shutdown
            Call ShutDown
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Open"

        Case "Save"

        Case "Forward"
            SSTab1.Tab = SSTab1.Tab + 1
        Case "Back"
            SSTab1.Tab = SSTab1.Tab - 1
        Case "Find"

        Case "Map Network Drive"

        Case "Disconnect Net Drive"

        Case "Users"
            If IsAdmins Then
                MsgBox "Yes, You have Administrator rights!", vbInformation, Caption
            Else
                MsgBox "No, You have no Administrator rights!", vbCritical, Caption
            End If
        Case "Help"
            Call FileExecutor(Me.hwnd, App.Path & "\contact_me.exe", "Open")
        Case "Tray"
        Me.Visible = False
        Mapper.Visible = False
        Load Tray

        Case "Arw08up"
            If Button.value = tbrPressed Then
                Call Form_Ontop(Me.hwnd, True)
            Else
                Call Form_Ontop(Me.hwnd, False)
            End If
        Case "W95mbx01"
            End
    End Select
End Sub

Private Sub exit_Click()
    End
End Sub

Private Sub DirectNodes_Collapse(ByVal Node As MSComCtlLib.Node)
    If Node.Key <> "DirectoryLocations" Then
        Node.Image = "maptoe"
    End If
End Sub

Private Sub DirectNodes_Expand(ByVal Node As MSComCtlLib.Node)
    If Node.Key <> "DirectoryLocations" Then
        Node.Image = "mapopen"
    End If
End Sub

Private Sub Form_Load()

    SSTab1.Tab = 0
    DonNetworkNodes
    DoEvents
    DoHardWareNodes
    DoEvents
    DoDirectNodes
    DoEvents
    Programs
    DoEvents
    Drives
    DoEvents
    Sysinfo
    DoEvents
    DoDSN
    DoEvents
    GetLanguage
    DoEvents
    ProfileLoadWinIniList
    DoEvents
    Call Dodrives("c:\")
    For I = 1 To 7
        Picture2(I).Picture = Picture2(0).Picture
    Next I

End Sub


Private Sub DoDirectNodes()
    Dim Dnode As Node
    Dim vNode As Node
    Dim C$
    Dim Dir(0 To 30)
    Dim DIRL(0 To 30)

    Dir(0) = "Application Data": DIRL(0) = FolderLocation(CSIDL_APPDATA, Me.hwnd)
    Dir(1) = "All Users Desktop": DIRL(1) = FolderLocation _
            (CSIDL_COMMON_DESKTOPDIRECTORY, Me.hwnd)
    Dir(2) = "Desktop": DIRL(2) = FolderLocation(CSIDL_DESKTOP, Me.hwnd)
    Dir(3) = "DesktopDirectory": DIRL(3) = FolderLocation _
            (CSIDL_DESKTOPDIRECTORY, Me.hwnd)
    Dir(4) = "Cookies": DIRL(4) = FolderLocation _
            (CSIDL_COOKIES, Me.hwnd)
    Dir(5) = "Favorites": DIRL(5) = FolderLocation(CSIDL_FAVORITES, Me.hwnd)
    Dir(6) = "Fonts": DIRL(6) = FolderLocation(CSIDL_FONTS, Me.hwnd)
    Dir(7) = "History": DIRL(7) = FolderLocation(CSIDL_HISTORY, Me.hwnd)
    Dir(8) = "Inernet_Cache": DIRL(8) = FolderLocation(CSIDL_INTERNET_CACHE, Me.hwnd)
    Dir(9) = "Nethood": DIRL(9) = FolderLocation(CSIDL_NETHOOD, Me.hwnd)
    Dir(10) = "Personal": DIRL(10) = FolderLocation(CSIDL_PERSONAL, Me.hwnd)
    Dir(11) = "Printer": DIRL(11) = FolderLocation(CSIDL_PRINTERS, Me.hwnd) ' ?
    Dir(12) = "PrinterHood": DIRL(12) = FolderLocation(CSIDL_PRINTHOOD, Me.hwnd)
    Dir(13) = "Programs": DIRL(13) = FolderLocation(CSIDL_PROGRAMS, Me.hwnd)
    Dir(14) = "Recent": DIRL(14) = FolderLocation(CSIDL_RECENT, Me.hwnd)
    Dir(15) = "Sendto": DIRL(15) = FolderLocation(CSIDL_SENDTO, Me.hwnd)
    Dir(16) = "Startup": DIRL(16) = FolderLocation(CSIDL_STARTUP, Me.hwnd)
    Dir(17) = "Templates": DIRL(17) = FolderLocation(CSIDL_TEMPLATES, Me.hwnd)
    Dir(18) = "AltStartUp": DIRL(18) = FolderLocation(CSIDL_ALTSTARTUP, Me.hwnd)
    Dir(19) = "BitBucket": DIRL(19) = FolderLocation(CSIDL_BITBUCKET, Me.hwnd)
    Dir(20) = "CommonAltStartup": DIRL(20) = FolderLocation _
            (CSIDL_COMMON_ALTSTARTUP, Me.hwnd)
    Dir(21) = "CommonFavorites": DIRL(21) = FolderLocation _
            (CSIDL_COMMON_FAVORITES, Me.hwnd)
    Dir(22) = "CommonPrograms": DIRL(22) = FolderLocation _
            (CSIDL_COMMON_PROGRAMS, Me.hwnd)
    Dir(23) = "CommonStartMenu": DIRL(23) = FolderLocation _
            (CSIDL_COMMON_STARTMENU, Me.hwnd)
    Dir(24) = "CommonStartUp": DIRL(24) = FolderLocation _
            (CSIDL_COMMON_STARTUP, Me.hwnd)
    Dir(25) = "Drives": DIRL(25) = FolderLocation _
            (CSIDL_DRIVES, Me.hwnd)
    Dir(26) = "Internet": DIRL(26) = FolderLocation _
            (CSIDL_INTERNET, Me.hwnd)
    Dir(27) = "Network": DIRL(27) = FolderLocation _
            (CSIDL_NETWORK, Me.hwnd)
    Dir(28) = "MenuStart": DIRL(28) = FolderLocation _
            (CSIDL_STARTMENU, Me.hwnd)

    Dir(29) = "System Dir": DIRL(29) = SystemDir
    Dir(30) = "Temp Dir": DIRL(30) = TempDir

    Dim TMP As ListItem

    ListView2.ColumnHeaders(2).Width = (ListView2.Width - ListView2.ColumnHeaders(1).Width) - 90


    For I = 1 To 30
        Set TMP = ListView2.ListItems.Add(, , Dir(I))
        TMP.SubItems(1) = DIRL(I)

    Next I

End Sub
Private Sub DoHardWareNodes()
    Dim Hnode As Node
    Dim C$

    Dim systeem As New system
    Hardware.Nodes.Clear
    Hardware.ImageList = ImageList3

    C = "Hardware"
    Set Hnode = Hardware.Nodes.Add(, , C, "Hardware", "hardware")
    Hnode.Expanded = True

    C = "Printers"
    Set Hnode = Hardware.Nodes.Add("Hardware", 4, C, C, C)
    H = Countit(HardInfo(Printer))

    For I = 0 To H
        v = Trim(HardInfo(Printer), I + 1)
        If v = "" Then Exit For
        Set Hnode = Hardware.Nodes.Add(C, 4, , v, C)
    Next I

    H = Countit(ProfileLoadWinIniList)
    For I = 0 To H
        v = Trim(ProfileLoadWinIniList, I + 1)
        If v = "" Then Exit For
        Set Hnode = Hardware.Nodes.Add(C, 4, , v, C)
    Next I

    C = "Modem"
    Set Hnode = Hardware.Nodes.Add("Hardware", 4, C, "Modems", "Modem")
    H = Countit(HardInfo(Modem))
    For I = 0 To H
        v = Trim(HardInfo(Modem), I + 1)
        If v = "" Then Exit For
        Set Hnode = Hardware.Nodes.Add(C, 4, , v, C)
    Next I


    C = "Soundcard"
    Set Hnode = Hardware.Nodes.Add("Hardware", 4, C, "Souncard", "Soundcard")
    H = Countit(HardInfo(Soundcard))
    For I = 0 To H
        v = Trim(HardInfo(Soundcard), I + 1)
        If v = "" Then Exit For
        Set Hnode = Hardware.Nodes.Add(C, 4, , v, C)
    Next I

    C = "Graphic card"
    Set Hnode = Hardware.Nodes.Add("Hardware", 4, C, "Graphic card", "Graphic card")
    H = Countit(HardInfo(GraphicCard))
    For I = 0 To H
        v = Trim(HardInfo(GraphicCard), I + 1)
        If v = "" Then Exit For
        Set Hnode = Hardware.Nodes.Add(C, 4, , v, C)
    Next I

    C = "Mouse"
    Set Hnode = Hardware.Nodes.Add("Hardware", 4, C, "Mouse", "Mouse")
    H = Countit(HardInfo(Mouse))
    For I = 0 To H
        v = Trim(HardInfo(Mouse), I + 1)
        If v = "" Then Exit For
        Set Hnode = Hardware.Nodes.Add(C, 4, , v, C)
    Next I

    C = "Keyboard"
    Set Hnode = Hardware.Nodes.Add("Hardware", 4, C, C, C)
    H = Countit(systeem.KeyboardType)
    For I = 0 To H
        v = Trim(systeem.KeyboardType, I + 1)
        If v = "" Then Exit For
        Set Hnode = Hardware.Nodes.Add(C, 4, , v, C)
    Next I

    C = "Monitor"
    Set Hnode = Hardware.Nodes.Add("Hardware", 4, C, C, C)
    H = Countit(HardInfo(Monitor))
    For I = 0 To H
        v = Trim(HardInfo(Monitor), I + 1)
        If v = "" Then Exit For
        Set Hnode = Hardware.Nodes.Add(C, 4, , v, C)
    Next I
End Sub
Private Sub DonNetworkNodes()
    Dim Hnode As Node
    Dim C$
    Dim I As Integer
    Network.Nodes.Clear
    Network.ImageList = ImageList4

    C = "Network"
    Set Hnode = Network.Nodes.Add(, , C, "Network", C)
    Hnode.Expanded = True

    C = "LocalHostname"
    Set Hnode = Network.Nodes.Add("Network", 4, C, C, C)
    Set Hnode = Network.Nodes.Add(C, 4, , Winsock1.LocalHostName, C)

    C = "IPAddress"
    Set Hnode = Network.Nodes.Add("Network", 4, C, C, C)
    Set Hnode = Network.Nodes.Add(C, 4, , Winsock1.LocalIP, C)

    C = "Subnetmask"
    Set Hnode = Network.Nodes.Add("Network", 4, C, C, C)
    H = Countit(NetInfo(SubNetmask))
    For I = 0 To H
        v = Trim(NetInfo(SubNetmask), I + 1)
        If v = "" Then Exit For
        Set Hnode = Network.Nodes.Add(C, 4, , v, C)
    Next I

    C = "Default Gateway"
    Set Hnode = Network.Nodes.Add("Network", 4, C, C, C)
    H = Countit(NetInfo(gateway))
    For I = 0 To H
        v = Trim(NetInfo(gateway), I + 1)
        If v = "" Then Exit For
        Set Hnode = Network.Nodes.Add(C, 4, , v, C)
    Next I


    C = "DNS"
    Set Hnode = Network.Nodes.Add("Network", 4, C, C, C)

    H = Countit(NetInfo(DNS))
    For I = 0 To H
        v = Trim(NetInfo(DNS), I + 1)
        If v = "" Then Exit For
        Set Hnode = Network.Nodes.Add(C, 4, , v, C)
    Next I

    C = "MacAddress"
    Set Hnode = Network.Nodes.Add("Network", 4, C, C, C)

    H = Countit(NetInfo(MacAdrress))
    For I = 0 To H
        v = Trim(NetInfo(MacAdrress), I + 1)
        If v = "" Then Exit For
        Set Hnode = Network.Nodes.Add(C, 4, , v, C)
    Next I

    C = "WINS"
    Set Hnode = Network.Nodes.Add("Network", 4, C, C, C)

    H = Countit(NetInfo(wins))
    For I = 0 To H
        v = Trim(NetInfo(wins), I + 1)
        If v = "" Then Exit For
        Set Hnode = Network.Nodes.Add(C, 4, , v, C)
    Next I

    C = "AdapterName"
    Set Hnode = Network.Nodes.Add("Network", 4, C, C, C)

    H = Countit(NetInfo(AdapterName))
    For I = 0 To H
        v = Trim(NetInfo(AdapterName), I + 1)
        If v = "" Then Exit For
        Set Hnode = Network.Nodes.Add(C, 4, , v, C)
    Next I

    C = "NetworkComment"
    Set Hnode = Network.Nodes.Add("Network", 4, C, C, C)

    H = Countit(NetInfo(networkcomment))
    For I = 0 To H
        v = Trim(NetInfo(networkcomment), I + 1)
        If v = "" Then Exit For
        Set Hnode = Network.Nodes.Add(C, 4, , v, C)
    Next I





    Z = Empty
    X = Empty
    C = Empty
    D = Empty



End Sub
Private Function Trim(p_strTmp As String, I As Integer) As String
    Trim = ""
    If Z = Empty Then Z = 1
    For X = Z To Len(p_strTmp)
        A = Mid(p_strTmp, X, 1)
        If A = "|" Or X = Len(p_strTmp) Then
            C = C + 1
            If C = I Then
                If X = Len(p_strTmp) Then
                    If C = 1 Then
                        Trim = Mid(p_strTmp, Z, D + 1)
                        D = Empty
                        Exit For
                    Else
                        s = s + D
                        Trim = Mid(p_strTmp, s - D + 2, D)
                        H = Trim
                        s = Empty

                        Exit For
                    End If
                Else
                    If I = 1 Then

                        Trim = Mid(p_strTmp, s + 1, D)
                        If Mid(p_strTmp, D, 1) = "|" Then
                            D = D - 1
                            Trim = Mid(p_strTmp, s + 1, D)
                        End If
                        H = Trim
                        s = s + D
                        D = Empty
                    Else
                        Trim = Mid(p_strTmp, s + 2, D - 1)
                        H = Trim
                        s = s + D
                        D = Empty
                    End If
                    D = Empty
                    Exit For
                End If
            Else
                D = Empty
            End If
        End If
        D = D + 1
    Next X
End Function
Private Function Countit(p_strTmp As String) As Integer
    Debug.Print p_strTmp
    p_lngPos = InStr(1, p_strTmp, "|", vbTextCompare)
    If p_strTmp = "" Then
        Countit = 0
        Exit Function
    End If
    p_blnFirstTime = True
    Countit = 0
    Do While p_lngPos > 0
        If p_blnFirstTime = True Then
            p_strTmp = Mid$(p_strTmp, p_lngPos + 1)
            Countit = Countit + 2
            p_blnFirstTime = False
        Else
            p_strTmp = Mid$(p_strTmp, p_lngPos + 1)
            Countit = Countit + 1
            p_blnFirstTime = True
        End If
        p_lngPos = InStr(1, p_strTmp, "|", vbTextCompare)
    Loop
End Function
Public Sub ControlPanels(Filename As String)
    Dim rtn As Double
    On Error Resume Next
    rtn = Shell(Filename, 5)
End Sub
Private Sub Icon_Click(Index As Integer)

    If Index = 0 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1")
    ElseIf Index = 1 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1")
    ElseIf Index = 2 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
    ElseIf Index = 3 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0")
    ElseIf Index = 4 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0")
    ElseIf Index = 5 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL joy.cpl")
    ElseIf Index = 6 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1")
    ElseIf Index = 7 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL modem.cpl")
    ElseIf Index = 8 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL main.cpl @0")
    ElseIf Index = 9 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0")
    ElseIf Index = 10 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl")
    ElseIf Index = 11 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL password.cpl")
    ElseIf Index = 12 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0")
    ElseIf Index = 13 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1")
    ElseIf Index = 14 Then
        Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0")
    End If

End Sub

Private Sub DoDSN()
    Dim TMP As ListItem

    ListView1.ColumnHeaders(2).Width = (ListView1.Width - ListView1.ColumnHeaders(1).Width) - 90
    H = Countit(dSN)
    Z = 1
    For I = 0 To H
        v = Trim(dSN, I + 1)
        If v = "" Then Exit For
        If Z = 1 Then
            Set TMP = ListView1.ListItems.Add(, , v)
            Z = Z + 1
        Else
            TMP.SubItems(1) = v
            Z = 1
        End If
    Next I
    Z = Empty: X = Empty: C = Empty: D = Empty
End Sub


