VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTerminal 
   Caption         =   "Modem Master Ver. 1.6     Terminal And Modem Diagnostics"
   ClientHeight    =   4935
   ClientLeft      =   2325
   ClientTop       =   2235
   ClientWidth     =   7155
   ForeColor       =   &H00000000&
   Icon            =   "vbterm.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   7155
   Begin ComctlLib.Toolbar tbrToolBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   19
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "OpenLogFile"
            Description     =   "Open Log File..."
            Object.ToolTipText     =   "Open Log File..."
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "CloseLogFile"
            Description     =   "Close Log File"
            Object.ToolTipText     =   "Close Log File"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "DialPhoneNumber"
            Description     =   "Dial Phone Number..."
            Object.ToolTipText     =   "Dial Phone Number..."
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "HangUpPhone"
            Description     =   "Hang Up Phone"
            Object.ToolTipText     =   "Hang Up Phone"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Properties"
            Description     =   "Properties..."
            Object.ToolTipText     =   "Properties..."
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "TransmitTextFile"
            Description     =   "Transmit Text File..."
            Object.ToolTipText     =   "Transmit Text File..."
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "A"
            Object.ToolTipText     =   "Open Port"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "B"
            Object.ToolTipText     =   "Dial Number"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "C"
            Object.ToolTipText     =   "Answer Call"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "D"
            Object.ToolTipText     =   "Current Moden Settings"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "E"
            Object.ToolTipText     =   "Originate Mode"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "F"
            Object.ToolTipText     =   "V.42BIS / MNPS"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "G"
            Object.ToolTipText     =   "Reset Factory Settings"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "H"
            Object.ToolTipText     =   "S Registers Available"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "I"
            Object.ToolTipText     =   "AT Commands Available"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "J"
            Object.ToolTipText     =   "Ampersand Available"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "K"
            Object.ToolTipText     =   "Print Terminal Image"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "L"
            Object.ToolTipText     =   "Set Terminal Fonts"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "M"
            Object.ToolTipText     =   "Display NVRAM Contents"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   240
         Left            =   4000
         TabIndex        =   2
         Top             =   75
         Width           =   240
         Begin VB.Image imgConnected 
            Height          =   240
            Left            =   0
            Picture         =   "vbterm.frx":030A
            Stretch         =   -1  'True
            ToolTipText     =   "Toggles Port"
            Top             =   0
            Width           =   240
         End
         Begin VB.Image imgNotConnected 
            Height          =   240
            Left            =   0
            Picture         =   "vbterm.frx":0454
            Stretch         =   -1  'True
            ToolTipText     =   "Toggles Port"
            Top             =   0
            Width           =   240
         End
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   210
      Top             =   3645
   End
   Begin VB.TextBox txtTerm 
      Height          =   3930
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   600
      Width           =   6135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   165
      Top             =   1815
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   45
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      RThreshold      =   1
      RTSEnable       =   -1  'True
      SThreshold      =   1
      InputMode       =   1
   End
   Begin MSComDlg.CommonDialog OpenLog 
      Left            =   105
      Top             =   1170
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "Open Communications Log File"
      Filter          =   "*.*"
      FilterIndex     =   501
      FontSize        =   9.02458e-38
   End
   Begin ComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4620
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Status:"
            TextSave        =   "Status:"
            Key             =   "Status"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Communications Port Status"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8310
            MinWidth        =   2
            Text            =   "Settings:"
            TextSave        =   "Settings:"
            Key             =   "Settings"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Communications Port Settings"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   1244
            MinWidth        =   1244
            TextSave        =   ""
            Key             =   "ConnectTime"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Connect Time"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   165
      Top             =   2445
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   19
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":059E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":08B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":0BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":0EEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":1206
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":1520
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":183A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":1B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":1E6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":2188
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":24A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":27BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":2AD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":2DF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":310A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":3424
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":373E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":3A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "vbterm.frx":3D72
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenLog 
         Caption         =   "&Open Log File..."
      End
      Begin VB.Menu mnuCloseLog 
         Caption         =   "&Close Log File"
         Enabled         =   0   'False
      End
      Begin VB.Menu m3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendText 
         Caption         =   "&Transmit Text File..."
         Enabled         =   0   'False
      End
      Begin VB.Menu Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu fin 
         Caption         =   "&New"
      End
      Begin VB.Menu fio 
         Caption         =   "&Open"
      End
      Begin VB.Menu fis 
         Caption         =   "&Save"
      End
      Begin VB.Menu fisa 
         Caption         =   "&Save As"
      End
      Begin VB.Menu prt1 
         Caption         =   "&Print Specs"
      End
      Begin VB.Menu BARA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile5"
         Index           =   5
         Visible         =   0   'False
      End
   End
   Begin VB.Menu e0 
      Caption         =   "&Edit"
      Begin VB.Menu e1 
         Caption         =   "&Cut"
      End
      Begin VB.Menu e2 
         Caption         =   "&Copy"
      End
      Begin VB.Menu e3 
         Caption         =   "&Paste"
      End
      Begin VB.Menu e4 
         Caption         =   "&Select All"
      End
   End
   Begin VB.Menu op0 
      Caption         =   "&Opts"
      Begin VB.Menu op1 
         Caption         =   "&Foreground Color"
      End
      Begin VB.Menu op2 
         Caption         =   "&Background Color"
      End
      Begin VB.Menu opt 
         Caption         =   "&Select Font"
      End
   End
   Begin VB.Menu mnuPort 
      Caption         =   "&Port"
      Begin VB.Menu mnuOpen 
         Caption         =   "Port &Open"
      End
      Begin VB.Menu MBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties..."
      End
      Begin VB.Menu mbar2 
         Caption         =   "-"
      End
      Begin VB.Menu sldb 
         Caption         =   "&Recieved Signal DBm"
      End
      Begin VB.Menu rls 
         Caption         =   "&Report Line Signal"
      End
      Begin VB.Menu rte 
         Caption         =   "&Report DCE-DTE-Comp."
      End
      Begin VB.Menu mbar3 
         Caption         =   "-"
      End
      Begin VB.Menu ap 
         Caption         =   "&Display Active Profile"
      End
      Begin VB.Menu sp 
         Caption         =   "&Display Stored Profile"
      End
      Begin VB.Menu sn 
         Caption         =   "&Display Stored Numbers"
      End
      Begin VB.Menu dcs 
         Caption         =   "&Display Control Settings"
      End
      Begin VB.Menu mbar4 
         Caption         =   "-"
      End
      Begin VB.Menu lp0 
         Caption         =   "&Load Profile 0"
      End
      Begin VB.Menu lp1 
         Caption         =   "&Load Profile 1"
      End
   End
   Begin VB.Menu mnuMSComm 
      Caption         =   "&Com"
      Begin VB.Menu mnuInputLen 
         Caption         =   "&InputLen..."
      End
      Begin VB.Menu mnuRThreshold 
         Caption         =   "&RThreshold..."
      End
      Begin VB.Menu mnuSThreshold 
         Caption         =   "&SThreshold..."
      End
      Begin VB.Menu mnuParRep 
         Caption         =   "P&arityReplace..."
      End
      Begin VB.Menu mnuDTREnable 
         Caption         =   "&DTREnable"
      End
      Begin VB.Menu BAR3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHCD 
         Caption         =   "&CDHolding..."
      End
      Begin VB.Menu mnuHCTS 
         Caption         =   "CTSH&olding..."
      End
      Begin VB.Menu mnuHDSR 
         Caption         =   "DSRHo&lding..."
      End
   End
   Begin VB.Menu mnuCall 
      Caption         =   "C&all"
      Begin VB.Menu mnuDial 
         Caption         =   "&Dial Phone Number..."
      End
      Begin VB.Menu mnuHangUp 
         Caption         =   "&Hang Up Phone"
         Enabled         =   0   'False
      End
      Begin VB.Menu ai1 
         Caption         =   "&Answer Incoming Call"
      End
   End
   Begin VB.Menu s0 
      Caption         =   "&Info"
      Begin VB.Menu rfc 
         Caption         =   "&Report Fax Classes"
      End
      Begin VB.Menu s1 
         Caption         =   "&S Registers Available"
      End
      Begin VB.Menu s2 
         Caption         =   "A&T Command List "
      End
      Begin VB.Menu s3 
         Caption         =   "&Ampersand Commands"
      End
      Begin VB.Menu s4 
         Caption         =   "&Product Code"
      End
      Begin VB.Menu s5 
         Caption         =   "&Checksum Info."
      End
      Begin VB.Menu s6 
         Caption         =   "P&roduct I.D."
      End
      Begin VB.Menu s12 
         Caption         =   "&Current Profiles"
      End
      Begin VB.Menu s7 
         Caption         =   "C&urrent Settings"
      End
      Begin VB.Menu s8 
         Caption         =   "&NVRAM Settings"
      End
      Begin VB.Menu s9 
         Caption         =   "Pro&duct Config"
      End
      Begin VB.Menu s10 
         Caption         =   "P&NP Enumeration "
      End
      Begin VB.Menu s11 
         Caption         =   "Print Terminal Image"
      End
   End
   Begin VB.Menu ss0 
      Caption         =   "&Setting"
      Begin VB.Menu mp1 
         Caption         =   "&Modem Properties"
      End
      Begin VB.Menu ss1 
         Caption         =   "Reset Factory Settings"
      End
      Begin VB.Menu ss2 
         Caption         =   "&Reset Modem (ATZ)"
      End
      Begin VB.Menu ss9 
         Caption         =   "&Hardware Flow Reset"
      End
      Begin VB.Menu ss10 
         Caption         =   "&Software Flow Reset"
      End
      Begin VB.Menu ss3 
         Caption         =   "&Extended Result Codes"
      End
      Begin VB.Menu ss4 
         Caption         =   "Speaker Set &Low"
      End
      Begin VB.Menu ss5 
         Caption         =   "Speaker Set &Med"
      End
      Begin VB.Menu ss6 
         Caption         =   "Speaker Set &High"
      End
      Begin VB.Menu ss7 
         Caption         =   "Data Retrain ON"
      End
      Begin VB.Menu ss8 
         Caption         =   "Data Retrain OFF"
      End
   End
   Begin VB.Menu mo 
      Caption         =   "&Modes"
      Begin VB.Menu m1 
         Caption         =   "Answer    Mode"
      End
      Begin VB.Menu m2 
         Caption         =   "Originate Mode"
      End
      Begin VB.Menu m33 
         Caption         =   "Pulse Dial Mode"
      End
      Begin VB.Menu m4 
         Caption         =   "Tone Dial Mode (D)"
      End
      Begin VB.Menu m5 
         Caption         =   "&Word Result Codes (D)"
      End
      Begin VB.Menu m6 
         Caption         =   "&Numeric Result Codes"
      End
      Begin VB.Menu m7 
         Caption         =   "Enable  ARQ Codes"
      End
      Begin VB.Menu m8 
         Caption         =   "Disable ARQ Codes"
      End
      Begin VB.Menu m9 
         Caption         =   "DCD ON Only"
      End
      Begin VB.Menu m10 
         Caption         =   "DCD Auto (D)"
      End
      Begin VB.Menu m11 
         Caption         =   "V.32 Modulation"
      End
      Begin VB.Menu m12 
         Caption         =   "V.42BIS / MNPS"
      End
      Begin VB.Menu m13 
         Caption         =   "&64 Char. MNP Block"
      End
      Begin VB.Menu m14 
         Caption         =   "&128 Char MNP Block"
      End
      Begin VB.Menu m15 
         Caption         =   "&192 Char MNP Block"
      End
      Begin VB.Menu m16 
         Caption         =   "&256 Char MNP Block"
      End
   End
   Begin VB.Menu sr0 
      Caption         =   "&Sregs"
      Begin VB.Menu sr1 
         Caption         =   "S 00-10"
         Begin VB.Menu srr00 
            Caption         =   "S 00"
         End
         Begin VB.Menu srr01 
            Caption         =   "S 01"
         End
         Begin VB.Menu srr02 
            Caption         =   "S 02"
         End
         Begin VB.Menu srr03 
            Caption         =   "S 03"
         End
         Begin VB.Menu srr04 
            Caption         =   "S 04"
         End
         Begin VB.Menu srr05 
            Caption         =   "S 05"
         End
         Begin VB.Menu srr06 
            Caption         =   "S 06"
         End
         Begin VB.Menu srr07 
            Caption         =   "S 07"
         End
         Begin VB.Menu srr08 
            Caption         =   "S 08"
         End
         Begin VB.Menu srr09 
            Caption         =   "S 09"
         End
         Begin VB.Menu srr10 
            Caption         =   "S 10"
         End
      End
      Begin VB.Menu sr2 
         Caption         =   "S11-20"
         Begin VB.Menu srr11 
            Caption         =   "S 11"
         End
         Begin VB.Menu srr12 
            Caption         =   "S 12"
         End
         Begin VB.Menu srr13 
            Caption         =   "S 13"
         End
         Begin VB.Menu srr14 
            Caption         =   "S 14"
         End
         Begin VB.Menu srr15 
            Caption         =   "S 15"
         End
         Begin VB.Menu srr16 
            Caption         =   "S 16"
         End
         Begin VB.Menu srr17 
            Caption         =   "S 17"
         End
         Begin VB.Menu srr18 
            Caption         =   "S 18"
         End
         Begin VB.Menu srr19 
            Caption         =   "S 19"
         End
         Begin VB.Menu srr20 
            Caption         =   "S 20"
         End
      End
      Begin VB.Menu sr3 
         Caption         =   "S21-20"
         Begin VB.Menu srr21 
            Caption         =   "S 21"
         End
         Begin VB.Menu srr22 
            Caption         =   "S 22"
         End
         Begin VB.Menu srr23 
            Caption         =   "S 23"
         End
         Begin VB.Menu srr24 
            Caption         =   "S 24"
         End
         Begin VB.Menu srr25 
            Caption         =   "S 25"
         End
         Begin VB.Menu srr26 
            Caption         =   "S 26"
         End
         Begin VB.Menu srr27 
            Caption         =   "S 27"
         End
         Begin VB.Menu srr28 
            Caption         =   "S 28"
         End
         Begin VB.Menu srr29 
            Caption         =   "S 29"
         End
         Begin VB.Menu srr30 
            Caption         =   "S 30"
         End
      End
      Begin VB.Menu sr4 
         Caption         =   "S31-40"
         Begin VB.Menu srr31 
            Caption         =   "S 31"
         End
         Begin VB.Menu srr32 
            Caption         =   "S 32"
         End
         Begin VB.Menu srr33 
            Caption         =   "S 33"
         End
         Begin VB.Menu srr34 
            Caption         =   "S 34"
         End
         Begin VB.Menu srr35 
            Caption         =   "S 35"
         End
         Begin VB.Menu srr36 
            Caption         =   "S 36"
         End
         Begin VB.Menu srr37 
            Caption         =   "S 37"
         End
         Begin VB.Menu srr38 
            Caption         =   "S 38"
         End
         Begin VB.Menu srr39 
            Caption         =   "S 39"
         End
         Begin VB.Menu srr40 
            Caption         =   "S 40"
         End
      End
      Begin VB.Menu sr5 
         Caption         =   "S 41-50"
         Begin VB.Menu srr41 
            Caption         =   "S 41"
         End
         Begin VB.Menu srr42 
            Caption         =   "S 42"
         End
         Begin VB.Menu srr43 
            Caption         =   "S 43"
         End
         Begin VB.Menu srr44 
            Caption         =   "S 44"
         End
         Begin VB.Menu srr45 
            Caption         =   "S 45"
         End
         Begin VB.Menu srr46 
            Caption         =   "S 46"
         End
         Begin VB.Menu srr47 
            Caption         =   "S 47"
         End
         Begin VB.Menu srr48 
            Caption         =   "S 48"
         End
         Begin VB.Menu srr49 
            Caption         =   "S 49"
         End
         Begin VB.Menu srr50 
            Caption         =   "S 50"
         End
      End
   End
   Begin VB.Menu t0 
      Caption         =   "&Tests"
      Begin VB.Menu t1 
         Caption         =   "Ram Diagnostics Check"
      End
      Begin VB.Menu t2 
         Caption         =   "Link Diagnostics I"
      End
      Begin VB.Menu t3 
         Caption         =   "Link diagnostics II"
      End
      Begin VB.Menu t4 
         Caption         =   "Local Analog Loopback"
      End
      Begin VB.Menu t5 
         Caption         =   "Start ABL Test"
      End
      Begin VB.Menu t6 
         Caption         =   "End ABL Test"
      End
      Begin VB.Menu t8 
         Caption         =   "Digital Loopback"
      End
   End
   Begin VB.Menu reg 
      Caption         =   "&Reg"
      Begin VB.Menu reg1 
         Caption         =   "&Register"
      End
   End
   Begin VB.Menu h0 
      Caption         =   "&Help"
      Begin VB.Menu h4 
         Caption         =   "&Email Author"
      End
      Begin VB.Menu h1 
         Caption         =   "&Operation"
      End
      Begin VB.Menu h2 
         Caption         =   "&About"
      End
      Begin VB.Menu h3 
         Caption         =   "&Commands"
      End
   End
End
Attribute VB_Name = "frmTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'copyright (c) PWS 2001, Fresno, CA
' may only br redistributed with authors permission, for info, call PWS (559)255-9110
'--------------------------------------------------
Option Explicit
'im filename As String
Dim Ret As Integer      ' Scratch integer.
Dim Temp As String      ' Scratch string.
Dim hLogFile As Integer ' Handle of open log file.
Dim StartTime As Date

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation _
As String, ByVal lpFile As String, ByVal lpParameters _
As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Private Const SW_SHOW = 5
Public Function OpenEmail(ByVal EmailAddress As String, _
    Optional Subject As String, Optional Body As String) _
    As Boolean

    Dim lWindow As Long
    Dim lRet As Long
    Dim sParams As String
    
    sParams = EmailAddress
    If LCase(Left(sParams, 7)) <> "mailto:" Then _
        sParams = "mailto:" & sParams
 
 If Subject <> "" Then sParams = sParams & "?subject=" & Subject
        
    If Body <> "" Then
        sParams = sParams & IIf(Subject = "", "?", "&")
        sParams = sParams & "body=" & Body
    End If

   lRet = ShellExecute(lWindow, "open", sParams, _
    vbNullString, vbNullString, SW_SHOW)
    
   OpenEmail = lRet = 0

End Function
' Stores starting time for port timer
Private Sub ai1_Click()
On Error Resume Next
MSComm1.Output = "ata" + Chr$(13)
End Sub
Private Sub ap_Click()
On Error Resume Next
MSComm1.Output = "at&V0" + Chr$(13)
End Sub
Private Sub dcs_Click()
On Error Resume Next
MSComm1.Output = "AT&V3" + Chr$(13)
End Sub
Private Sub e1_Click()
On Error Resume Next
Clipboard.SetText frmTerminal.txtTerm.SelText
End Sub
Private Sub e2_Click()
On Error Resume Next
Clipboard.SetText frmTerminal.txtTerm.SelText
End Sub
Private Sub e3_Click()
On Error Resume Next
frmTerminal.txtTerm.SelText = Clipboard.GetText()
End Sub
Private Sub e4_Click()
On Error Resume Next
txtTerm.SelStart = 0
    txtTerm.SelLength = Len(txtTerm.Text)
End Sub
Private Sub fin_Click()
Call FileNew
End Sub
Private Sub fio_Click()
Call FileOpenProc
End Sub
Private Sub fis_Click()
Call FileSave
End Sub
Private Sub fisa_Click()
Call Savefilegui
End Sub
Private Sub Form_Load()
    Dim CommPort As String, Handshaking As String, Settings As String
        
    On Error Resume Next
    
    ' Set the default color for the terminal
    txtTerm.SelLength = Len(txtTerm)
    txtTerm.SelText = ""
    txtTerm.ForeColor = vbBlue
       
    ' Set Title
    App.Title = "Visual Basic Terminal"
    
    ' Set up status indicator light
    imgNotConnected.ZOrder
       
    ' Center Form
    frmTerminal.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    ' Load Registry Settings
    
    Settings = GetSetting(App.Title, "Properties", "Settings", "") ' frmTerminal.MSComm1.Settings]\
    If Settings <> "" Then
        MSComm1.Settings = Settings
        If Err Then
            MsgBox Error$, 48
            Exit Sub
        End If
    End If
    
    CommPort = GetSetting(App.Title, "Properties", "CommPort", "") ' frmTerminal.MSComm1.CommPort
    If CommPort <> "" Then MSComm1.CommPort = CommPort
    
    Handshaking = GetSetting(App.Title, "Properties", "Handshaking", "") 'frmTerminal.MSComm1.Handshaking
    If Handshaking <> "" Then
        MSComm1.Handshaking = Handshaking
        If Err Then
            MsgBox Error$, 48
            Exit Sub
        End If
    End If
    
    Echo = GetSetting(App.Title, "Properties", "Echo", "") ' Echo
    On Error GoTo 0

End Sub
Private Sub Form_Resize()
  On Error Resume Next
   ' Resize the Term (display) control
   txtTerm.Move 0, tbrToolBar.Height, frmTerminal.ScaleWidth, frmTerminal.ScaleHeight - sbrStatus.Height - tbrToolBar.Height
   
   ' Position the status indicator light
   Frame1.Left = ScaleWidth - Frame1.Width * 1.5
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim Counter As Long

    If MSComm1.PortOpen Then
       ' Wait 10 seconds for data to be transmitted.
       Counter = Timer + 10
       Do While MSComm1.OutBufferCount
          Ret = DoEvents()
          If Timer > Counter Then
             Select Case MsgBox("Data cannot be sent", 34)
                ' Cancel.
                Case 3
                   Cancel = True
                   Exit Sub
                ' Retry.
                Case 4
                   Counter = Timer + 10
                ' Ignore.
                Case 5
                   Exit Do
             End Select
          End If
       Loop

       MSComm1.PortOpen = 0
    End If

    ' If the log file is open, flush and close it.
    If hLogFile Then mnuCloseLog_Click
    End
End Sub
Private Sub h1_Click()
op.Show
End Sub
Private Sub h2_Click()
about.Show
End Sub
Private Sub h3_Click()
comds.Show
End Sub

Private Sub h4_Click()
OpenEmail "scott93727@aol.com", "Software Registration", "Please send registration information for the program:"
End Sub
Private Sub imgConnected_Click()
    ' Call the mnuOpen_Click routine to toggle connect and disconnect
    Call mnuOpen_Click
End Sub
Private Sub imgNotConnected_Click()
    ' Call the mnuOpen_Click routine to toggle connect and disconnect
    Call mnuOpen_Click
End Sub
Private Sub lp0_Click()
On Error Resume Next
MSComm1.Output = "ATZ0" + Chr$(13)
End Sub
Private Sub lp1_Click()
On Error Resume Next
MSComm1.Output = "ATZ1" + Chr$(13)
End Sub
Private Sub m1_Click()
On Error GoTo er1:
MSComm1.Output = "ata" + Chr$(13)
GoTo end1:
er1:
MsgBox "Port Must Be Open"
Resume Next
end1:
End Sub
Private Sub m10_Click()
On Error Resume Next
MSComm1.Output = "AT&C1" + Chr$(13)
End Sub
Private Sub m11_Click()
On Error Resume Next
MSComm1.Output = "AT&A2" + Chr$(13)
End Sub
Private Sub m12_Click()
On Error Resume Next
MSComm1.Output = "AT&A3" + Chr$(13)
End Sub
Private Sub m13_Click()
On Error Resume Next
MSComm1.Output = "AT\A0" + Chr$(13)
End Sub
Private Sub m14_Click()
On Error Resume Next
MSComm1.Output = "AT\A13" + Chr$(13)
End Sub
Private Sub m15_Click()
On Error Resume Next
MSComm1.Output = "AT\A2" + Chr$(13)
End Sub
Private Sub m16_Click()
On Error Resume Next
MSComm1.Output = "AT\A3" + Chr$(13)
End Sub
Private Sub m2_Click()
On Error GoTo er1:
MSComm1.Output = "atd" + Chr$(13)
GoTo end1:
er1:
MsgBox "Port Must Be Open"
Resume Next
end1:
End Sub
Private Sub m33_Click()
On Error GoTo er1:
MSComm1.Output = "atp" + Chr$(13)
GoTo end1:
er1:
MsgBox "Port Must Be Open"
Resume Next
end1:
End Sub
Private Sub m4_Click()
On Error GoTo er1:
MSComm1.Output = "att" + Chr$(13)
GoTo end1:
er1:
MsgBox "Port Must Be Open"
Resume Next
end1:
End Sub
Private Sub m5_Click()
On Error Resume Next
MSComm1.Output = "ATV1" + Chr$(13)
End Sub

Private Sub m6_Click()
On Error Resume Next
MSComm1.Output = "ATV0" + Chr$(13)
End Sub
Private Sub m7_Click()
On Error Resume Next
MSComm1.Output = "AT&A1" + Chr$(13)
End Sub
Private Sub m8_Click()
On Error Resume Next
MSComm1.Output = "AT&A0" + Chr$(13)
End Sub
Private Sub m9_Click()
On Error Resume Next
MSComm1.Output = "AT&C0" + Chr$(13)
End Sub
Private Sub mnuCloseLog_Click()
    ' Close the log file.
    Close hLogFile
    hLogFile = 0
    mnuOpenLog.Enabled = True
    tbrToolBar.Buttons("OpenLogFile").Enabled = True
    mnuCloseLog.Enabled = False
    tbrToolBar.Buttons("CloseLogFile").Enabled = False
    frmTerminal.Caption = "Visual Basic Terminal"
End Sub
Private Sub mnuDial_Click()
    On Local Error Resume Next
    Static Num As String
    
    Num = "1-206-936-6735" ' This is the MSDN phone number
    
    ' Get a number from the user.
    Num = InputBox$("Enter Phone Number:", "Dial Number", Num)
    If Num = "" Then Exit Sub
    
    ' Open the port if it isn't already open.
    If Not MSComm1.PortOpen Then
       mnuOpen_Click
       If Err Then Exit Sub
    End If
      
    ' Enable hang up button and menu item
    mnuHangUp.Enabled = True
    tbrToolBar.Buttons("HangUpPhone").Enabled = True
              
    ' Dial the number.
    MSComm1.Output = "ATDT" & Num & vbCrLf
    
    ' Start the port timer
    StartTiming
End Sub
' Toggle the DTREnabled property.
Private Sub mnuDTREnable_Click()
    ' Toggle DTREnable property
    MSComm1.DTREnable = Not MSComm1.DTREnable
    mnuDTREnable.Checked = MSComm1.DTREnable
End Sub
Private Sub mnuFileExit_Click()
    ' Use Form_Unload since it has code to check for unsent data and an open log file.
    Form_Unload Ret
End Sub
' Toggle the DTREnable property to hang up the line.
Private Sub mnuHangup_Click()
    On Error Resume Next
    
    MSComm1.Output = "ATH"      ' Send hangup string
    Ret = MSComm1.DTREnable     ' Save the current setting.
    MSComm1.DTREnable = True    ' Turn DTR on.
    MSComm1.DTREnable = False   ' Turn DTR off.
    MSComm1.DTREnable = Ret     ' Restore the old setting.
    mnuHangUp.Enabled = False
    tbrToolBar.Buttons("HangUpPhone").Enabled = False
    
    ' If port is actually still open, then close it
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
    
    ' Notify user of error
    If Err Then MsgBox Error$, 48
    
    mnuSendText.Enabled = False
    tbrToolBar.Buttons("TransmitTextFile").Enabled = False
    mnuHangUp.Enabled = False
    tbrToolBar.Buttons("HangUpPhone").Enabled = False
    mnuDial.Enabled = True
    tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
    sbrStatus.Panels("Settings").Text = "Settings: "
    
    ' Turn off indicator light and uncheck open menu
    mnuOpen.Checked = False
    imgNotConnected.ZOrder
            
    ' Stop the port timer
    StopTiming
    sbrStatus.Panels("Status").Text = "Status: "
    On Error GoTo 0
End Sub
' Display the value of the CDHolding property.
Private Sub mnuHCD_Click()
    If MSComm1.CDHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "CDHolding = " + Temp
End Sub
' Display the value of the CTSHolding property.
Private Sub mnuHCTS_Click()
    If MSComm1.CTSHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "CTSHolding = " + Temp
End Sub
' Display the value of the DSRHolding property.
Private Sub mnuHDSR_Click()
    If MSComm1.DSRHolding Then
        Temp = "True"
    Else
        Temp = "False"
    End If
    MsgBox "DSRHolding = " + Temp
End Sub

' This procedure sets the InputLen property, which determines how
' many bytes of data are read each time Input is used
' to retreive data from the input buffer.
' Setting InputLen to 0 specifies that
' the entire contents of the buffer should be read.
Private Sub mnuInputLen_Click()
    On Error Resume Next

    Temp = InputBox$("Enter New InputLen:", "InputLen", Str$(MSComm1.InputLen))
    If Len(Temp) Then
        MSComm1.InputLen = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If
End Sub
Private Sub mnuProperties_Click()
  ' Show the CommPort properties form
  frmProperties.Show vbModal
  
End Sub
' Toggles the state of the port (open or closed).
Private Sub mnuOpen_Click()
    On Error Resume Next
    Dim OpenFlag

    MSComm1.PortOpen = Not MSComm1.PortOpen
    If Err Then MsgBox Error$, 48
    
    OpenFlag = MSComm1.PortOpen
    
    mnuOpen.Checked = OpenFlag
    mnuSendText.Enabled = OpenFlag
    tbrToolBar.Buttons("TransmitTextFile").Enabled = OpenFlag
        
    If MSComm1.PortOpen Then
        ' Enable dial button and menu item
        mnuDial.Enabled = True
        tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
        
        ' Enable hang up button and menu item
        mnuHangUp.Enabled = True
        tbrToolBar.Buttons("HangUpPhone").Enabled = True
        
        imgConnected.ZOrder
        sbrStatus.Panels("Settings").Text = "Settings: " & MSComm1.Settings
        StartTiming
    Else
        ' Enable dial button and menu item
        mnuDial.Enabled = True
        tbrToolBar.Buttons("DialPhoneNumber").Enabled = True
        
        ' Disable hang up button and menu item
        mnuHangUp.Enabled = False
        tbrToolBar.Buttons("HangUpPhone").Enabled = False
        
        imgNotConnected.ZOrder
        sbrStatus.Panels("Settings").Text = "Settings: "
        StopTiming
    End If
    
End Sub
Private Sub mnuOpenLog_Click()
   Dim replace
   On Error Resume Next
   OpenLog.Flags = cdlOFNHideReadOnly Or cdlOFNExplorer
   OpenLog.CancelError = True
      
   ' Get the log filename from the user.
   OpenLog.DialogTitle = "Open Communications Log File"
   OpenLog.Filter = "Log Files (*.LOG)|*.log|All Files (*.*)|*.*"
   
   Do
      OpenLog.filename = ""
      OpenLog.ShowOpen
      If Err = cdlCancel Then Exit Sub
      Temp = OpenLog.filename

      ' If the file already exists, ask if the user wants to overwrite the file or add to it.
      Ret = Len(Dir$(Temp))
      If Err Then
         MsgBox Error$, 48
         Exit Sub
      End If
      If Ret Then
         replace = MsgBox("Replace existing file - " + Temp + "?", 35)
      Else
         replace = 0
      End If
   Loop While replace = 2

   ' User clicked the Yes button, so delete the file.
   If replace = 6 Then
      Kill Temp
      If Err Then
         MsgBox Error$, 48
         Exit Sub
      End If
   End If

   ' Open the log file.
   hLogFile = FreeFile
   Open Temp For Binary Access Write As hLogFile
   If Err Then
      MsgBox Error$, 48
      Close hLogFile
      hLogFile = 0
      Exit Sub
   Else
      ' Go to the end of the file so that new data can be appended.
      Seek hLogFile, LOF(hLogFile) + 1
   End If

   frmTerminal.Caption = "Visual Basic Terminal - " + OpenLog.FileTitle
   mnuOpenLog.Enabled = False
   tbrToolBar.Buttons("OpenLogFile").Enabled = False
   mnuCloseLog.Enabled = True
   tbrToolBar.Buttons("CloseLogFile").Enabled = True
End Sub
' This procedure sets the ParityReplace property, which holds the
' character that will replace any incorrect characters
' that are received because of a parity error.
Private Sub mnuParRep_Click()
    On Error Resume Next

    Temp = InputBox$("Enter Replace Character", "ParityReplace", frmTerminal.MSComm1.ParityReplace)
    frmTerminal.MSComm1.ParityReplace = Left$(Temp, 1)
    If Err Then MsgBox Error$, 48
End Sub
' This procedure sets the RThreshold property, which determines
' how many bytes can arrive at the receive buffer before the OnComm
' event is triggered and the CommEvent property is set to comEvReceive.
Private Sub mnuRThreshold_Click()
    On Error Resume Next
    
    Temp = InputBox$("Enter New RThreshold:", "RThreshold", Str$(MSComm1.RThreshold))
    If Len(Temp) Then
        MSComm1.RThreshold = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If

End Sub
Private Sub mp1_Click()
frmProperties.Show vbModal
End Sub
' The OnComm event is used for trapping communications events and errors.
Private Static Sub MSComm1_OnComm()
    Dim EVMsg$
    Dim ERMsg$
    
    ' Branch according to the CommEvent property.
    Select Case MSComm1.CommEvent
        ' Event messages.
        Case comEvReceive
            Dim Buffer As Variant
            Buffer = MSComm1.Input
            Debug.Print "Receive - " & StrConv(Buffer, vbUnicode)
            ShowData txtTerm, (StrConv(Buffer, vbUnicode))
        Case comEvSend
        Case comEvCTS
            EVMsg$ = "Change in CTS Detected"
        Case comEvDSR
            EVMsg$ = "Change in DSR Detected"
        Case comEvCD
            EVMsg$ = "Change in CD Detected"
        Case comEvRing
            EVMsg$ = "The Phone is Ringing"
        Case comEvEOF
            EVMsg$ = "End of File Detected"

        ' Error messages.
        Case comBreak
            ERMsg$ = "Break Received"
        Case comCDTO
            ERMsg$ = "Carrier Detect Timeout"
        Case comCTSTO
            ERMsg$ = "CTS Timeout"
        Case comDCB
            ERMsg$ = "Error retrieving DCB"
        Case comDSRTO
            ERMsg$ = "DSR Timeout"
        Case comFrame
            ERMsg$ = "Framing Error"
        Case comOverrun
            ERMsg$ = "Overrun Error"
        Case comRxOver
            ERMsg$ = "Receive Buffer Overflow"
        Case comRxParity
            ERMsg$ = "Parity Error"
        Case comTxFull
            ERMsg$ = "Transmit Buffer Full"
        Case Else
            ERMsg$ = "Unknown error or event"
    End Select
    
    If Len(EVMsg$) Then
        ' Display event messages in the status bar.
        sbrStatus.Panels("Status").Text = "Status: " & EVMsg$
                
        ' Enable timer so that the message in the status bar
        ' is cleared after 2 seconds
        Timer2.Enabled = True
        
    ElseIf Len(ERMsg$) Then
        ' Display event messages in the status bar.
        sbrStatus.Panels("Status").Text = "Status: " & ERMsg$
        
        ' Display error messages in an alert message box.
        Beep
        Ret = MsgBox(ERMsg$, 1, "Click Cancel to quit, OK to ignore.")
        
        ' If the user clicks Cancel (2)...
        If Ret = 2 Then
            MSComm1.PortOpen = False    ' Close the port and quit.
        End If
        
        ' Enable timer so that the message in the status bar
        ' is cleared after 2 seconds
        Timer2.Enabled = True
    End If
End Sub
Private Sub mnuSendText_Click()
   Dim hSend, BSize, LF&
   
   On Error Resume Next
   
   mnuSendText.Enabled = False
   tbrToolBar.Buttons("TransmitTextFile").Enabled = False
   
   ' Get the text filename from the user.
   OpenLog.DialogTitle = "Send Text File"
   OpenLog.Filter = "Text Files (*.TXT)|*.txt|All Files (*.*)|*.*"
   Do
      OpenLog.CancelError = True
      OpenLog.filename = ""
      OpenLog.ShowOpen
      If Err = cdlCancel Then
        mnuSendText.Enabled = True
        tbrToolBar.Buttons("TransmitTextFile").Enabled = True
        Exit Sub
      End If
      Temp = OpenLog.filename

      ' If the file doesn't exist, go back.
      Ret = Len(Dir$(Temp))
      If Err Then
         MsgBox Error$, 48
         mnuSendText.Enabled = True
         tbrToolBar.Buttons("TransmitTextFile").Enabled = True
         Exit Sub
      End If
      If Ret Then
         Exit Do
      Else
         MsgBox Temp + " not found!", 48
      End If
   Loop

   ' Open the log file.
   hSend = FreeFile
   Open Temp For Binary Access Read As hSend
   If Err Then
      MsgBox Error$, 48
   Else
      ' Display the Cancel dialog box.
      CancelSend = False
      frmCancelSend.Label1.Caption = "Transmitting Text File - " + Temp
      frmCancelSend.Show
      
      ' Read the file in blocks the size of the transmit buffer.
      BSize = MSComm1.OutBufferSize
      LF& = LOF(hSend)
      Do Until EOF(hSend) Or CancelSend
         ' Don't read too much at the end.
         If LF& - Loc(hSend) <= BSize Then
            BSize = LF& - Loc(hSend) + 1
         End If
      
         ' Read a block of data.
         Temp = Space$(BSize)
         Get hSend, , Temp
      
         ' Transmit the block.
         MSComm1.Output = Temp
         If Err Then
            MsgBox Error$, 48
            Exit Do
         End If
      
         ' Wait for all the data to be sent.
         Do
            Ret = DoEvents()
         Loop Until MSComm1.OutBufferCount = 0 Or CancelSend
      Loop
   End If
   
   Close hSend
   mnuSendText.Enabled = True
   tbrToolBar.Buttons("TransmitTextFile").Enabled = True
   CancelSend = True
   frmCancelSend.Hide
End Sub
' This procedure sets the SThreshold property, which determines
' how many characters (at most) have to be waiting
' in the output buffer before the CommEvent property
' is set to comEvSend and the OnComm event is triggered.
Private Sub mnuSThreshold_Click()
    On Error Resume Next
    
    Temp = InputBox$("Enter New SThreshold Value", "SThreshold", Str$(MSComm1.SThreshold))
    If Len(Temp) Then
        MSComm1.SThreshold = Val(Temp)
        If Err Then MsgBox Error$, 48
    End If
End Sub
' This procedure adds data to the Term control's Text property.
' It also filters control characters, such as BACKSPACE,
' carriage return, and line feeds, and writes data to
' an open log file.
' BACKSPACE characters delete the character to the left,
' either in the Text property, or the passed string.
' Line feed characters are appended to all carriage
' returns.  The size of the Term control's Text
' property is also monitored so that it never
' exceeds MAXTERMSIZE characters.
Private Static Sub ShowData(Term As Control, Data As String)
    On Error GoTo Handler
    Const MAXTERMSIZE = 16000
    Dim TermSize As Long, i
    
    ' Make sure the existing text doesn't get too large.
    TermSize = Len(Term.Text)
    If TermSize > MAXTERMSIZE Then
       Term.Text = Mid$(Term.Text, 4097)
       TermSize = Len(Term.Text)
    End If

    ' Point to the end of Term's data.
    Term.SelStart = TermSize

    ' Filter/handle BACKSPACE characters.
    Do
       i = InStr(Data, Chr$(8))
       If i Then
          If i = 1 Then
             Term.SelStart = TermSize - 1
             Term.SelLength = 1
             Data = Mid$(Data, i + 1)
          Else
             Data = Left$(Data, i - 2) & Mid$(Data, i + 1)
          End If
       End If
    Loop While i

    ' Eliminate line feeds.
    Do
       i = InStr(Data, Chr$(10))
       If i Then
          Data = Left$(Data, i - 1) & Mid$(Data, i + 1)
       End If
    Loop While i

    ' Make sure all carriage returns have a line feed.
    i = 1
    Do
       i = InStr(i, Data, Chr$(13))
       If i Then
          Data = Left$(Data, i) & Chr$(10) & Mid$(Data, i + 1)
          i = i + 1
       End If
    Loop While i

    ' Add the filtered data to the SelText property.
    Term.SelText = Data
  
    ' Log data to file if requested.
    If hLogFile Then
       i = 2
       Do
          Err = 0
          Put hLogFile, , Data
          If Err Then
             i = MsgBox(Error$, 21)
             If i = 2 Then
                mnuCloseLog_Click
             End If
          End If
       Loop While i <> 2
    End If
    Term.SelStart = Len(Term.Text)
Exit Sub

Handler:
    MsgBox Error$
    Resume Next
End Sub
Private Sub op1_Click()
OpenLog.CancelError = True
On Error GoTo ErrHandler
OpenLog.Flags = cdlCCRGBInit
OpenLog.ShowColor
txtTerm.ForeColor = OpenLog.Color
ErrHandler:
End Sub
Private Sub op2_Click()
OpenLog.CancelError = True
On Error GoTo ErrHandler
OpenLog.Flags = cdlCCRGBInit
OpenLog.ShowColor
txtTerm.BackColor = OpenLog.Color
ErrHandler:
End Sub
Private Sub opt_Click()
Dim ft
On Error GoTo ErrHandler3
OpenLog.Flags = cdlCFBoth
    OpenLog.ShowFont
    For ft = 0 To 161
    txtTerm.FontName = OpenLog.FontName
    Next ft
    For ft = 0 To 161
    txtTerm.FontBold = OpenLog.FontBold
    Next ft
    For ft = 0 To 161
    txtTerm.FontItalic = OpenLog.FontItalic
    Next ft
    For ft = 0 To 161
    txtTerm.FontStrikethru = OpenLog.FontStrikethru
    Next ft
    For ft = 1 To 161
    txtTerm.FontUnderline = OpenLog.FontUnderline
    Next ft
    GoTo fte
ErrHandler3:
MsgBox "Retry And Click On a Font Name"
fte:

End Sub

Private Sub prt1_Click()
Dim BeginPage, EndPage, NumCopies, i
OpenLog.CancelError = True
On Error GoTo ErrHandler
OpenLog.ShowPrinter
BeginPage = OpenLog.FromPage
EndPage = OpenLog.ToPage
NumCopies = OpenLog.Copies
For i = 1 To NumCopies
frmTerminal.PrintForm
' Printer code goes here
Next i
Exit Sub
ErrHandler:
MsgBox "Printer Not Ready"
End Sub
Private Sub reg1_Click()
regist.Show
End Sub
Private Sub rfc_Click()
On Error Resume Next
MSComm1.Output = "at+fclass=?" + Chr$(13)
End Sub
Private Sub rls_Click()
On Error Resume Next
MSComm1.Output = "at%Q" + Chr$(13)
End Sub
Private Sub rte_Click()
On Error Resume Next
MSComm1.Output = "ATW1" + Chr$(13)
End Sub
Private Sub s1_Click()
On Error GoTo er1:
MSComm1.Output = "ats$" + Chr$(13)
GoTo end1:
er1:
MsgBox "Port Must Be Open"
Resume Next
end1:
End Sub
Private Sub s10_Click()
On Error Resume Next
MSComm1.Output = "ati9" + Chr$(13)
End Sub
Private Sub s11_Click()
Dim BeginPage, EndPage, NumCopies, i
OpenLog.CancelError = True
On Error GoTo ErrHandler
OpenLog.ShowPrinter
BeginPage = OpenLog.FromPage
EndPage = OpenLog.ToPage
NumCopies = OpenLog.Copies
For i = 1 To NumCopies
frmTerminal.PrintForm
' Printer code goes here
Next i
Exit Sub
ErrHandler:
MsgBox "Printer Not Ready"
End Sub
Private Sub s12_Click()
On Error Resume Next
MSComm1.Output = "at&V0&V1&V2&V3" + Chr$(13)
End Sub
Private Sub s2_Click()
On Error GoTo er1:
MSComm1.Output = "at$" + Chr$(13)
GoTo end1:
er1:
MsgBox "Port Must Be Open"
Resume Next
end1:
End Sub
Private Sub s3_Click()
On Error GoTo er1:
MSComm1.Output = "at&$" + Chr$(13)
GoTo end1:
er1:
MsgBox "Port Must Be Open"
Resume Next
end1:
End Sub
Private Sub s4_Click()
On Error Resume Next
MSComm1.Output = "atI0" + Chr$(13)
End Sub
Private Sub s5_Click()
On Error Resume Next
MSComm1.Output = "atI1" + Chr$(13)
End Sub
Private Sub s6_Click()
On Error Resume Next
MSComm1.Output = "atI3" + Chr$(13)
End Sub
Private Sub s7_Click()
On Error Resume Next
MSComm1.Output = "atI4" + Chr$(13)
End Sub
Private Sub s8_Click()
On Error Resume Next
MSComm1.Output = "atI5" + Chr$(13)
End Sub
Private Sub s9_Click()
On Error Resume Next
MSComm1.Output = "atI7" + Chr$(13)
End Sub
Private Sub sldb_Click()
On Error Resume Next
MSComm1.Output = "AT%L" + Chr$(13)
End Sub
Private Sub sn_Click()
On Error Resume Next
MSComm1.Output = "at&V2" + Chr$(13)
End Sub
Private Sub sp_Click()
On Error Resume Next
MSComm1.Output = "at&V1" + Chr$(13)
End Sub
Private Sub srr00_Click()
On Error Resume Next
MSComm1.Output = "ats0?" + Chr$(13)
End Sub
Private Sub srr01_Click()
On Error Resume Next
MSComm1.Output = "ats1?" + Chr$(13)
End Sub
Private Sub srr02_Click()
On Error Resume Next
MSComm1.Output = "ats2?" + Chr$(13)
End Sub
Private Sub srr03_Click()
On Error Resume Next
MSComm1.Output = "ats3?" + Chr$(13)
End Sub
Private Sub srr04_Click()
On Error Resume Next
MSComm1.Output = "ats4?" + Chr$(13)
End Sub
Private Sub srr05_Click()
On Error Resume Next
MSComm1.Output = "ats5?" + Chr$(13)
End Sub
Private Sub srr06_Click()
On Error Resume Next
MSComm1.Output = "ats6?" + Chr$(13)
End Sub
Private Sub srr07_Click()
On Error Resume Next
MSComm1.Output = "ats7?" + Chr$(13)
End Sub
Private Sub srr08_Click()
On Error Resume Next
MSComm1.Output = "ats8?" + Chr$(13)
End Sub
Private Sub srr09_Click()
On Error Resume Next
MSComm1.Output = "ats9?" + Chr$(13)
End Sub
Private Sub srr10_Click()
On Error Resume Next
MSComm1.Output = "ats10?" + Chr$(13)
End Sub
Private Sub srr11_Click()
On Error Resume Next
MSComm1.Output = "ats11?" + Chr$(13)
End Sub
Private Sub srr12_Click()
On Error Resume Next
MSComm1.Output = "ats12?" + Chr$(13)
End Sub
Private Sub srr13_Click()
On Error Resume Next
MSComm1.Output = "ats13?" + Chr$(13)
End Sub
Private Sub srr14_Click()
On Error Resume Next
MSComm1.Output = "ats14?" + Chr$(13)
End Sub
Private Sub srr15_Click()
On Error Resume Next
MSComm1.Output = "ats145" + Chr$(13)
End Sub
Private Sub srr16_Click()
On Error Resume Next
MSComm1.Output = "ats16?" + Chr$(13)
End Sub
Private Sub srr17_Click()
On Error Resume Next
MSComm1.Output = "ats17?" + Chr$(13)
End Sub
Private Sub srr18_Click()
On Error Resume Next
MSComm1.Output = "ats18?" + Chr$(13)
End Sub
Private Sub srr19_Click()
On Error Resume Next
MSComm1.Output = "ats19?" + Chr$(13)
End Sub
Private Sub srr20_Click()
On Error Resume Next
MSComm1.Output = "ats20?" + Chr$(13)
End Sub
Private Sub srr21_Click()
On Error Resume Next
MSComm1.Output = "ats21?" + Chr$(13)
End Sub
Private Sub srr22_Click()
On Error Resume Next
MSComm1.Output = "ats22?" + Chr$(13)
End Sub
Private Sub srr23_Click()
On Error Resume Next
MSComm1.Output = "ats23?" + Chr$(13)
End Sub
Private Sub srr24_Click()
On Error Resume Next
MSComm1.Output = "ats24?" + Chr$(13)
End Sub
Private Sub srr25_Click()
On Error Resume Next
MSComm1.Output = "ats25?" + Chr$(13)
End Sub
Private Sub srr26_Click()
On Error Resume Next
MSComm1.Output = "ats26?" + Chr$(13)
End Sub
Private Sub srr27_Click()
On Error Resume Next
MSComm1.Output = "ats27?" + Chr$(13)
End Sub
Private Sub srr28_Click()
On Error Resume Next
MSComm1.Output = "ats28?" + Chr$(13)
End Sub
Private Sub srr29_Click()
On Error Resume Next
MSComm1.Output = "ats29?" + Chr$(13)
End Sub
Private Sub srr30_Click()
On Error Resume Next
MSComm1.Output = "ats30?" + Chr$(13)
End Sub
Private Sub srr31_Click()
On Error Resume Next
MSComm1.Output = "ats31?" + Chr$(13)
End Sub
Private Sub srr32_Click()
On Error Resume Next
MSComm1.Output = "ats32?" + Chr$(13)
End Sub
Private Sub srr33_Click()
On Error Resume Next
MSComm1.Output = "ats33?" + Chr$(13)
End Sub
Private Sub srr34_Click()
On Error Resume Next
MSComm1.Output = "ats34?" + Chr$(13)
End Sub
Private Sub srr35_Click()
On Error Resume Next
MSComm1.Output = "ats35?" + Chr$(13)
End Sub
Private Sub srr36_Click()
On Error Resume Next
MSComm1.Output = "ats36?" + Chr$(13)
End Sub
Private Sub srr37_Click()
On Error Resume Next
MSComm1.Output = "ats37?" + Chr$(13)
End Sub
Private Sub srr38_Click()
On Error Resume Next
MSComm1.Output = "ats38?" + Chr$(13)
End Sub
Private Sub srr39_Click()
On Error Resume Next
MSComm1.Output = "ats39?" + Chr$(13)
End Sub
Private Sub srr40_Click()
On Error Resume Next
MSComm1.Output = "ats40?" + Chr$(13)
End Sub
Private Sub srr41_Click()
On Error Resume Next
MSComm1.Output = "ats41?" + Chr$(13)
End Sub
Private Sub srr42_Click()
On Error Resume Next
MSComm1.Output = "ats42?" + Chr$(13)
End Sub
Private Sub srr43_Click()
On Error Resume Next
MSComm1.Output = "ats43?" + Chr$(13)
End Sub
Private Sub srr44_Click()
On Error Resume Next
MSComm1.Output = "ats44?" + Chr$(13)
End Sub
Private Sub srr45_Click()
On Error Resume Next
MSComm1.Output = "ats45?" + Chr$(13)
End Sub
Private Sub srr46_Click()
On Error Resume Next
MSComm1.Output = "ats46?" + Chr$(13)
End Sub
Private Sub srr47_Click()
On Error Resume Next
MSComm1.Output = "ats47?" + Chr$(13)
End Sub
Private Sub srr48_Click()
On Error Resume Next
MSComm1.Output = "ats48?" + Chr$(13)
End Sub
Private Sub srr49_Click()
On Error Resume Next
MSComm1.Output = "ats49?" + Chr$(13)
End Sub
Private Sub srr50_Click()
On Error Resume Next
MSComm1.Output = "ats50?" + Chr$(13)
End Sub
Private Sub ss1_Click()
On Error GoTo er1:
MSComm1.Output = "at&f0&f" + Chr$(13)
GoTo end1:
er1:
MsgBox "Port Must Be Open"
Resume Next
end1:
End Sub
Private Sub ss10_Click()
On Error Resume Next
MSComm1.Output = "at&f2" + Chr$(13)
End Sub
Private Sub ss2_Click()
On Error Resume Next
MSComm1.Output = "atz" + Chr$(13)
End Sub
Private Sub ss3_Click()
On Error Resume Next
MSComm1.Output = "atx6" + Chr$(13)
End Sub
Private Sub ss4_Click()
On Error Resume Next
MSComm1.Output = "atI1" + Chr$(13)
End Sub
Private Sub ss5_Click()
On Error Resume Next
MSComm1.Output = "atI2" + Chr$(13)
End Sub
Private Sub ss6_Click()
On Error Resume Next
MSComm1.Output = "atI3" + Chr$(13)
End Sub
Private Sub ss7_Click()
On Error Resume Next
MSComm1.Output = "atO1" + Chr$(13)
End Sub
Private Sub ss8_Click()
On Error Resume Next
MSComm1.Output = "atO0" + Chr$(13)
End Sub
Private Sub ss9_Click()
On Error Resume Next
MSComm1.Output = "at&F1" + Chr$(13)
End Sub
Private Sub t1_Click()
On Error Resume Next
MSComm1.Output = "atI2" + Chr$(13)
End Sub
Private Sub t2_Click()
On Error Resume Next
MSComm1.Output = "atI6" + Chr$(13)
End Sub
Private Sub t3_Click()
On Error Resume Next
MSComm1.Output = "atI11" + Chr$(13)
End Sub
Private Sub t4_Click()
On Error Resume Next
MSComm1.Output = "AT&M0" + Chr$(13)
MSComm1.Output = "ATS18=0&T1" + Chr$(13)
MSComm1.Output = " Echo Test In Progress 1234567890" + Chr$(13)
MSComm1.Output = "IF YOU SEE TEXT AND NUMBERS ABOVE MODEM PASTED" + Chr$(13)
MSComm1.Output = "+++" + Chr$(13)
MSComm1.Output = "AT&T0" + Chr$(13)
End Sub
Private Sub t5_Click()
On Error Resume Next
MSComm1.Output = "at&T1" + Chr$(13)
End Sub
Private Sub t6_Click()
On Error Resume Next
MSComm1.Output = "at&T0" + Chr$(13)
End Sub
Private Sub t8_Click()
On Error Resume Next
MSComm1.Output = "AT&M0" + Chr$(13)
MSComm1.Output = "ATS18=0&T8" + Chr$(13)
MSComm1.Output = "AT&T0" + Chr$(13)
End Sub
Private Sub Timer2_Timer()
sbrStatus.Panels("Status").Text = "Status: "
Timer2.Enabled = False
End Sub
' Keystrokes trapped here are sent to the MSComm
' control where they are echoed back via the
' OnComm (comEvReceive) event, and displayed
' with the ShowData procedure.
Private Sub txtTerm_KeyPress(KeyAscii As Integer)
    ' If the port is opened...
    If MSComm1.PortOpen Then
        ' Send the keystroke to the port.
        MSComm1.Output = Chr$(KeyAscii)
        
        ' Unless Echo is on, there is no need to
        ' let the text control display the key.
        ' A modem usually echos back a character
        If Not Echo Then
            ' Place position at end of terminal
            txtTerm.SelStart = Len(txtTerm)
            KeyAscii = 0
        End If
    End If
     
End Sub
Private Sub tbrToolBar_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
Case "OpenLogFile"
    Call mnuOpenLog_Click
Case "CloseLogFile"
    Call mnuCloseLog_Click
Case "DialPhoneNumber"
    Call mnuDial_Click
Case "HangUpPhone"
    Call mnuHangup_Click
Case "Properties"
    Call mnuProperties_Click
Case "TransmitTextFile"
    Call mnuSendText_Click
Case "A"
    Call mnuOpen_Click
Case "B"
    Call mnuDial_Click
Case "C"
    Call ai1_Click
Case "D"
    Call s7_Click
Case "E"
    Call m2_Click
Case "F"
    Call m12_Click
Case "G"
    Call ss1_Click
Case "H"
    Call s1_Click
Case "I"
    Call s2_Click
Case "J"
    Call s3_Click
Case "K"
    Dim BeginPage, EndPage, NumCopies, i
OpenLog.CancelError = True
On Error GoTo ErrHandler
OpenLog.ShowPrinter
BeginPage = OpenLog.FromPage
EndPage = OpenLog.ToPage
NumCopies = OpenLog.Copies
For i = 1 To NumCopies
frmTerminal.PrintForm
' Printer code goes here
Next i
Exit Sub
ErrHandler:
MsgBox "Printer Not Ready"
Case "L"
Dim ft
On Error GoTo ErrHandler3
OpenLog.Flags = cdlCFBoth
    OpenLog.ShowFont
    For ft = 0 To 161
    txtTerm.FontName = OpenLog.FontName
    Next ft
    For ft = 0 To 161
    txtTerm.FontBold = OpenLog.FontBold
    Next ft
    For ft = 0 To 161
    txtTerm.FontItalic = OpenLog.FontItalic
    Next ft
    For ft = 0 To 161
    txtTerm.FontStrikethru = OpenLog.FontStrikethru
    Next ft
    For ft = 1 To 161
    txtTerm.FontUnderline = OpenLog.FontUnderline
    Next ft
    GoTo fte
ErrHandler3:
MsgBox "Retry And Click On a Font Name"
fte:

Case "M"
    Call s8_Click

End Select
End Sub
Private Sub Timer1_Timer()
    ' Display the Connect Time
    sbrStatus.Panels("ConnectTime").Text = Format(Now - StartTime, "hh:nn:ss") & " "
End Sub
' Call this function to start the Connect Time timer
Private Sub StartTiming()
    StartTime = Now
    Timer1.Enabled = True
End Sub
' Call this function to stop timing
Private Sub StopTiming()
    Timer1.Enabled = False
    sbrStatus.Panels("ConnectTime").Text = ""
End Sub
Private Sub mnuRecentFile_Click(Index As Integer)
    ' Call the file open procedure, passing a
    ' reference to the selected file name
    OpenFile (mnuRecentFile(Index).Caption)
    ' Update the list of recently opened files in the File menu control array.
    GetRecentFiles
End Sub
