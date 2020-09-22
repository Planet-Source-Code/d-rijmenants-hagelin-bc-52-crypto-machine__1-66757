VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "C-52 Simulator"
   ClientHeight    =   8400
   ClientLeft      =   0
   ClientTop       =   -330
   ClientWidth     =   11055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dialog2 
      Left            =   1680
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   1080
      Top             =   5400
   End
   Begin VB.PictureBox picMenu 
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   2310
      Left            =   720
      ScaleHeight     =   2310
      ScaleWidth      =   1725
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   1725
      Begin VB.Label lblMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Help"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   20
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Label lblMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Edit Key"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   18
         Top             =   750
         Width           =   1695
      End
      Begin VB.Label lblMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Exit"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   17
         Top             =   2020
         Width           =   1695
      End
      Begin VB.Label lblMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Load Key"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   16
         Top             =   990
         Width           =   1695
      End
      Begin VB.Label lblMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Save Key"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   15
         Top             =   1230
         Width           =   1695
      End
      Begin VB.Label lblMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Customize BC-52"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   14
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label lblMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " View Key"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   13
         Top             =   510
         Width           =   1695
      End
      Begin VB.Label lblMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Auto Typing"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Clipboard"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   1560
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Image imgOffsetDn 
      Height          =   495
      Left            =   2880
      MouseIcon       =   "frmMain.frx":030A
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4365
      Width           =   615
   End
   Begin VB.Image imgOffsetUp 
      Height          =   495
      Left            =   2880
      MouseIcon       =   "frmMain.frx":0614
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3870
      Width           =   615
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Click here for simulator menu  or use F1 for help"
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   480
      TabIndex        =   19
      Top             =   1560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image imgHandleSwitch 
      Height          =   615
      Left            =   2520
      MouseIcon       =   "frmMain.frx":091E
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   495
   End
   Begin VB.Image imgMenu 
      Height          =   7575
      Left            =   360
      Stretch         =   -1  'True
      ToolTipText     =   " Click Here for Menu "
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblOffset 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A=A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image picHandle 
      Height          =   1020
      Left            =   4320
      Picture         =   "frmMain.frx":0C28
      Top             =   8520
      Width           =   75
   End
   Begin VB.Image imgHandle 
      Height          =   1020
      Left            =   2930
      Picture         =   "frmMain.frx":10AA
      Stretch         =   -1  'True
      Top             =   2250
      Width           =   75
   End
   Begin VB.Image imgAdvance 
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmMain.frx":152C
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Shape shpDot 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   150
   End
   Begin VB.Image picSoundOn 
      Height          =   255
      Left            =   480
      Picture         =   "frmMain.frx":1836
      Stretch         =   -1  'True
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image picSoundOff 
      Height          =   255
      Left            =   240
      Picture         =   "frmMain.frx":1D28
      Stretch         =   -1  'True
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgResetCounter 
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmMain.frx":221A
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   26
      Left            =   3840
      Picture         =   "frmMain.frx":2524
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   25
      Left            =   3600
      Picture         =   "frmMain.frx":3D4E
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   24
      Left            =   3360
      Picture         =   "frmMain.frx":5578
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   23
      Left            =   3120
      Picture         =   "frmMain.frx":6DA2
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   22
      Left            =   2880
      Picture         =   "frmMain.frx":85CC
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   21
      Left            =   2640
      Picture         =   "frmMain.frx":9DF6
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   20
      Left            =   2400
      Picture         =   "frmMain.frx":B620
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   19
      Left            =   2160
      Picture         =   "frmMain.frx":CE4A
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   18
      Left            =   1920
      Picture         =   "frmMain.frx":E674
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   17
      Left            =   1680
      Picture         =   "frmMain.frx":FE9E
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   16
      Left            =   1440
      Picture         =   "frmMain.frx":116C8
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   15
      Left            =   1200
      Picture         =   "frmMain.frx":12EF2
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   14
      Left            =   960
      Picture         =   "frmMain.frx":1471C
      Stretch         =   -1  'True
      Top             =   8760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   13
      Left            =   3840
      Picture         =   "frmMain.frx":15F46
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   12
      Left            =   3600
      Picture         =   "frmMain.frx":178B0
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   11
      Left            =   3360
      Picture         =   "frmMain.frx":190DA
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   10
      Left            =   3120
      Picture         =   "frmMain.frx":1A904
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   9
      Left            =   2880
      Picture         =   "frmMain.frx":1C12E
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   8
      Left            =   2640
      Picture         =   "frmMain.frx":1D958
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   7
      Left            =   2400
      Picture         =   "frmMain.frx":1F182
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   6
      Left            =   2160
      Picture         =   "frmMain.frx":209AC
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   5
      Left            =   1920
      Picture         =   "frmMain.frx":221D6
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   4
      Left            =   1680
      Picture         =   "frmMain.frx":23A00
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   3
      Left            =   1440
      Picture         =   "frmMain.frx":2522A
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   195
      Index           =   2
      Left            =   1200
      Picture         =   "frmMain.frx":26A54
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picKey 
      Height          =   200
      Index           =   1
      Left            =   960
      Picture         =   "frmMain.frx":2827E
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   26
      Left            =   6930
      MouseIcon       =   "frmMain.frx":29AA8
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   5920
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   25
      Left            =   3820
      MouseIcon       =   "frmMain.frx":29DB2
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   7335
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   24
      Left            =   4530
      MouseIcon       =   "frmMain.frx":2A0BC
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   7335
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   23
      Left            =   4060
      MouseIcon       =   "frmMain.frx":2A3C6
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   5920
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   3
      Left            =   5270
      MouseIcon       =   "frmMain.frx":2A6D0
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   7335
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   21
      Left            =   7650
      MouseIcon       =   "frmMain.frx":2A9DA
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   5920
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   20
      Left            =   6210
      MouseIcon       =   "frmMain.frx":2ACE4
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   5920
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   19
      Left            =   4298
      MouseIcon       =   "frmMain.frx":2AFEE
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   6620
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   18
      Left            =   5500
      MouseIcon       =   "frmMain.frx":2B2F8
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   5920
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   17
      Left            =   3330
      MouseIcon       =   "frmMain.frx":2B602
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   5920
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   16
      Left            =   9810
      MouseIcon       =   "frmMain.frx":2B90C
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   5920
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   15
      Left            =   9100
      MouseIcon       =   "frmMain.frx":2BC16
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   5920
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   14
      Left            =   7420
      MouseIcon       =   "frmMain.frx":2BF20
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   7335
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   13
      Left            =   8140
      MouseIcon       =   "frmMain.frx":2C22A
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   7335
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   12
      Left            =   9330
      MouseIcon       =   "frmMain.frx":2C534
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   6620
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   11
      Left            =   8606
      MouseIcon       =   "frmMain.frx":2C83E
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   6620
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   10
      Left            =   7900
      MouseIcon       =   "frmMain.frx":2CB48
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   6620
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   9
      Left            =   8370
      MouseIcon       =   "frmMain.frx":2CE52
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   5920
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   8
      Left            =   7180
      MouseIcon       =   "frmMain.frx":2D15C
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   6620
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   7
      Left            =   6460
      MouseIcon       =   "frmMain.frx":2D466
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   6620
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   6
      Left            =   5734
      MouseIcon       =   "frmMain.frx":2D770
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   6620
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   5
      Left            =   4770
      MouseIcon       =   "frmMain.frx":2DA7A
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   5925
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   4
      Left            =   5020
      MouseIcon       =   "frmMain.frx":2DD84
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   6620
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   22
      Left            =   5980
      MouseIcon       =   "frmMain.frx":2E08E
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   7335
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   2
      Left            =   6700
      MouseIcon       =   "frmMain.frx":2E398
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   7335
      Width           =   660
   End
   Begin VB.Image imgKey 
      Height          =   660
      Index           =   1
      Left            =   3580
      MouseIcon       =   "frmMain.frx":2E6A2
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   6620
      Width           =   660
   End
   Begin VB.Image imgHelp 
      Height          =   300
      Left            =   9870
      Stretch         =   -1  'True
      ToolTipText     =   " Help "
      Top             =   45
      Width           =   300
   End
   Begin VB.Image imgSound 
      Height          =   300
      Left            =   9470
      Picture         =   "frmMain.frx":2E9AC
      Stretch         =   -1  'True
      ToolTipText     =   " Sound Off "
      Top             =   45
      Width           =   300
   End
   Begin VB.Image imgAbout 
      Height          =   300
      Left            =   10290
      Stretch         =   -1  'True
      ToolTipText     =   " About C-52 "
      Top             =   45
      Width           =   300
   End
   Begin VB.Label lblCounter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3570
      TabIndex        =   8
      Top             =   1485
      Width           =   495
   End
   Begin VB.Label lblOutput 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   1245
      Width           =   4695
   End
   Begin VB.Label lblInput 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   855
      Width           =   4695
   End
   Begin VB.Image picTitleBar 
      Height          =   375
      Left            =   0
      MousePointer    =   15  'Size All
      Stretch         =   -1  'True
      ToolTipText     =   " Move Window "
      Top             =   0
      Width           =   9375
   End
   Begin VB.Image imgExit 
      Height          =   300
      Left            =   10680
      Stretch         =   -1  'True
      ToolTipText     =   " Exit Program "
      Top             =   45
      Width           =   300
   End
   Begin VB.Image WheelUp 
      Height          =   500
      Index           =   6
      Left            =   9150
      MouseIcon       =   "frmMain.frx":2EE9E
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4505
      Width           =   375
   End
   Begin VB.Image WheelUp 
      Height          =   500
      Index           =   5
      Left            =   8564
      MouseIcon       =   "frmMain.frx":2F1A8
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4505
      Width           =   375
   End
   Begin VB.Image WheelUp 
      Height          =   500
      Index           =   4
      Left            =   7978
      MouseIcon       =   "frmMain.frx":2F4B2
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4505
      Width           =   375
   End
   Begin VB.Image WheelUp 
      Height          =   500
      Index           =   3
      Left            =   7392
      MouseIcon       =   "frmMain.frx":2F7BC
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4505
      Width           =   375
   End
   Begin VB.Image WheelUp 
      Height          =   500
      Index           =   2
      Left            =   6806
      MouseIcon       =   "frmMain.frx":2FAC6
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4505
      Width           =   375
   End
   Begin VB.Image WheelUp 
      Height          =   500
      Index           =   1
      Left            =   6220
      MouseIcon       =   "frmMain.frx":2FDD0
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4505
      Width           =   375
   End
   Begin VB.Image WheelDn 
      Height          =   500
      Index           =   6
      Left            =   9150
      MouseIcon       =   "frmMain.frx":300DA
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3675
      Width           =   375
   End
   Begin VB.Image WheelDn 
      Height          =   500
      Index           =   5
      Left            =   8564
      MouseIcon       =   "frmMain.frx":303E4
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3675
      Width           =   375
   End
   Begin VB.Image WheelDn 
      Height          =   500
      Index           =   4
      Left            =   7978
      MouseIcon       =   "frmMain.frx":306EE
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3675
      Width           =   375
   End
   Begin VB.Image WheelDn 
      Height          =   500
      Index           =   3
      Left            =   7392
      MouseIcon       =   "frmMain.frx":309F8
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3675
      Width           =   375
   End
   Begin VB.Image WheelDn 
      Height          =   500
      Index           =   2
      Left            =   6806
      MouseIcon       =   "frmMain.frx":30D02
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3675
      Width           =   375
   End
   Begin VB.Image WheelDn 
      Height          =   500
      Index           =   1
      Left            =   6220
      MouseIcon       =   "frmMain.frx":3100C
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3675
      Width           =   375
   End
   Begin VB.Label lblWindow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1140
      Index           =   6
      Left            =   9150
      TabIndex        =   5
      Top             =   3790
      Width           =   375
   End
   Begin VB.Label lblWindow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1140
      Index           =   5
      Left            =   8564
      TabIndex        =   4
      Top             =   3790
      Width           =   375
   End
   Begin VB.Label lblWindow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1140
      Index           =   4
      Left            =   7978
      TabIndex        =   3
      Top             =   3790
      Width           =   375
   End
   Begin VB.Label lblWindow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1140
      Index           =   3
      Left            =   7392
      TabIndex        =   2
      Top             =   3790
      Width           =   375
   End
   Begin VB.Label lblWindow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1140
      Index           =   2
      Left            =   6806
      TabIndex        =   1
      Top             =   3790
      Width           =   375
   End
   Begin VB.Label lblWindow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1140
      Index           =   1
      Left            =   6220
      TabIndex        =   0
      Top             =   3790
      Width           =   375
   End
   Begin VB.Image imgOpen 
      Height          =   1815
      Left            =   5400
      MouseIcon       =   "frmMain.frx":31316
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Image imgBackground 
      Height          =   8400
      Left            =   0
      Picture         =   "frmMain.frx":31620
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pblnMoveFrom As Boolean
Private pBlnKeyLock As Boolean

Private Sub imgAbout_Click()
'about
If Me.picMenu.Visible = True Then Me.picMenu.Visible = False
frmInfo.Show (vbModal)
End Sub

Private Sub imgAdvance_Click()
'advance paper and check for maximum display length
If Me.picMenu.Visible = True Then Me.picMenu.Visible = False
Me.lblInput.Caption = Me.lblInput.Caption & " "
Me.lblOutput.Caption = Me.lblOutput.Caption & " "
If Len(Me.lblInput.Caption) > 34 Then
    Me.lblInput.Caption = Right(Me.lblInput.Caption, 34)
    Me.lblOutput.Caption = Right(Me.lblOutput.Caption, 34)
    Else
End If
PlaySound (2)
End Sub

Private Sub imgAdvance_DblClick()
Call imgAdvance_Click
End Sub

Private Sub imgBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.picMenu.Visible = True Then Me.picMenu.Visible = False
End Sub

Private Sub imgHandleSwitch_Click()
If Me.picMenu.Visible = True Then Me.picMenu.Visible = False
If gblnModeCipher = True Then
    gblnModeCipher = False
    Me.imgHandle.Picture = Nothing
    Else
    gblnModeCipher = True
    Me.imgHandle.Picture = Me.picHandle.Picture
End If
PlaySound (2)
End Sub

Private Sub imgHelp_Click()
'show helpfile
If Me.picMenu.Visible = True Then Me.picMenu.Visible = False
Call PlaySound(2)
Me.Dialog1.HelpFile = App.Path & "\BC-52.hlp"
Me.Dialog1.HelpCommand = cdlHelpContents
Me.Dialog1.ShowHelp
End Sub

Private Sub imgMenu_Click()
Dim k As Integer
If Me.Timer1.Enabled = True Then
    Me.lblInfo.Visible = False
    Me.Timer1.Enabled = False
    End If
If Me.lblOffset.Visible = True Then Me.lblOffset.Visible = False
If Me.picMenu.Visible = False Then
    For k = 0 To 8
        Me.lblMenu(k).BackColor = &HFFFFFF
    Next
    Me.imgMenu.ToolTipText = ""
    Me.picMenu.Visible = True
    Else
    Me.picMenu.Visible = False
    Me.imgMenu.ToolTipText = " Click Here for Menu "
End If
End Sub

Private Sub imgMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.picMenu.Visible = False Then Me.picMenu.Visible = True
End Sub

Private Sub imgOffsetDn_Click()
'dial offset -
If Me.picMenu.Visible = True Then Me.picMenu.Visible = False
If Me.lblOffset.Visible = False Then
    Me.lblOffset.Visible = True
    Me.lblOffset.Caption = "A=" & Chr(gintWheelOffset + 65)
    Else
    gintWheelOffset = gintWheelOffset - 1: If gintWheelOffset < 0 Then gintWheelOffset = 25
    Me.lblOffset.Caption = "A=" & Chr(gintWheelOffset + 65)
End If
Call PlaySound(2)
End Sub

Private Sub imgOffsetDn_DblClick()
Call imgOffsetDn_Click
End Sub

Private Sub imgOffsetUp_Click()
'dial offset +
If Me.picMenu.Visible = True Then Me.picMenu.Visible = False
If Me.lblOffset.Visible = False Then
    Me.lblOffset.Visible = True
    Me.lblOffset.Caption = "A=" & Chr(gintWheelOffset + 65)
    Else
    gintWheelOffset = gintWheelOffset + 1: If gintWheelOffset > 25 Then gintWheelOffset = 0
    Me.lblOffset.Caption = "A=" & Chr(gintWheelOffset + 65)
End If
Call PlaySound(2)
End Sub

Private Sub imgOffsetUp_DblClick()
Call imgOffsetUp_Click
End Sub

Private Sub imgOpen_Click()
If Me.picMenu.Visible = True Then Me.picMenu.Visible = False
frmKey.Show (vbModal)
End Sub

Private Sub imgResetCounter_Click()
'rest sounter
If Me.picMenu.Visible = True Then Me.picMenu.Visible = False
Me.lblCounter.Caption = "000"
glngGroupCount = 0
gstrCounter = 0
PlaySound (2)
End Sub

Private Sub imgSound_Click()
If Me.picMenu.Visible = True Then Me.picMenu.Visible = False
If gblnSound = True Then
    gblnSound = False
    Me.imgSound.ToolTipText = " Sound On "
    Me.imgSound.Picture = Me.picSoundOff.Picture
    Else
    gblnSound = True
    Me.imgSound.ToolTipText = " Sound Off "
    Me.imgSound.Picture = Me.picSoundOn.Picture
End If
End Sub


Private Sub lblInfo_Click()
Me.Timer1.Enabled = False
Me.lblInfo.Visible = False
Call imgMenu_Click
End Sub

Private Sub lblInput_Click()
'clipboard
frmClipBoard.Show (vbModal)
End Sub

Private Sub lblMenu_Click(Index As Integer)
'menu click
Me.picMenu.Visible = False
Select Case Index
Case 0
    frmClipBoard.Show (vbModal)
Case 1
    frmQuick.Show (vbModal)
Case 2
    frmKeySheet.Show (vbModal)
Case 3
    frmCustom.Show (vbModal)
Case 4
    Call SaveFile
Case 5
    Call OpenFile
Case 6
    Call EndProgram
Case 7
    frmKey.Show (vbModal)
Case 8
    Call imgHelp_Click
End Select

End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Integer
For k = 0 To 8
    Me.lblMenu(k).BackColor = &HFFFFFF
Next
Me.lblMenu(Index).BackColor = &HC0C0C0
End Sub

Private Sub lblOutput_Click()
'clipboard
frmClipBoard.Show (vbModal)
End Sub

Private Sub WheelDn_Click(Index As Integer)
If Me.picMenu.Visible = True Then Me.picMenu.Visible = False
Wpos(Index) = Wpos(Index) - 1
PlaySound (2)
If Wpos(Index) < 0 Then Wpos(Index) = Wmax(Index)
SetWheelView (Index)
End Sub

Private Sub WheelDn_DblClick(Index As Integer)
Call WheelDn_Click(Index)
End Sub

Private Sub WheelUp_Click(Index As Integer)
If Me.picMenu.Visible = True Then Me.picMenu.Visible = False
PlaySound (2)
Wpos(Index) = Wpos(Index) + 1
If Wpos(Index) > Wmax(Index) Then Wpos(Index) = 1
SetWheelView (Index)
End Sub

Private Sub WheelUp_DblClick(Index As Integer)
Call WheelUp_Click(Index)
End Sub

Private Sub imgExit_Click()
'exit program
'Call PlaySound(2)
If Me.picMenu.Visible = True Then Me.picMenu.Visible = False
Call EndProgram
End Sub

'--------------------------------------------------------------------
' form movement
'--------------------------------------------------------------------

Private Sub picTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'get mouse movement
Dim POINT As POINTAPI
GetCursorPos POINT
LastPoint.X = POINT.X
LastPoint.Y = POINT.Y
pblnMoveFrom = True
End Sub

Private Sub picTitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'if mouse is down, move the form
Dim iDX As Long, iDY As Long
Dim POINT As POINTAPI
If Not pblnMoveFrom Then
    Exit Sub
    End If
GetCursorPos POINT
iDX& = (POINT.X - LastPoint.X) * iTPPX&
iDY& = (POINT.Y - LastPoint.Y) * iTPPY&
LastPoint.X = POINT.X
LastPoint.Y = POINT.Y
Me.Move Me.Left + iDX&, Me.Top + iDY&
End Sub

Private Sub picTitleBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'release form
pblnMoveFrom = False
End Sub

'---keys
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If pBlnKeyLock = True Then Exit Sub
Dim k As Integer
Dim i As Integer

If Me.picMenu.Visible = True Then Me.picMenu.Visible = False
If Me.lblOffset.Visible = True Then Me.lblOffset.Visible = False

Select Case KeyCode
Case 64 To 90, 32
    'Call PlaySound(2)
    pBlnKeyLock = True
    If KeyCode = 32 And gblnModeCipher = True Then
        'if cipher mode, replace space by gstrSpaceLetter
        KeyCode = Asc(gstrSpaceLetter)
    ElseIf KeyCode = 32 And gblnModeCipher = False Then
        KeyCode = 0
        Exit Sub
    End If
    'set dial
    gintAlphaWheel = KeyCode - 64
    Me.imgKey(gintAlphaWheel).Picture = Me.picKey(gintAlphaWheel).Picture
    k = gintLastDialView
    If gblnFastRun = False Then
        Do While k <> gintAlphaWheel
            Call SetDialView(k)
            PauzeTime (5)
            k = k + 1
            If k > 26 Then k = k - 26
            PlaySound (2)
        Loop
    End If
    Call SetDialView(gintAlphaWheel)
Case 46
    'delet ribbon
    glngGroupCount = 0
    gstrInput = ""
    gstrOutput = ""
    frmMain.lblInput.Caption = ""
    frmMain.lblOutput.Caption = ""
    glngGroupCount = 0
Case 116
    'clipboard
    frmClipBoard.Show (vbModal)
Case 119
    'key sheet
    frmKeySheet.Show (vbModal)
Case 117
    'autotyping
    frmQuick.Show (vbModal)
Case 45
    'INS, memorize wheel positions
    Call MemorizeWheels
Case 36
    'HOME, reset all wheels to memorized pos
    For k = 1 To 6
        i = gintPosMemo(k)
        If i <> 0 And i <= Wmax(k) Then
            Wpos(k) = i
            Else
            Wpos(k) = 1
        End If
        SetWheelView (k)
    Next
Case 27
    'ESC, abort autotyping
    gstrAutoType = False
Case 121
    'customize wheels and machine
    frmCustom.Show (vbModal)
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 64 To 90, 32
    'encrypt letter
    If KeyCode = 32 And gblnModeCipher = True Then
        'if cipher mode, replace space by gstrSpaceLetter
        KeyCode = Asc(gstrSpaceLetter)
    ElseIf KeyCode = 32 And gblnModeCipher = False Then
        KeyCode = 0
        pBlnKeyLock = False
        Exit Sub
    End If
    gintAlphaWheel = KeyCode - 64
    Me.imgKey(gintAlphaWheel).Picture = Nothing
    If gblnFastRun = False Then
        PlaySound (1)
        Else
        Call PlaySound(2)
        End If
Call Crypto(gintAlphaWheel)
End Select
pBlnKeyLock = False
End Sub

Private Sub imgKey_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_KeyDown(Index + 64, 0)
If Me.lblOffset.Visible = True Then Me.lblOffset.Visible = False
End Sub

Private Sub imgKey_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_KeyUp(Index + 64, 0)
If Me.lblOffset.Visible = True Then Me.lblOffset.Visible = False
End Sub

Private Sub Timer1_Timer()
' display flashing message at start pointing to readme
Static flashCount As Integer
flashCount = flashCount + 1
Select Case flashCount
    Case 1, 3, 5, 7
        Me.lblInfo.Visible = False
    Case 2, 4, 6, 8
        Me.lblInfo.Visible = True
    Case 24
        Me.Timer1.Enabled = False
        Me.lblInfo.Visible = False
End Select
End Sub



