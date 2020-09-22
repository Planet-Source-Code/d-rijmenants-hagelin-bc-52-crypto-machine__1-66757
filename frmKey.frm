VERSION 5.00
Begin VB.Form frmKey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " BC-52 Key Setting"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8265
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lugs on Drum Bars"
      Height          =   4935
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   3975
      Begin VB.VScrollBar VScroll1 
         Height          =   1935
         Left            =   3480
         Max             =   32
         Min             =   1
         TabIndex        =   2
         Top             =   360
         Value           =   1
         Width           =   255
      End
      Begin VB.Label lblLugMsg 
         BackStyle       =   0  'Transparent
         Height          =   650
         Left            =   240
         TabIndex        =   89
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Use Scrollbar to select Bar and Click Lugs 1 - 6 to place or remove Lug."
         Height          =   495
         Left            =   240
         TabIndex        =   87
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "= Lug placed"
         Height          =   255
         Left            =   600
         TabIndex        =   86
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lbll 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   240
         TabIndex        =   85
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblLug 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bar 1"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   84
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblLug 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   3000
         TabIndex        =   80
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblLug 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   2640
         TabIndex        =   79
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblLug 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   2280
         TabIndex        =   78
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblLug 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   1920
         TabIndex        =   77
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblLug 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   1560
         TabIndex        =   76
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblLug 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   1200
         TabIndex        =   75
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wheel Configuration"
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3975
      Begin VB.PictureBox Picture2 
         Height          =   735
         Left            =   240
         ScaleHeight     =   675
         ScaleWidth      =   3435
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   360
         Width           =   3495
         Begin VB.Label lblC52 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "BC-52"
            Height          =   255
            Left            =   0
            TabIndex        =   74
            Top             =   105
            Width           =   3495
         End
         Begin VB.Label lblWheelSet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   73
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblWheelSet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   72
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblWheelSet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   1320
            TabIndex        =   71
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblWheelSet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   1680
            TabIndex        =   70
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblWheelSet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   2040
            TabIndex        =   69
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblWheelSet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   2400
            TabIndex        =   68
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.PictureBox picWheelBox 
         Height          =   1215
         Left            =   240
         ScaleHeight     =   1155
         ScaleWidth      =   3435
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   3360
         Width           =   3495
         Begin VB.Label lblBox 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Wheel Selection Box"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   75
            Width           =   3135
         End
         Begin VB.Label lblWheelInBox 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "47"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   2760
            TabIndex        =   65
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblWheelInBox 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "46"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   2263
            TabIndex        =   64
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblWheelInBox 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "43"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   1766
            TabIndex        =   63
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblWheelInBox 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "42"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   1269
            TabIndex        =   62
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblWheelInBox 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "41"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   772
            TabIndex        =   61
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblWheelInBox 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "38"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   275
            TabIndex        =   60
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblWheelInBox 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "37"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   2760
            TabIndex        =   59
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblWheelInBox 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "34"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   2263
            TabIndex        =   58
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblWheelInBox 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "31"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   1766
            TabIndex        =   57
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblWheelInBox 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "29"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   1269
            TabIndex        =   56
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblWheelInBox 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "26"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   772
            TabIndex        =   55
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblWheelInBox 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "25"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   275
            TabIndex        =   54
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Label lblClear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clear All Pins"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   240
         TabIndex        =   88
         Top             =   3000
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "= Active Pin"
         Height          =   255
         Left            =   2685
         TabIndex        =   83
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   2400
         TabIndex        =   82
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPrompt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Prompt"
         Height          =   375
         Left            =   0
         TabIndex        =   81
         Top             =   2040
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label lblWheelOut 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wheel"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1320
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "47"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   47
         Left            =   3480
         TabIndex        =   51
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "46"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   46
         Left            =   3480
         TabIndex        =   50
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "45"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   45
         Left            =   3120
         TabIndex        =   49
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "44"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   44
         Left            =   3120
         TabIndex        =   48
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "43"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   43
         Left            =   3120
         TabIndex        =   47
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "42"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   42
         Left            =   3120
         TabIndex        =   46
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "41"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   41
         Left            =   3120
         TabIndex        =   45
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "40"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   40
         Left            =   2760
         TabIndex        =   44
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "39"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   39
         Left            =   2760
         TabIndex        =   43
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "38"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   38
         Left            =   2760
         TabIndex        =   42
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "37"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   37
         Left            =   2760
         TabIndex        =   41
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   36
         Left            =   2760
         TabIndex        =   40
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "35"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   35
         Left            =   2400
         TabIndex        =   39
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "34"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   34
         Left            =   2400
         TabIndex        =   38
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "33"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   33
         Left            =   2400
         TabIndex        =   37
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "32"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   32
         Left            =   2400
         TabIndex        =   36
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "31"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   31
         Left            =   2400
         TabIndex        =   35
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   30
         Left            =   2040
         TabIndex        =   34
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "29"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   29
         Left            =   2040
         TabIndex        =   33
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "28"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   28
         Left            =   2040
         TabIndex        =   32
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "27"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   27
         Left            =   2040
         TabIndex        =   31
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "26"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   26
         Left            =   2040
         TabIndex        =   30
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "25"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   25
         Left            =   1680
         TabIndex        =   29
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "24"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   24
         Left            =   1680
         TabIndex        =   28
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "23"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   23
         Left            =   1680
         TabIndex        =   27
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "22"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   22
         Left            =   1680
         TabIndex        =   26
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "21"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   21
         Left            =   1680
         TabIndex        =   25
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   20
         Left            =   1320
         TabIndex        =   24
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "19"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   19
         Left            =   1320
         TabIndex        =   23
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "18"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   18
         Left            =   1320
         TabIndex        =   22
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "17"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   17
         Left            =   1320
         TabIndex        =   21
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "16"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   16
         Left            =   1320
         TabIndex        =   20
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "15"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   15
         Left            =   960
         TabIndex        =   19
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "14"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   14
         Left            =   960
         TabIndex        =   18
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "13"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   13
         Left            =   960
         TabIndex        =   17
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   12
         Left            =   960
         TabIndex        =   16
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "11"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   11
         Left            =   960
         TabIndex        =   15
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   10
         Left            =   600
         TabIndex        =   14
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "09"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   9
         Left            =   600
         TabIndex        =   13
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "08"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   8
         Left            =   600
         TabIndex        =   12
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "07"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   600
         TabIndex        =   11
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "06"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   600
         TabIndex        =   10
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "05"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "04"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "03"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "02"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tPin(6) As String
Private tLug(32) As String
Private tSel(6) As Integer

Private outPin As String
Private outSel As Integer

Private boxPin(12) As String
Private oldChange As Boolean

Private tLugPos As Integer
Private Const colRed = &H8080FF
Private Const colWhite = &HFFFFFF

Private Sub Form_Activate()
Dim k As Integer

oldChange = gblnKeyHasChanged

With Me
.Caption = "BC-52 Key - " & CutFilePath(gstrkeyFile)
For k = 0 To 47
    Me.lblPin(k).Visible = False
Next
Me.lblColor.Visible = False
Me.lblWheelOut.Visible = False
Me.lblClear.Visible = False
Me.lblPrompt.Visible = True
Me.lblPrompt.Caption = "Select wheel from BC-52 to remove it" & vbCrLf & "or select wheel in box"
Me.lblBox.Caption = "Wheel Selection Box"
Me.lblC52.Caption = "BC-52"

'show/hide selection box
If gblnCxType = True Then
    .picWheelBox.Visible = False
    .lblPrompt.Caption = "Select wheel from CX-52 to extract it" & vbCrLf & "and adjust the pins"
    .lblC52.Caption = "CX-52"
    Else
    .picWheelBox.Visible = True
    .lblPrompt.Caption = "Select wheel from C-52 to remove it"
    .lblC52.Caption = "C-52"
    End If
    
'show wheel setup
'transferr setup to tmp arrays
For k = 1 To 6
    .lblWheelSet(k).Caption = Len(Wpin(k))
    tPin(k) = Wpin(k)
    tSel(k) = Wsel(k)
Next
'clear used rotors in box
For k = 1 To 12
    .lblWheelInBox(k).Visible = True
Next
For k = 1 To 6
.lblWheelInBox(Wsel(k)).Visible = False
Next

Me.lblPrompt.Visible = True
outSel = 0
'transfer bars
For k = 1 To 32
    tLug(k) = BLug(k)
Next

Me.VScroll1.Value = 1
'set first bar
Call SetBarView(1)

End With
End Sub

Private Sub lblClear_Click()
Dim k As Integer
Dim retval As String
retval = MsgBox("Clear all pins on wheel " & Trim(Str(Len(outPin))) & "?", vbOKCancel + vbQuestion)
If retval = vbCancel Then Exit Sub
outPin = String(Len(outPin), "0")
For k = 1 To Len(outPin)
    Me.lblPin(k).BackColor = colWhite
Next
gblnKeyHasChanged = True
End Sub

Private Sub lblWheelSet_Click(Index As Integer)
'click on wheel in BC-52

Dim k As Integer
Dim tmp As String

gblnKeyHasChanged = True
If tSel(Index) <> 0 And outSel = 0 Then
    'pull out wheel
    outPin = tPin(Index)
    outSel = tSel(Index)
    'clear wheel in c52
    tPin(Index) = ""
    tSel(Index) = 0
    Me.lblWheelSet(Index).Caption = ""
    Me.lblWheelOut.Caption = "Wheel " & Trim(Str(Len(outPin)))
    Me.lblWheelOut.Visible = True
    'show pins
    Me.lblColor.Visible = True
    Me.lblPin(0).Visible = True 'dummy
    Me.lblClear.Visible = True
    For k = 1 To Len(outPin)
        Me.lblPin(k).Visible = True
        'pin on/off
        If Mid(outPin, k, 1) = "1" Then
            Me.lblPin(k).BackColor = colRed
            Else
            Me.lblPin(k).BackColor = colWhite
        End If
    Next
    Me.lblPrompt.Visible = False
    Me.lblBox.Caption = "Click here to put wheel in box"
    If gblnCxType = True Then
        Me.lblC52.Caption = "Click free place to insert wheel in CX-52"
        Else
        Me.lblC52.Caption = "Click free place to insert wheel in C-52"
        End If
    ElseIf tSel(Index) = 0 And outSel <> 0 Then
    'put wheel in BC-52
    tPin(Index) = outPin
    tSel(Index) = outSel
    Me.lblWheelSet(Index).Caption = Trim(Str(Len(outPin)))
    'clear wheel out
    outPin = ""
    outSel = 0
    For k = 0 To 47
        Me.lblPin(k).Visible = False
    Next
    Me.lblColor.Visible = False
    Me.lblWheelOut.Visible = False
    Me.lblClear.Visible = False
    Me.lblPrompt.Visible = True
    If gblnCxType = True Then
        Me.lblPrompt.Caption = "Select wheel from CX-52 to extract it" & vbCrLf & "and adjust the pins"
        Me.lblC52.Caption = "CX-52"
        Else
        Me.lblPrompt.Caption = "Select wheel from C-52 to remove it" & vbCrLf & "or select wheel in box"
        Me.lblC52.Caption = "C-52"
        End If
Me.lblBox.Caption = "Wheel Selection Box"
End If
End Sub

Private Sub lblWheelInBox_Click(Index As Integer)
'click on wheel in box
Dim k As Integer
Dim tmp As String

gblnKeyHasChanged = True
If outSel = 0 Then
    'get wheel from box
    outPin = boxPin(Index)
    If outPin = "" Then outPin = String(Len(W_textLabel(Index)), "0")
    outSel = Index
    'clear wheel in box
    boxPin(Index) = ""
    Me.lblWheelInBox(Index).Visible = False
    Me.lblWheelOut.Caption = "Wheel " & Trim(Str(Len(outPin)))
    Me.lblWheelOut.Visible = True
    'show pins
    Me.lblColor.Visible = True
    Me.lblPin(0).Visible = True 'dummy
    Me.lblClear.Visible = True
    For k = 1 To Len(outPin)
        Me.lblPin(k).Visible = True
        'pin on/off
        If Mid(outPin, k, 1) = "1" Then
            Me.lblPin(k).BackColor = colRed
            Else
            Me.lblPin(k).BackColor = colWhite
        End If
    Next
    Me.lblPrompt.Visible = False
    Me.lblBox.Caption = "Click here to put wheel in box"
    Me.lblC52.Caption = "Click free place to insert wheel in C-52"
End If
End Sub

Private Sub picWheelBox_Click()
If outSel = 0 Then Exit Sub
'put  rotor in box
Dim k As Integer
k = W_len(Len(outPin))
boxPin(k) = outPin
Me.lblWheelInBox(k).Visible = True
'clear wheel out
outPin = ""
outSel = 0
For k = 0 To 47
    Me.lblPin(k).Visible = False
Next
Me.lblColor.Visible = False
Me.lblWheelOut.Visible = False
Me.lblClear.Visible = False
Me.lblPrompt.Visible = True
Me.lblBox.Caption = "Wheel Selection Box"

If gblnCxType = True Then
    Me.lblPrompt.Caption = "Select wheel from CX-52 to remove it" & vbCrLf & "or select wheel in box"
    Me.lblC52.Caption = "CX-52"
    Else
    Me.lblPrompt.Caption = "Select wheel from C-52 to remove it" & vbCrLf & "or select wheel in box"
    Me.lblC52.Caption = "C-52"
    End If
End Sub

Private Sub lblBox_Click()
Call picWheelBox_Click
End Sub

Private Sub lblPin_Click(Index As Integer)
gblnKeyHasChanged = True
'click on pin
If Mid(outPin, Index, 1) = "0" Then
    'set pin
    Mid(outPin, Index, 1) = "1"
    Me.lblPin(Index).BackColor = colRed
    Else
    'clear pin
    Mid(outPin, Index, 1) = "0"
    Me.lblPin(Index).BackColor = colWhite
End If
End Sub

Private Sub lblLug_Click(Index As Integer)
'set lug
If Index = 0 Then Exit Sub
If Mid(tLug(tLugPos), Index, 1) = "0" Then
    'set lug
    Mid(tLug(tLugPos), Index, 1) = "1"
    Me.lblLug(Index).BackColor = colRed
    Else
    'clear pin
    Mid(tLug(tLugPos), Index, 1) = "0"
    Me.lblLug(Index).BackColor = colWhite
End If
gblnKeyHasChanged = True
End Sub

Private Sub SetBarView(bar As Integer)
Dim k As Integer
'set view from bars
Me.lblLug(0).Caption = "Bar " & Trim(Str(bar))
For k = 1 To 6
    If Mid(tLug(tLugPos), k, 1) = "0" Then
        Me.lblLug(k).BackColor = colWhite
        Else
        Me.lblLug(k).BackColor = colRed
    End If
Next
k = gintAdvanceBar(bar)
If k <> 0 Then
    Me.lblLugMsg.Caption = "Bar " & Trim(Str(bar)) & " is used for moving wheel " & Str(Str(k))
    If gintBarStepping(k) = 1 Then
        Me.lblLugMsg.Caption = Me.lblLugMsg.Caption & vbCrLf & "Wheel " & Str(Str(k)) & " moves if this bar is activated"
    ElseIf gintBarStepping(k) = 2 Then
        Me.lblLugMsg.Caption = Me.lblLugMsg.Caption & vbCrLf & "Wheel " & Str(Str(k)) & " moves if this bar is not activated"
    ElseIf gintBarStepping(k) = 3 Then
        Me.lblLugMsg.Caption = Me.lblLugMsg.Caption & vbCrLf & "Wheel " & Str(Str(k)) & " always moves!"
    ElseIf gintBarStepping(k) = 4 Then
        Me.lblLugMsg.Caption = Me.lblLugMsg.Caption & vbCrLf & "Wheel " & Str(Str(k)) & " will never move!"
    End If
    If gblnCipherBars = True Then Me.lblLugMsg.Caption = Me.lblLugMsg.Caption & vbCrLf & "This bar is also used for ciphering"
Else
    Me.lblLugMsg.Caption = ""
End If
End Sub

Private Sub cmdOK_Click()
Dim k As Integer
'check for empty wheels
For k = 1 To 6
    If tPin(k) = "" Then
    MsgBox "Please insert 6 wheels in BC-52", vbCritical
    Exit Sub
    End If
Next
'validate settings
'transfer tmps to effective sheels
For k = 1 To 6
    Wpin(k) = tPin(k)
    Wmax(k) = Len(Wpin(k))
    Wsel(k) = tSel(k)
    Wpos(k) = 1
Next
Call CreateLabels
Call resetAllWheels

For k = 1 To 6
    SetWheelView (k)
Next
'bars
For k = 1 To 32
    BLug(k) = tLug(k)
Next

For k = 0 To 47
    Me.lblPin(k).Visible = False
Next
Me.lblColor.Visible = False
Me.lblWheelOut.Visible = False
Me.lblClear.Visible = False
Me.lblPrompt.Visible = True
Me.lblPrompt.Caption = "Select wheel from BC-52 to remove it" & vbCrLf & "or select wheel in box"
Me.lblBox.Caption = "Wheel Selection Box"
Me.lblC52.Caption = "BC-52"

'If gblnKeyHasChanged = True Then gstrkeyFile = "Untitled"

Me.Hide

End Sub

Private Sub cmdCancel_Click()
gblnKeyHasChanged = oldChange
Me.Hide
End Sub

Private Sub VScroll1_Change()
tLugPos = Me.VScroll1.Value
Call SetBarView(tLugPos)
End Sub
