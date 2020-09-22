VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " About BC-52"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5760
   ControlBox      =   0   'False
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   5535
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   1035
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   300
         Picture         =   "frmInfo.frx":000C
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   480
      End
   End
   Begin VB.Label lblInfo 
      Height          =   2655
      Left            =   1320
      TabIndex        =   5
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hagelin BC-52  Cipher Machine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   1320
      TabIndex        =   1
      Top             =   2040
      Width           =   3855
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Me.lblInfo.Caption = "Version 3.6" & vbCrLf & vbCrLf & _
"Program written by D. Rijmenants" & vbCrLf & vbCrLf & "© D. Rijmenants 2006" & vbCrLf & vbCrLf & _
"This program is freeware and can be used and distributed under the following restrictions: It is strictly forbidden to use this software, copies or parts of it for commercial purposes, or to sell, to lease or make profit from this program by any means." & vbCrLf & vbCrLf & "For more info read the Help File (F1)."
End Sub
