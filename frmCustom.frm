VERSION 5.00
Begin VB.Form frmCustom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " BC-52 Customizing"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Caption         =   "Program Speed"
      Height          =   855
      Left            =   120
      TabIndex        =   62
      Top             =   5280
      Width           =   4815
      Begin VB.CheckBox chkSpeed 
         Caption         =   "Disable graphics delay"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Machine Type"
      Height          =   1335
      Left            =   120
      TabIndex        =   55
      Top             =   120
      Width           =   4815
      Begin VB.OptionButton optC52 
         Caption         =   "C-52 (6 different wheels with 25 up to 47 pins)"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   4335
      End
      Begin VB.OptionButton optCX52 
         Caption         =   "CX-52 (6 wheels with 47 pins each)"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "&Default"
      Height          =   375
      Left            =   5040
      TabIndex        =   29
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Space Letter"
      Height          =   1695
      Left            =   5040
      TabIndex        =   47
      Top             =   4440
      Width           =   3975
      Begin VB.ComboBox cmbSpace 
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Text            =   "Combo1"
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Select the letter that will represent a space during ciphering. When deciphering, this letter will be replace by a space."
         Height          =   735
         Left            =   240
         TabIndex        =   48
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Wheel Labeling Setup"
      Height          =   3615
      Left            =   120
      TabIndex        =   32
      Top             =   1560
      Width           =   4815
      Begin VB.ComboBox cmbLabel 
         Height          =   315
         Index           =   12
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2880
         Width           =   615
      End
      Begin VB.ComboBox cmbLabel 
         Height          =   315
         Index           =   11
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2880
         Width           =   615
      End
      Begin VB.ComboBox cmbLabel 
         Height          =   315
         Index           =   10
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2880
         Width           =   615
      End
      Begin VB.ComboBox cmbLabel 
         Height          =   315
         Index           =   9
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2880
         Width           =   615
      End
      Begin VB.ComboBox cmbLabel 
         Height          =   315
         Index           =   8
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2880
         Width           =   615
      End
      Begin VB.ComboBox cmbLabel 
         Height          =   315
         Index           =   7
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2880
         Width           =   615
      End
      Begin VB.ComboBox cmbLabel 
         Height          =   315
         Index           =   6
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2160
         Width           =   615
      End
      Begin VB.ComboBox cmbLabel 
         Height          =   315
         Index           =   5
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2160
         Width           =   615
      End
      Begin VB.ComboBox cmbLabel 
         Height          =   315
         Index           =   4
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2160
         Width           =   615
      End
      Begin VB.ComboBox cmbLabel 
         Height          =   315
         Index           =   3
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2160
         Width           =   615
      End
      Begin VB.ComboBox cmbLabel 
         Height          =   315
         Index           =   2
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   615
      End
      Begin VB.ComboBox cmbLabel 
         Height          =   315
         Index           =   1
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmCustom.frx":0000
         Height          =   855
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "47"
         Height          =   255
         Index           =   11
         Left            =   3840
         TabIndex        =   45
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "46"
         Height          =   255
         Index           =   10
         Left            =   3120
         TabIndex        =   44
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "43"
         Height          =   255
         Index           =   9
         Left            =   2400
         TabIndex        =   43
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "42"
         Height          =   255
         Index           =   8
         Left            =   1680
         TabIndex        =   42
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "41"
         Height          =   255
         Index           =   7
         Left            =   960
         TabIndex        =   41
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "38"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   40
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "37"
         Height          =   255
         Index           =   5
         Left            =   3840
         TabIndex        =   39
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "34"
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   38
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "31"
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   37
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "29"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   36
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "26"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   35
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "25"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Set the label that is visible on the exteriour of the machine when Pin 01 is in front of the Pin Reading Pawl."
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   1320
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6720
      TabIndex        =   30
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7920
      TabIndex        =   0
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wheel Advancing and Pin Reading Method"
      Height          =   4215
      Left            =   5040
      TabIndex        =   31
      Top             =   120
      Width           =   3975
      Begin VB.ComboBox cmbStep 
         Height          =   315
         Index           =   2
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2880
         Width           =   615
      End
      Begin VB.ComboBox cmbStep 
         Height          =   315
         Index           =   3
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2880
         Width           =   615
      End
      Begin VB.ComboBox cmbStep 
         Height          =   315
         Index           =   4
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2880
         Width           =   615
      End
      Begin VB.ComboBox cmbStep 
         Height          =   315
         Index           =   5
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2880
         Width           =   615
      End
      Begin VB.ComboBox cmbStep 
         Height          =   315
         Index           =   6
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2880
         Width           =   615
      End
      Begin VB.ComboBox cmbBar 
         Height          =   315
         Index           =   6
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox cmbBar 
         Height          =   315
         Index           =   5
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox cmbBar 
         Height          =   315
         Index           =   4
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox cmbBar 
         Height          =   315
         Index           =   3
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox cmbBar 
         Height          =   315
         Index           =   2
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1200
         Width           =   615
      End
      Begin VB.CheckBox chkCipherBars 
         Caption         =   "Use advance bars also for ciphering"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   3480
         Width           =   3015
      End
      Begin VB.CheckBox chkHoldPins 
         Caption         =   "Hold pin positions while turning drum"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   3840
         Width           =   3015
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmCustom.frx":00CB
         Height          =   855
         Left            =   240
         TabIndex        =   61
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   9
         Left            =   375
         TabIndex        =   60
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3"
         Height          =   255
         Index           =   8
         Left            =   1095
         TabIndex        =   59
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "4"
         Height          =   255
         Index           =   7
         Left            =   1830
         TabIndex        =   58
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5"
         Height          =   255
         Index           =   6
         Left            =   2550
         TabIndex        =   57
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "6"
         Height          =   255
         Index           =   5
         Left            =   3285
         TabIndex        =   56
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Select for wheels 2 to 6 the bar that will advance that wheel. (wheel 1 always moves)"
         Height          =   495
         Left            =   240
         TabIndex        =   54
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "6"
         Height          =   255
         Index           =   4
         Left            =   3285
         TabIndex        =   53
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "5"
         Height          =   255
         Index           =   3
         Left            =   2550
         TabIndex        =   52
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "4"
         Height          =   255
         Index           =   2
         Left            =   1830
         TabIndex        =   51
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "3"
         Height          =   255
         Index           =   1
         Left            =   1095
         TabIndex        =   50
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   0
         Left            =   375
         TabIndex        =   49
         Top             =   960
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim k As Integer
Dim j As Integer
Dim tmp As String

For k = 1 To 12
    For j = 1 To Len(W_textLabel(k))
        If Mid(W_textLabel(k), j, 1) <> "." Then
            Me.cmbLabel(k).AddItem Mid(W_textLabel(k), j, 1)
            Else
            Me.cmbLabel(k).AddItem Format(j, "00")
        End If
    Next
Next
For k = 1 To 26
    Me.cmbSpace.AddItem Chr(k + 64)
Next k
For k = 2 To 6
    For j = 1 To 32
        Me.cmbBar(k).AddItem Format(j, "00")
    Next
Next

For k = 2 To 6
    For j = 1 To 4
        Me.cmbStep(k).AddItem Format(j, "0")
    Next
Next

End Sub

Private Sub Form_Activate()
Dim k As Integer
Me.cmdCancel.SetFocus

If gblnCxType = True Then
    Me.optCX52.Value = True
    Else
    Me.optC52.Value = True
    End If

For k = 1 To 12
    Me.cmbLabel(k).ListIndex = gintLabelView(k) - 1
Next
If gblnCipherBars = True Then
    Me.chkCipherBars.Value = 1
    Else
    Me.chkCipherBars.Value = 0
End If
If gblnHoldPins = True Then
    Me.chkHoldPins.Value = 1
    Else
    Me.chkHoldPins.Value = 0
End If
Me.cmbSpace.ListIndex = Asc(gstrSpaceLetter) - 65

For k = 1 To 32
    If gintAdvanceBar(k) <> 0 Then
        Me.cmbBar(gintAdvanceBar(k)).ListIndex = k - 1
    End If
Next

For k = 2 To 6
    Me.cmbStep(k).ListIndex = gintBarStepping(k) - 1
Next

If gblnFastRun = True Then
    Me.chkSpeed.Value = 1
    Else
    Me.chkSpeed.Value = 0
End If

End Sub

Private Sub cmdOK_Click()
'save settings to regestry
Dim k As Integer
Dim j As Integer
Dim tmp As String

'check for double advance bars
For k = 2 To 6
    For j = k + 1 To 6
        If Me.cmbBar(k).ListIndex = Me.cmbBar(j).ListIndex Then
            MsgBox "Please select different bar numbers for each wheel!", vbCritical
            Exit Sub
        End If
    Next
Next


For k = 1 To 12
    j = Me.cmbLabel(k).ListIndex
    gintLabelView(k) = j + 1
    tmp = tmp & Format(j + 1, "00")
Next

'set advance bars
For k = 1 To 32
    gintAdvanceBar(k) = 0
Next
For k = 2 To 6
    j = Me.cmbBar(k).ListIndex + 1
    gintAdvanceBar(j) = k
    tmp = tmp & Format(j, "00")
Next

If Me.chkCipherBars.Value = 1 Then
    tmp = tmp & "1"
    gblnCipherBars = True
    Else
    tmp = tmp & "0"
    gblnCipherBars = False
End If

If Me.chkHoldPins.Value = 1 Then
    tmp = tmp & "1"
    gblnHoldPins = True
    Else
    tmp = tmp & "0"
    gblnHoldPins = False
End If

If Me.optCX52.Value = True Then
    If gblnCxType = False Then Call SetCX52wheels
    gblnCxType = True
    tmp = tmp & "1"
    Else
    If gblnCxType = True Then Call SetC52wheels
    gblnCxType = False
    tmp = tmp & "0"
    End If

gstrSpaceLetter = Me.cmbSpace.Text
tmp = tmp & Me.cmbSpace.Text

For k = 2 To 6
    gintBarStepping(k) = Me.cmbStep(k).ListIndex + 1
    tmp = tmp & Format(gintBarStepping(k), "0")
Next

gblnFastRun = Me.chkSpeed.Value

If Me.chkSpeed.Value = 0 Then
    gblnFastRun = False
    Else
    gblnFastRun = True
End If
SaveSetting App.EXEName, "config", "fastrun", Trim(Str(Me.chkSpeed.Value))

gstrMachineSetup = tmp

SaveSetting App.EXEName, "config", "setup", gstrMachineSetup
Me.Hide
For k = 1 To 6
    Call SetWheelView(k)
Next k
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdDefault_Click()
Dim k As Integer
For k = 1 To 12
    gintLabelView(k) = Val(Mid(DefaultSetup, (k * 2) - 1, 2))
    Me.cmbLabel(k).ListIndex = gintLabelView(k) - 1
Next
For k = 2 To 6
Me.cmbBar(k).ListIndex = k - 2
Next

For k = 2 To 6
Me.cmbStep(k).ListIndex = 0
Next

Me.chkCipherBars.Value = 0
Me.chkHoldPins.Value = 1
Me.cmbSpace.ListIndex = 23
Me.optC52.Value = True
Me.chkSpeed.Value = 0

End Sub
