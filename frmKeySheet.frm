VERSION 5.00
Begin VB.Form frmKeySheet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " BC-52 Key Settings"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4815
   ControlBox      =   0   'False
   Icon            =   "frmKeySheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtKey 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy &To Clipboard"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   3255
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "frmKeySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub cmdCopy_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText Me.txtKey.Text
End Sub

Private Sub Form_Activate()
Dim k As Integer
Dim j As Integer
Dim pMax As Integer
Dim tmp As String
Dim pin As String
Dim strLine As String
Dim strKeyText As String

Me.Caption = "BC-52 Key - " & CutFilePath(gstrkeyFile)

'
strLine = "----------------------------------" & vbCrLf

tmp = strLine & "        BC-52 KEY SETTINGS " & vbCrLf & strLine

For k = 1 To 6
    tmp = tmp & Format(Wmax(k), "00") & " "
Next

tmp = tmp & " NR  1 2 3 4 5 6" & vbCrLf & strLine

'get max pins
For k = 1 To 6
If Wmax(k) > pMax Then pMax = Wmax(k)
Next

For k = 1 To pMax

    For j = 1 To 6
    'read pins
    If k <= Wmax(j) Then
        pin = Mid(Wpin(j), k, 1)
        If pin = "1" Then
            tmp = tmp & Format(k, "00") & " "
            Else
            tmp = tmp & "-- "
        End If
        Else
        tmp = tmp & "   "
    End If
    Next

    tmp = tmp & " " & Format(k, "00") & "  "
    
    'read lugs
    If k < 33 Then
        For j = 1 To 6
        pin = Mid(BLug(k), j, 1)
        If pin = "1" Then
            tmp = tmp & Trim(Val(j)) & " "
            Else
            tmp = tmp & "- "
            End If
        Next
    End If
    
    tmp = tmp & vbCrLf

Next

tmp = tmp & strLine

Me.txtKey.Text = tmp
End Sub

