Attribute VB_Name = "modMain"

'-----------------------------------------------------------
'
'       Hagelin BC-52 Cipher Machine Simulator v3.6
'
'                Written by D. Rijmenants
'
'                       (c) 2006
'
'-----------------------------------------------------------


Option Explicit

Public Wsel(6) As Integer
Public Wpos(6) As Integer
Public Wpin(6) As String
Public Wmax(6) As Integer
Public W_len(47) As Integer
Public BLug(32) As String
Public Wlab(12, 47) As String

Public gintLabelView(12) As Integer
Public gintBarStepping(6) As Integer
Public gstrCipherPrintWheel As String
Public W_textLabel(12) As String
Public gblnModeCipher As Boolean
Public gstrSpaceLetter As String
Public gintPosMemo(6) As Integer
Public gintAdvanceBar(32) As Integer
Public gintLastDialView As Integer
Public gintWheelOffset As Integer
Public gstrMachineSetup As String
Public gblnCancelSave As Boolean
'
Public gstrExitVal As String
'
Public gstrInput As String
Public gstrOutput As String
Public gintAlphaWheel As Integer
Public glngGroupCount As Long
Public gstrCounter As Integer
Public gstrAutoType As Boolean
Public gblnSound As Boolean
Public gstrkeyFile As String
Public gblnKeyHasChanged As Boolean
Public gblnFastRun As Boolean

Public DialX As Variant
Public DialY As Variant

Public gblnHoldPins As Boolean
Public gblnCipherBars As Boolean
Public gblnCxType As Boolean
Public Const DefaultSetup = "1616181921232325252628280102030405010X11111"
'                            010203040506070809101112

'cursor functions to move forms
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public LastPoint As POINTAPI
Public iTPPY As Long
Public iTPPX As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type

'sound api
Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
        (ByVal lpszSoundName As Any, ByVal uFlags As Long) As Long
Public Const SND_ASYNC = &H1     ' Play asynchronously
Public Const SND_NODEFAULT = &H2 ' Don't use default sound
Public Const SND_MEMORY = &H4    ' lpszSoundName points to a memory file
Public SoundBuffer As String

'time function
Public Declare Function GetTickCount Lib "kernel32" () As Long


Sub Main()
Dim k As Integer
Dim j As Integer
Dim tmp As String
iTPPX& = Screen.TwipsPerPixelX
iTPPY& = Screen.TwipsPerPixelY
gblnSound = True

DialX = Array(4500, 4620, 4740, 4830, 4890, 4920, 4950, 4920, 4890, 4860, 4785, 4665, 4500, 4320, 4185, 4080, 3990, 3915, 3885, 3870, 3870, 3915, 3990, 4095, 4200, 4350)
DialY = Array(3825, 3870, 3930, 4005, 4095, 4185, 4305, 4395, 4470, 4545, 4605, 4680, 4725, 4725, 4695, 4635, 4575, 4500, 4410, 4290, 4185, 4080, 3990, 3915, 3855, 3825)

'define rotor labels
'                 12345678901234567890123456789012345678901234567
W_textLabel(1) = "ABCDEFGHIJKLMNOPQRSTUVXYZ" '25**
W_textLabel(2) = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" '26
W_textLabel(3) = "ABCDEFGHI.JKLMNOPQR.STUVWXYZ." '29
W_textLabel(4) = "ABCDE.FGHIJ.KLMNOP.QRSTU.VWXYZ." '31
W_textLabel(5) = "A.B.C.D.E.F.G.H.I.J.K.L.M.N.O.P.Q." '34**
W_textLabel(6) = "ABC.DE.FG.HIJ.KL.MN.OPQ.RS.TU.VWX.YZ." '37
W_textLabel(7) = "A.B.C.D.E.F.G.H.I.J.K.L.M.N.O.P.Q.R.S." '38**
W_textLabel(8) = "AB.CD.EF.G.HI.JK.LM.N.OP.QR.ST.U.VW.XY.Z." '41
W_textLabel(9) = "A.B.C.D.E.F.G.H.I.J.K.L.M.N.O.P.Q.R.S.T.U." '42**
'                  12345678901234567890123456789012345678901234567
W_textLabel(10) = "AB.C.DE.F.GH.I.JK.L.MN.O.PQ.R.ST.U.VW.X.YZ." '43
W_textLabel(11) = "A.B.C.D.E.F.G.H.I.J.K.L.M.N.O.P.Q.R.S.T.U.V.X." '46**
W_textLabel(12) = "AB.C.D.E.FG.H.I.J.KL.M.N.O.P.QR.S.T.U.VW.X.Y.Z." '47

For k = 1 To 12
    W_len(Len(W_textLabel(k))) = k
Next

'read setup
gstrMachineSetup = GetSetting(App.EXEName, "config", "setup", DefaultSetup)
If Len(gstrMachineSetup) <> 43 Then
    gstrMachineSetup = DefaultSetup
    SaveSetting App.EXEName, "config", "setup", DefaultSetup
    End If
If Asc(Mid(gstrMachineSetup, 38, 1)) < 65 Or Asc(Mid(gstrMachineSetup, 38, 1)) > 90 Then
    gstrMachineSetup = DefaultSetup
    SaveSetting App.EXEName, "config", "setup", DefaultSetup
    End If
DoReload:
'check labels
For k = 1 To 12
    gintLabelView(k) = Val(Mid(gstrMachineSetup, (k * 2) - 1, 2))
    If gintLabelView(k) < 1 Or gintLabelView(k) > Len(W_textLabel(k)) Then
        'error setup
        gstrMachineSetup = DefaultSetup
        SaveSetting App.EXEName, "config", "setup", DefaultSetup
        GoTo DoReload
    End If
Next
'check bar advancing
For k = 13 To 17
    j = Val(Mid(gstrMachineSetup, (k * 2) - 1, 2))
    If j < 1 Or j > 32 Then
        'error setup
        tmp = DefaultSetup
        SaveSetting App.EXEName, "config", "setup", DefaultSetup
        GoTo DoReload
        Else
    gintAdvanceBar(j) = k - 11
    End If
Next

If Mid(gstrMachineSetup, 35, 1) = "1" Then gblnCipherBars = True
If Mid(gstrMachineSetup, 36, 1) = "1" Then gblnHoldPins = True
If Mid(gstrMachineSetup, 37, 1) = "1" Then gblnCxType = True

gstrSpaceLetter = Mid(gstrMachineSetup, 38, 1)

For k = 39 To 43
    gintBarStepping(k - 37) = Val(Mid(gstrMachineSetup, k, 1))
Next

Load frmMain
Load frmClipBoard
Load frmQuick
Load frmInfo
Load frmKeySheet
Load frmKey
Load frmCustom

App.HelpFile = App.Path & "\BC-52.hlp"

frmMain.shpDot.FillColor = RGB(166, 166, 166)
frmMain.shpDot.BorderColor = RGB(166, 166, 166)
frmMain.shpDot.Left = DialX(0)
frmMain.shpDot.Top = DialY(0)

gstrCipherPrintWheel = "ZYXWVUTSRQPONMLKJIHGFEDCBA"
gblnModeCipher = True
gintLastDialView = 1
frmMain.Show

'get key file
gstrkeyFile = GetSetting(App.EXEName, "config", "key", "")

If gstrkeyFile = "" Then
    gstrkeyFile = App.Path & "\Example Key.C52"
    MsgBox "Example key settings are loaded!", vbExclamation
    End If

Call LoadKeySettings(gstrkeyFile)

tmp = GetSetting(App.EXEName, "config", "fastrun", "0")
If tmp = "0" Then
    gblnFastRun = False
    Else
    gblnFastRun = True
End If

End Sub

Public Sub SetWheelView(wheel As Integer)
'set wheels view
Dim k As Integer
Dim j As Integer
Dim P As Integer
Dim Lpos As Integer
'calculatte labeling offset to pin
Lpos = Wpos(wheel) + (gintLabelView(Wsel(wheel)) - 1)
If Lpos > Wmax(wheel) Then Lpos = Lpos - Wmax(wheel)
frmMain.lblWindow(wheel) = ""
For j = Lpos + 2 To Lpos - 2 Step -1
    If j < 1 Then
        P = j + Wmax(wheel)
    ElseIf j > Wmax(wheel) Then
        P = j - Wmax(wheel)
    Else
        P = j
    End If
    frmMain.lblWindow(wheel) = frmMain.lblWindow(wheel) & Wlab(wheel, P) & vbCrLf
Next j

End Sub



'-----------------------------key ------------------------------------

Public Sub LoadKeySettings(ByVal aFile As String)
'load key settings from "C52.dat" file
Dim infile As Integer
Dim FileName As String
Dim tmpInput(40) As String
Dim k As Integer
Dim j As Integer
Dim tmp As String

On Error GoTo errHandle
If aFile = "" Then GoTo errHandle

infile = FreeFile
FileName = aFile
If Dir(FileName) = "" Then
    'no file
    GoTo errHandle
    End If
Open FileName For Input As infile

Line Input #infile, tmp
If tmp <> "BC52SIM" Then
    MsgBox "Failed loading key. Unknown data format", vbCritical
    GoTo errHandle
    End If

For k = 1 To 38
    Line Input #infile, tmpInput(k)
Next k

Close infile
'load pins
For k = 1 To 6
    Wpin(k) = tmpInput(k)
    Wmax(k) = Len(tmpInput(k))
    Wsel(k) = W_len(Len(tmpInput(k)))
Next

If Wmax(1) = 47 And Wmax(2) = 47 Then
    'CX-52 key settings
    If gblnCxType = False Then
        MsgBox "The loaded key settings are for a CX-52 model only." & vbCrLf & vbCrLf & "Machine settings are changed to the CX-52 model.", vbInformation
        'change setup
        Mid(gstrMachineSetup, 37, 1) = "1"
        SaveSetting App.EXEName, "config", "setup", gstrMachineSetup
        End If
    gblnCxType = True
    Else
    'C-52 key settings
    If gblnCxType = True Then
        MsgBox "The loaded key settings are for a C-52 model only." & vbCrLf & vbCrLf & "Machine settings are changed to the C-52 model.", vbInformation
        'change setup
        Mid(gstrMachineSetup, 37, 1) = "0"
        SaveSetting App.EXEName, "config", "setup", gstrMachineSetup
        End If
    gblnCxType = False
    End If

'load and check lugs
For k = 1 To 32
    If Len(tmpInput(k + 6)) <> 6 Then GoTo errHandle
    BLug(k) = tmpInput(k + 6)
Next

'create labels
Call CreateLabels
Call resetAllWheels
Call MemorizeWheels
gblnKeyHasChanged = False
gstrkeyFile = aFile

Exit Sub
errHandle:
Close infile

'load default settings

If gblnCxType = True Then
    'set 6 wheels with 47 pins
    Call SetCX52wheels
    MsgBox "Default key settings for CX-52 model are loaded!" & vbCrLf & vbCrLf & "Please set Pins and Lugs before using the CX-52", vbCritical + vbOKOnly, " BC-52"
    Else
    'set 6 different wheels
    Call SetC52wheels
    MsgBox "Default key settings for C-52 model are loaded!" & vbCrLf & vbCrLf & "Please set Pins and Lugs before using the C-52", vbCritical + vbOKOnly, " BC-52"
End If

End Sub

Public Sub SaveKeySettings(ByVal aFile As String)
'save current key settings to file
Dim infile As Integer
Dim FileName As String
Dim tmp As String
Dim k As Integer

On Error GoTo errHandle

infile = FreeFile

Open aFile For Output As infile

'header
tmp = "BC52SIM"
Print #infile, tmp

'save pins
For k = 1 To 6
    tmp = Wpin(k)
    Print #infile, tmp
Next k
'save lugs
For k = 1 To 32
    tmp = BLug(k)
    Print #infile, tmp
Next


Close infile
gblnKeyHasChanged = False
gstrkeyFile = aFile
Exit Sub

errHandle:
Close infile
gblnCancelSave = True
MsgBox "Failed saving the key settings." & vbCrLf & vbCrLf & Err.Description, vbCritical + vbOKOnly
End Sub

Public Sub EraseStettings()
' erase saved settings (kill "atomix.dat" file)
Dim FileName As String
Dim infile As Integer
On Error GoTo errHandle
infile = FreeFile
FileName = App.Path & "\atomix.dat"
Kill FileName
Exit Sub
errHandle:
MsgBox "Failed deleting the Key settings." & vbCrLf & vbCrLf & Err.Description, vbCritical + vbOKOnly
End Sub

Public Sub CreateLabels()
'create wheel labels array for current wheels
Dim k As Integer
Dim j As Integer
For k = 1 To 6
    For j = 1 To Wmax(k)
        If Mid(W_textLabel(Wsel(k)), j, 1) <> "." Then
            Wlab(k, j) = Mid(W_textLabel(Wsel(k)), j, 1)
            Else
            Wlab(k, j) = Format(j, "00")
            End If
    Next j
Next
End Sub

Public Sub EndProgram()
' exit program
Dim k As Integer
Dim retval As Integer

If gblnKeyHasChanged = True Then
    retval = MsgBox("The BC-52 key settings are changed." & vbCrLf & vbCrLf & "Do you want to save these changes?", vbQuestion + vbYesNoCancel)
    If retval = vbCancel Then Exit Sub
    If retval = vbYes Then
        gblnCancelSave = False
        Call SaveFile
        If gblnCancelSave = True Then Exit Sub
        End If
    End If
    
'save current key
If gstrkeyFile <> "Untitled" Then SaveSetting App.EXEName, "config", "key", gstrkeyFile
SaveSetting App.EXEName, "config", "fastrun", gblnFastRun

Unload frmMain
Unload frmInfo
Unload frmClipBoard
Unload frmQuick
Unload frmKeySheet
Unload frmKey
Unload frmCustom

End
End Sub

Public Function ReadPin(ByVal wheel As Integer) As Integer
'read current pin of a wheel
Dim PinPos As Integer
If Mid(Wpin(wheel), Wpos(wheel), 1) = "1" Then
    ReadPin = 1
    Else
    ReadPin = 0
    End If
End Function

Public Sub Crypto(akey As Integer)
Dim KeyOffset As Integer
Dim KeyOut As Integer
Dim CharOut As String
Dim currPin(6) As Integer
Dim k As Integer
Dim j As Integer
Dim i As Integer
Dim barHasSlided As Boolean
'memorize pins
For k = 1 To 6
    currPin(k) = ReadPin(k)
Next

KeyOffset = 0

'turn drum

For k = 1 To 32
    i = gintAdvanceBar(k)
    barHasSlided = False
    For j = 1 To 6
        'if No pin hold, read pins settings each time
        If gblnHoldPins = False Then currPin(j) = ReadPin(j)
        'check for sliding bars
        If Mid(BLug(k), j, 1) = "1" And currPin(j) = 1 Then
            'pin and lug:
            barHasSlided = True
            If i = 0 Then
                'only cipher bar
                KeyOffset = KeyOffset + 1
            Else
                'advance bar:
                If gintBarStepping(i) = 1 Or gintBarStepping(i) = 3 Then Call MoveWheel(i)
                'if gintAdvanceBar is also cipher bar +1:
                If gblnCipherBars = True Then KeyOffset = KeyOffset + 1
            End If
            Exit For
        End If
    Next
If barHasSlided = False Then
    If i <> 0 Then
        If gintBarStepping(i) = 2 Or gintBarStepping(i) = 3 Then Call MoveWheel(i)
    End If
End If
If gstrAutoType = False Then PauzeTime (10)
SetDialView (akey + KeyOffset)
Next k

Call MoveWheel(1)

KeyOut = akey + KeyOffset + gintWheelOffset
DoStrip:
If KeyOut > 26 Then KeyOut = KeyOut - 26: GoTo DoStrip

'print plain and cipher output
SetDialView (KeyOut)
CharOut = Mid(gstrCipherPrintWheel, KeyOut, 1)
If gblnModeCipher = False And CharOut = gstrSpaceLetter Then CharOut = " "
gstrInput = gstrInput & Chr(akey + 64)
gstrOutput = gstrOutput & CharOut

With frmMain
.lblInput.Caption = .lblInput.Caption & Chr(akey + 64)
.lblOutput.Caption = .lblOutput.Caption & CharOut

If gblnModeCipher = True Then
    'check for groups
    glngGroupCount = glngGroupCount + 1
    If glngGroupCount = 5 Then
        glngGroupCount = 0
        .lblInput.Caption = .lblInput.Caption & " "
        .lblOutput.Caption = .lblOutput.Caption & " "
    End If
End If

'check for maximum display length
If Len(.lblInput.Caption) > 34 Then
    .lblInput.Caption = Right(.lblInput.Caption, 34)
    .lblOutput.Caption = Right(.lblOutput.Caption, 34)
    Else
End If

'set counter
gstrCounter = gstrCounter + 1
If gstrCounter > 999 Then gstrCounter = 0
.lblCounter.Caption = Format(gstrCounter, "000")

End With
End Sub

Public Sub SetDialView(ByVal aLetter As Integer)
Dim sq As Single
DoStrip:
If aLetter > 26 Then aLetter = aLetter - 26: GoTo DoStrip
frmMain.shpDot.Left = DialX(aLetter - 1)
frmMain.shpDot.Top = DialY(aLetter - 1)
gintLastDialView = aLetter
End Sub

Public Sub MoveWheel(wheel As Integer)
'move a wheel
Wpos(wheel) = Wpos(wheel) + 1
If Wpos(wheel) > Wmax(wheel) Then Wpos(wheel) = Wpos(wheel) - Wmax(wheel)
Call SetWheelView(wheel)
End Sub

Public Sub resetAllWheels()
Dim k As Integer
Dim i As Integer
For k = 1 To 6
    Wpos(k) = 1 - (gintLabelView(Wsel(k)) - 1)
    If Wpos(k) < 1 Then Wpos(k) = Wpos(k) + Wmax(k)
    Call SetWheelView(k)
Next
End Sub

Public Sub AutoTyping()
'autotyping function
Dim tmpQuick As String
Dim tmp As Integer
Dim i As Long
Dim tm As Long
'delet all but alphabet
tmpQuick = frmQuick.txtQuick.Text
If tmpQuick = "" Then Exit Sub
gstrAutoType = True
'select autotyping speed
Select Case frmQuick.cmbSpeed.Text
Case "Slow"
    tm = 2000
Case "Normal"
    tm = 250
Case "Fast"
    tm = 0
End Select
For i = 1 To Len(tmpQuick)
    tmp = Asc(UCase(Mid(tmpQuick, i, 1)))
        If (tmp > 64 And tmp < 91) Or tmp = 32 Then
            If gblnModeCipher = True And tmp = 32 Then tmp = Asc(gstrSpaceLetter)
            If tmp <> 32 Then
                PauzeTime (tm)
                If frmQuick.cmbSpeed.Text <> "Fast" Then PlaySound (1)
                Call Crypto(tmp - 64)
            End If
        End If
    DoEvents
    If gstrAutoType = False Then
        MsgBox "Auto Typing aborted.", vbInformation, " BC-52"
        Exit For
    End If
Next i
gstrAutoType = False
End Sub

Public Sub PlaySound(aSound As Integer)
'play sound
Dim Ret
Select Case aSound
Case 0
    Exit Sub
Case 1
    If gblnSound = False Then Exit Sub
    SoundBuffer = StrConv(LoadResData(1, "Sounds"), vbUnicode)
Case 2
    If gblnSound = False Then Exit Sub
    SoundBuffer = StrConv(LoadResData(2, "Sounds"), vbUnicode)
Case 3
    SoundBuffer = StrConv(LoadResData(3, "Sounds"), vbUnicode)
End Select
Ret = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
End Sub

Public Sub PauzeTime(TimeToWait As Long)
'make a pauze
If gblnFastRun = True Then Exit Sub
Dim currTime As Long
Dim passedTime As Long
currTime = GetTickCount()
Do
    passedTime = Abs(GetTickCount() - currTime)
Loop While passedTime < TimeToWait
End Sub

Public Sub MemorizeWheels()
Dim k As Integer
For k = 1 To 6
    gintPosMemo(k) = Wpos(k)
Next
End Sub

Public Sub SetCX52wheels()
Dim k As Integer
For k = 1 To 6
    Wpin(k) = String(47, "0")
Next
For k = 1 To 6
    Wmax(k) = Len(Wpin(k))
    Wsel(k) = W_len(Len(Wpin(k)))
    Wpos(k) = 1
    gintPosMemo(k) = 1
Next

Call CreateLabels
Call resetAllWheels
Call MemorizeWheels

For k = 1 To 6
    Call SetWheelView(k)
Next k

'lugs
For k = 0 To 32
    BLug(k) = "000000"
Next

gstrkeyFile = "Untitled"
gblnKeyHasChanged = True
MsgBox "CX-52 model selected and all pins and lugs cleared." & vbCrLf & vbCrLf & "Please set pins and lugs before using the CX-52!", vbExclamation

End Sub

Public Sub SetC52wheels()
Dim k As Integer
'set 6 different wheels
Wpin(1) = "00000000000000000000000000000" '29
Wpin(2) = "0000000000000000000000000000000" '31
Wpin(3) = "0000000000000000000000000000000000000" '37
Wpin(4) = "00000000000000000000000000000000000000000" '41
Wpin(5) = "0000000000000000000000000000000000000000000" '43
Wpin(6) = "00000000000000000000000000000000000000000000000" '47
For k = 1 To 6
    Wmax(k) = Len(Wpin(k))
    Wsel(k) = W_len(Len(Wpin(k)))
    Wpos(k) = 1
    gintPosMemo(k) = 1
Next

Call CreateLabels
Call resetAllWheels
Call MemorizeWheels

For k = 1 To 6
    Call SetWheelView(k)
Next k

'lugs
For k = 0 To 32
    BLug(k) = "000000"
Next

gstrkeyFile = "Untitled"
gblnKeyHasChanged = True
MsgBox "C-52 model selected and all pins and lugs cleared." & vbCrLf & vbCrLf & "Please set pins and lugs before using the C-52!", vbExclamation

End Sub

Public Sub OpenFile()
Dim tmpfilename As String
Dim retval As Integer

If gblnKeyHasChanged = True Then
    retval = MsgBox("The BC-52 key settings are changed." & vbCrLf & vbCrLf & "Do you want to save these changes?", vbQuestion + vbYesNoCancel)
    If retval = vbCancel Then Exit Sub
    If retval = vbYes Then
        Call SaveFile
        If gblnCancelSave = True Then Exit Sub
        End If
    End If

On Error Resume Next
With frmMain.Dialog2
    .FileName = ""
    .DialogTitle = " Load BC-52 Key..."
    .Flags = &H1000 Or &H4
    .DefaultExt = ".C52"
    .InitDir = gstrkeyFile
    .Filter = "BC-52 Key Files (*.C52)|*.C52"
    .FilterIndex = 1
    .ShowOpen
    If Err = 32755 Or .FileName = "" Then Exit Sub
    tmpfilename = .FileName
End With

Call LoadKeySettings(tmpfilename)

End Sub

Public Sub SaveFile()
Dim tmpfilename As String
On Error Resume Next
With frmMain.Dialog2
    .FileName = CutFilePath(gstrkeyFile)
    .DialogTitle = " Save BC-52 Key..."
    .Flags = &H4 Or &H2
    .DefaultExt = ".C52"
    .InitDir = gstrkeyFile
    .Filter = "BC-52 Key Files (*.C52)|*.C52"
    .FilterIndex = 1
    .ShowSave
    If Err = 32755 Or .FileName = "" Then gblnCancelSave = True: Exit Sub
    tmpfilename = .FileName
End With

Call SaveKeySettings(tmpfilename)

End Sub

Public Function CutFilePath(strFile As String) As String
'returns only the filename without full path
Dim k As Integer
Dim pos As Integer
For k = 1 To Len(strFile)
    If Mid(strFile, k, 1) = "\" Then pos = k
Next
If pos = 0 Then
    CutFilePath = strFile
    Else
    CutFilePath = Mid(strFile, pos + 1)
    End If
End Function

