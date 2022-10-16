VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "frmMain"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   5355
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame FrameLetterCase 
      Caption         =   "FrameLetterCase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   180
      TabIndex        =   13
      Top             =   3060
      Width           =   5010
      Begin VB.OptionButton OptionCase 
         Caption         =   "OptionCase"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Top             =   720
         Width           =   3030
      End
      Begin VB.OptionButton OptionCase 
         Caption         =   "OptionCase"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Width           =   3030
      End
   End
   Begin VB.Frame FrameMACFormat 
      Caption         =   "FrameMACFormat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   135
      TabIndex        =   6
      Top             =   1260
      Width           =   5100
      Begin VB.OptionButton OptionMACFormat 
         Caption         =   "OptionMACFormat"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   225
         TabIndex        =   12
         Top             =   1035
         Width           =   4605
      End
      Begin VB.OptionButton OptionMACFormat 
         Caption         =   "OptionMACFormat"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   225
         TabIndex        =   8
         Top             =   675
         Width           =   4605
      End
      Begin VB.OptionButton OptionMACFormat 
         Caption         =   "OptionMACFormat"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Top             =   360
         Width           =   3030
      End
   End
   Begin VB.ComboBox NoOfPing 
      Appearance      =   0  '2D
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "frmMain.frx":0442
      Left            =   180
      List            =   "frmMain.frx":0444
      Style           =   2  'Dropdown-Liste
      TabIndex        =   4
      Top             =   630
      Width           =   2445
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "cmdAbort"
      Height          =   375
      Left            =   4050
      TabIndex        =   10
      Top             =   675
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "cmdSearch"
      Height          =   375
      Left            =   4005
      TabIndex        =   9
      Top             =   180
      Width           =   1095
   End
   Begin VB.TextBox txtIPAdress 
      Height          =   285
      Index           =   4
      Left            =   2880
      TabIndex        =   5
      Text            =   "txtIPAdress"
      Top             =   585
      Width           =   420
   End
   Begin VB.TextBox txtIPAdress 
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   3
      Text            =   "txtIPAdress"
      Top             =   270
      Width           =   420
   End
   Begin VB.TextBox txtIPAdress 
      Height          =   285
      Index           =   2
      Left            =   2025
      TabIndex        =   2
      Text            =   "txtIPAdress"
      Top             =   270
      Width           =   420
   End
   Begin VB.TextBox txtIPAdress 
      Height          =   285
      Index           =   1
      Left            =   1215
      TabIndex        =   1
      Text            =   "txtIPAdress"
      Top             =   225
      Width           =   420
   End
   Begin VB.TextBox txtIPAdress 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Text            =   "txtIPAdress"
      Top             =   180
      Width           =   690
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblInfo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2250
      TabIndex        =   11
      Top             =   4635
      Width           =   540
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Type STARTUPINFO
   cb              As Long
   lpReserved      As String
   lpDesktop       As String
   lpTitle         As String
   dwX             As Long
   dwY             As Long
   dwXSize         As Long
   dwYSize         As Long
   dwXCountChars   As Long
   dwYCountChars   As Long
   dwFillAttribute As Long
   dwFlags         As Long
   wShowWindow     As Integer
   cbReserved2     As Integer
   lpReserved2     As Byte
   hStdInput       As Long
   hStdOutput      As Long
   hStdError       As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess    As Long
    hThread     As Long
    dwProcessId As Long
    dwThreadId  As Long
End Type

Const STILL_ACTIVE = &H103

Dim Abort As Boolean

Private Sub Form_Load()

    Dim i As Long
    
    For i = 0 To txtIPAdress.Count - 1
        txtIPAdress(i).Text = GetSetting(App.EXEName, "IPAdress", "No" & CStr(i))
        txtIPAdress(i).ToolTipText = "IP-Adresse 0-255"
        txtIPAdress(i).Alignment = vbCenter
        If i > 0 Then
            txtIPAdress(i).Width = txtIPAdress(i - 1).Width
            txtIPAdress(i).Height = txtIPAdress(i - 1).Height
            txtIPAdress(i).FontSize = txtIPAdress(i - 1).FontSize

            If i < txtIPAdress.Count - 1 Then
                txtIPAdress(i).Left = txtIPAdress(i - 1).Left + txtIPAdress(i - 1).Width + txtIPAdress(i - 1).Width / 2
                txtIPAdress(i).Top = txtIPAdress(i - 1).Top

            Else
                txtIPAdress(i).Left = txtIPAdress(i - 1).Left
                txtIPAdress(i).Top = txtIPAdress(i - 1).Top + txtIPAdress(i - 1).Height + txtIPAdress(i - 1).Height / 2
            
            End If
            
        End If
    
    Next i
    
    i = txtIPAdress.Count - 1
    i = txtIPAdress(i).Top + txtIPAdress(i).Height / 2 - NoOfPing.Height / 2
    NoOfPing.Left = txtIPAdress(0).Left
    NoOfPing.Top = i
    i = 2
    i = txtIPAdress(i).Left + txtIPAdress(i).Width - txtIPAdress(0).Left
    NoOfPing.Width = i
    
    NoOfPing.Clear
    NoOfPing.AddItem "kein Pingtest"
    NoOfPing.Tag = " Pingtest in Folge"
    For i = 1 To 9
        NoOfPing.AddItem i & NoOfPing.Tag
        
    Next i
    NoOfPing.ListIndex = Val(GetSetting(App.EXEName, "Setup", "NoOfPing", 0))
    
    cmdSearch.Height = txtIPAdress(0).Height
    cmdSearch.Top = txtIPAdress(0).Top
    
    cmdSearch.Left = txtIPAdress(txtIPAdress.Count - 1).Left + txtIPAdress(txtIPAdress.Count - 1).Width + txtIPAdress(txtIPAdress.Count - 1).Width / 2
    
    cmdAbort.Height = txtIPAdress(txtIPAdress.Count - 1).Height
    cmdAbort.Top = txtIPAdress(txtIPAdress.Count - 1).Top
    cmdAbort.Left = cmdSearch.Left
    
    FrameMACFormat.Left = txtIPAdress(0).Left
    FrameMACFormat.Top = cmdAbort.Top + cmdAbort.Height + cmdAbort.Height / 2
    FrameMACFormat.Caption = " Ausgabeformat MAC-Adressen "
    
    OptionMACFormat(0).Left = OptionMACFormat(0).Top
    OptionMACFormat(1).Left = OptionMACFormat(0).Left
    OptionMACFormat(1).Top = OptionMACFormat(0).Top + 1.3 * OptionMACFormat(0).Height
    OptionMACFormat(2).Left = OptionMACFormat(1).Left
    OptionMACFormat(2).Top = OptionMACFormat(1).Top + 1.3 * OptionMACFormat(1).Height
    
    OptionMACFormat(0).Caption = "00-00-00-00-00-00"
    OptionMACFormat(1).Caption = "00:00:00:00:00:00"
    OptionMACFormat(2).Caption = "0000.0000.0000"
    
    OptionMACFormat(0).Value = IIf(Val(GetSetting(App.EXEName, "MACAdress", "Format")) = 0, True, False)
    OptionMACFormat(1).Value = IIf(Val(GetSetting(App.EXEName, "MACAdress", "Format")) = 1, True, False)
    OptionMACFormat(2).Value = IIf(Val(GetSetting(App.EXEName, "MACAdress", "Format")) = 2, True, False)
    
    FrameMACFormat.Height = OptionMACFormat(0).Top + OptionMACFormat(2).Top + OptionMACFormat(2).Height
    
    FrameLetterCase.Left = FrameMACFormat.Left
    FrameLetterCase.Top = FrameMACFormat.Top + FrameMACFormat.Height + cmdAbort.Height / 2
    FrameLetterCase.Caption = " Ausgabe mit "
    
    OptionCase(0).Left = OptionMACFormat(0).Left
    OptionCase(1).Left = OptionCase(0).Left
    OptionCase(0).Top = OptionMACFormat(0).Top
    OptionCase(1).Top = OptionCase(0).Top + 1.3 * OptionCase(0).Height
    
    OptionCase(0).Caption = UCase("Grossbuchstaben")
    OptionCase(1).Caption = LCase("Kleinbuchstaben")
    
    OptionCase(0).Value = IIf(Val(GetSetting(App.EXEName, "MACAdress", "Letter")) = 0, True, False)
    OptionCase(1).Value = IIf(Val(GetSetting(App.EXEName, "MACAdress", "Letter")) = 1, True, False)
    
    FrameLetterCase.Height = OptionCase(0).Top + OptionCase(1).Top + OptionCase(1).Height
    
    Me.Caption = "MAC-Adressen ermitteln - V" & App.Major & "." & App.Minor & "." & App.Revision
    
    cmdSearch.Caption = "Starten"
    cmdSearch.ToolTipText = App.CompanyName
    cmdSearch.Enabled = True
    
    cmdAbort.Caption = "Abbrechen"
    cmdAbort.ToolTipText = App.CompanyName
    cmdAbort.Enabled = False
    
    lblInfo = ""
    lblInfo.ToolTipText = App.CompanyName
    
    lblInfo.Top = FrameLetterCase.Top + FrameLetterCase.Height + cmdAbort.Height / 2

    Me.Width = cmdSearch.Left + cmdSearch.Width + txtIPAdress(0).Left + (Me.Width - Me.ScaleWidth)
    Me.Height = (Me.Height - Me.ScaleHeight) + lblInfo.Top + lblInfo.Height + lblInfo.Height / 2
    
    lblInfo.Left = Me.ScaleWidth / 2 - lblInfo.Width / 2
    
    FrameMACFormat.Width = Me.ScaleWidth - 2 * FrameMACFormat.Left
    FrameLetterCase.Width = FrameMACFormat.Width

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call SaveSetup

End Sub

Private Sub SaveSetup()

    Dim i As Long
    
    For i = 0 To txtIPAdress.Count - 1
        SaveSetting App.EXEName, "IPAdress", "No" & CStr(i), txtIPAdress(i).Text
    
    Next i

    SaveSetting App.EXEName, "Setup", "NoOfPing", NoOfPing.ListIndex

    i = 0
    i = IIf(OptionMACFormat(1).Value, 1, i)
    i = IIf(OptionMACFormat(2).Value, 2, i)
    
    SaveSetting App.EXEName, "MACAdress", "Format", i

    i = 0
    i = IIf(OptionCase(1).Value, 1, i)

    SaveSetting App.EXEName, "MACAdress", "Letter", i

End Sub

Private Sub cmdSearch_Click()

    Dim i As Long, j As Long

    cmdSearch.Enabled = False
    cmdAbort.Enabled = True
    Abort = False

    For i = 0 To txtIPAdress.Count - 1
        txtIPAdress(i).Enabled = False
    
    Next i
    NoOfPing.Enabled = False
    
    FrameMACFormat.Enabled = False
    OptionMACFormat(0).Enabled = False
    OptionMACFormat(1).Enabled = False
    OptionMACFormat(2).Enabled = False
    
    FrameLetterCase.Enabled = False
    OptionCase(0).Enabled = False
    OptionCase(1).Enabled = False
    
    Dim StartIP As Long
    Dim EndIP As Long
    Dim IPAdress As String, IPFile As String
    Dim sMACAdress As String, sHostName As String
    
    Dim lHandle1 As Long, lHandle2 As Long
    Dim bError As Boolean
    Dim cmdline As String
    
    Dim KeyRetVal As Long
    
    StartIP = Val(Trim(txtIPAdress(3).Text))
    EndIP = Val(Trim(txtIPAdress(4).Text))
    
    KeyRetVal = vbOKOnly
    
    IPFile = App.Path & "\" & App.EXEName & ".CSV"
    lHandle1 = FreeFile
    Open IPFile For Output As #lHandle1
        IPFile = Date & ";" & Time & ";;"
        Print #lHandle1, IPFile
        Print #lHandle1,
        Print #lHandle1, "Name;IP-Adresse;Physikalische Adresse;"
    
        For i = StartIP To EndIP
            IPAdress = Val(Trim(txtIPAdress(0).Text)) & "." & _
                        Val(Trim(txtIPAdress(1).Text)) & "." & _
                        Val(Trim(txtIPAdress(2).Text)) & "." & _
                        Val(Trim(Str(i)))
    
            lblInfo.Caption = "Ping zu " & IPAdress & " aktiv, bitte warten..."
            lblInfo.Refresh
            
            For j = 1 To NoOfPing.ListIndex
                If Not PingA(IPAdress) Then
                    lblInfo.Caption = j & "." & Trim(NoOfPing.Tag) & " zu " & IPAdress & " aktiv, bitte warten..."
                    lblInfo.Refresh
                
                    DoEvents
                    
                Else
                    Exit For
                    
                End If
            
            Next j
            
            If PingA(IPAdress) > 0 Then
                lblInfo.Caption = "Info's werden von " & IPAdress & " abgefragt, bitte warten..."
                lblInfo.Refresh
            
                DoEvents
                
                sMACAdress = Trim(GetMacAdresse(IPAdress))
                sHostName = Trim(HostByAddress(IPAdress))
                If Len(sMACAdress) Then
                
                    sMACAdress = MACFormat(sMACAdress)
                    IPFile = sHostName & ";" & IPAdress & ";" & sMACAdress & ";"
                    Print #lHandle1, IPFile
                
                Else
                
                    IPFile = sHostName & ";" & IPAdress & ";" & ";"
                    
                    If Len(sHostName) Then
                        IPFile = sHostName & "_" & IPAdress
            
                    Else
                        IPFile = Replace(IPAdress, ".", "_")
            
                    End If
    
                    cmdline = IPFile & ".TXT"
            
                    IPFile = App.Path & "\" & App.EXEName & ".BAT"
                    lHandle2 = FreeFile
                    Open IPFile For Output As #lHandle2
                        Print #lHandle2, "@echo off"
                        Print #lHandle2, "arp -a "; IPAdress; ">"; cmdline
            
                    Close #lHandle2
      
                    sMACAdress = ""
                    bError = ShellGetHandle(IPFile, sMACAdress, lHandle2)
   
                    If bError Then
                        MsgBox IPFile & " kann nicht gestartet werden", vbCritical + vbOKOnly, Me.Caption
            
                    Else
                        
                        Do While Wait(lHandle2)
                            DoEvents
      
                        Loop
                        
                        Kill IPFile
                    
                        IPFile = cmdline
                        lHandle2 = FreeFile: bError = False
                        Open IPFile For Input As #lHandle2
                            While Not (EOF(lHandle2) Or bError)
                                Line Input #lHandle2, cmdline
                                cmdline = Trim(cmdline)
                                If Replace(cmdline, IPAdress, "") <> cmdline Then
                                    If Left(cmdline, Len(IPAdress)) = IPAdress Then
                                        sMACAdress = Trim(Replace(cmdline, IPAdress, ""))
                                        sMACAdress = Trim(UCase(Left(sMACAdress, InStr(1, sMACAdress, " "))))
                                        bError = True
      
                                    End If
                                    
                                End If
                                
                            Wend
                            
                        Close #lHandle2
                        
                        Kill IPFile
   
                        If Len(sMACAdress) < 2 Then
                            sMACAdress = ""
                            
                        End If
                        
                    End If
                    
                    sMACAdress = MACFormat(sMACAdress)
    
                    cmdline = sHostName & ";" & IPAdress & ";" & sMACAdress & ";"
                    Print #lHandle1, cmdline
                
                End If
                
            Else
                If KeyRetVal <> vbIgnore Then
                    KeyRetVal = MsgBox(" Die IP-Adresse " & IPAdress & " ist derzeit nicht erreichbar." & Space(10), vbExclamation + vbAbortRetryIgnore + vbDefaultButton2, Me.Caption)
                    If KeyRetVal = vbAbort Then
                        i = EndIP + 1
        
                        IPFile = ";" & IPAdress & ";nicht erreichbar, Abbruch durch Benutzer;"
                        Print #lHandle1, IPFile
        
                    ElseIf KeyRetVal = vbRetry Then
                        i = i - 1
                
                    ElseIf KeyRetVal = vbIgnore Then
                        If (i < EndIP) Then
                            KeyRetVal = MsgBox(" Alle nicht erreichbaren IP-Adressen ignorieren ?" & Space(10), vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption)
                            KeyRetVal = IIf(KeyRetVal = vbYes, vbIgnore, vbAbort)
                
                        End If
                        
                        IPFile = ";" & IPAdress & ";nicht erreichbar;"
                        Print #lHandle1, IPFile
                
                    End If
                    
                Else
                    IPFile = ";" & IPAdress & ";nicht erreichbar;"
                    Print #lHandle1, IPFile
                    
                End If
            
            End If
    
            DoEvents
            
            If Abort Then
                i = EndIP + 1
                
            End If
    
        Next i
    
    Close #lHandle1
    
    For i = 0 To txtIPAdress.Count - 1
        txtIPAdress(i).Enabled = True
    
    Next i
    NoOfPing.Enabled = True
    FrameMACFormat.Enabled = True
    OptionMACFormat(0).Enabled = True
    OptionMACFormat(1).Enabled = True
    OptionMACFormat(2).Enabled = True
    
    FrameLetterCase.Enabled = True
    OptionCase(0).Enabled = True
    OptionCase(1).Enabled = True
    
    cmdAbort.Enabled = False
    cmdSearch.Enabled = True
    
    lblInfo.Caption = "Suchlauf beendet"
    
    Call SaveSetup

End Sub

Private Sub cmdAbort_Click()

    Abort = True

End Sub

Private Sub txtIPAdress_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case KeyAscii
        Case 8, 48 To 57
        
        Case Else
            KeyAscii = 0
        
    End Select

End Sub

Private Sub txtIPAdress_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    If Val(Trim(txtIPAdress(Index))) > 255 Then
        txtIPAdress(Index) = 255
        
    End If
    

End Sub

Private Function ShellGetHandle(ByVal sDateiname As String, ByVal sCmdline As String, lHandle As Long) As Boolean
   
   Dim udtProcessInfo As PROCESS_INFORMATION
   Dim udtStartupInfo As STARTUPINFO
   Dim lSuccess As Long
   
   udtStartupInfo.cb = Len(udtStartupInfo)
   udtStartupInfo.dwFlags = &H1
   udtStartupInfo.wShowWindow = 0
   
   lSuccess = CreateProcess(0&, _
                            sDateiname & sCmdline, _
                            0&, _
                            0&, _
                            1&, _
                            &H20, _
                            0&, _
                            0&, _
                            udtStartupInfo, _
                            udtProcessInfo)
   
   If lSuccess = 1 Then
      lHandle = udtProcessInfo.hProcess
      ShellGetHandle = False
   
   Else
      ShellGetHandle = True
   
   End If

End Function

Private Function Wait(ByVal lHandle As Long) As Boolean

    Dim ExitCode As Long

    Call GetExitCodeProcess(lHandle, ExitCode)
    Wait = IIf(ExitCode = STILL_ACTIVE, True, False)
    If Wait = False And ExitCode Then
        'Process mit Fehlercode beendet
        MsgBox ExitCode & " - Batch mit Fehler beendet", vbCritical + vbOKOnly, Me.Caption
    
    End If

    Exit Function

End Function

Private Function MACFormat(sMACAdress As String) As String

    Dim i As Long, Dummy As String
    
    i = 0
    i = IIf(OptionMACFormat(1).Value, 1, i)
    i = IIf(OptionMACFormat(2).Value, 2, i)
    
    Select Case i
        Case 1
            MACFormat = Replace(sMACAdress, "-", ":")
            
        Case 2
            MACFormat = Replace(sMACAdress, ":", "")
            MACFormat = Replace(MACFormat, "-", "")
            For i = 0 To Len(MACFormat) - 1 Step 4
                If i Mod 4 = 0 Then
                    If i > 0 Then
                        Dummy = Dummy & "."
                    
                    End If
                    Dummy = Dummy & Mid(MACFormat, i + 1, 4)
                    
                End If
                
            Next i
            MACFormat = Dummy
        
        Case Else
            MACFormat = Replace(sMACAdress, ":", "-")
        
    End Select

    If OptionCase(1).Value Then
        MACFormat = LCase(MACFormat)
        
    Else
        MACFormat = UCase(MACFormat)
    
    End If
    MACFormat = "MAC - " & MACFormat
    
End Function

