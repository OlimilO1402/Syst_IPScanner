VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmIPPingScanner 
   Caption         =   "IPScanner"
   ClientHeight    =   6375
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12495
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   12495
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Panel1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   12345
      TabIndex        =   6
      Top             =   480
      Width           =   12375
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5235
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   2295
      End
      Begin VB.PictureBox Panel2 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   2400
         ScaleHeight     =   5505
         ScaleWidth      =   9705
         TabIndex        =   7
         Top             =   0
         Width           =   9735
         Begin VB.TextBox TxtIPInfo 
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5295
            Left            =   2400
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Beides
            TabIndex        =   10
            Top             =   0
            Width           =   7095
         End
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5235
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   2295
         End
      End
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   1680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton BtnScanNextXXIPs 
      Caption         =   "Scan next 50 IPs"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.TextBox TxtIPBase 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4800
      TabIndex        =   0
      Text            =   "192.168.178"
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label LblDTime 
      AutoSize        =   -1  'True
      Caption         =   "            "
      Height          =   315
      Left            =   9000
      TabIndex        =   5
      Top             =   0
      Width           =   1620
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Valid IP-addresses:"
      Height          =   195
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Invalid IP-addresses:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1470
   End
   Begin VB.Label Label2 
      Caption         =   "IP-Base:"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "&Option"
      Begin VB.Menu mnuOptOption 
         Caption         =   "Option"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " &? "
      Begin VB.Menu mnuHelpExternalIP 
         Caption         =   "My external IP"
      End
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "Info"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuPopAddPort 
         Caption         =   "Add port"
      End
      Begin VB.Menu mnuPopEditPort 
         Caption         =   "Edit port"
      End
      Begin VB.Menu mnuOptPingIP 
         Caption         =   "Ping IP"
      End
      Begin VB.Menu mnuOptPingIPport 
         Caption         =   "Ping IP:port"
      End
      Begin VB.Menu mnuOptNslookupIP 
         Caption         =   "Nslookup IP"
      End
      Begin VB.Menu mnuOptNslookupIPport 
         Caption         =   "Nslookup IP:port"
      End
   End
End
Attribute VB_Name = "FrmIPPingScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents IPPingScanner As IPPingScanner
Attribute IPPingScanner.VB_VarHelpID = -1
Private m_CancelFlag As VbMsgBoxResult
Private WithEvents Splitter1 As Splitter
Attribute Splitter1.VB_VarHelpID = -1
Private Splitter2 As Splitter
Private mnuPopListBox As ListBox

Private Sub Form_Load()
    Set IPPingScanner = New IPPingScanner
    'TxtIPBase.Text = MApp.IPBase.IPToStr
    Panel1.BorderStyle = BorderStyleConstants.vbTransparent
    Panel2.BorderStyle = 0
    
    Set Splitter1 = MNew.Splitter(False, Me, Panel1, "Splitter1", List1, Panel2)
    Splitter1.BorderStyle = bsXPStyl
    Splitter1.LeftTopPos = List1.Width
    
    Set Splitter2 = MNew.Splitter(False, Me, Panel2, "Splitter2", List2, TxtIPInfo)
    Splitter2.BorderStyle = bsXPStyl
    Splitter2.LeftTopPos = List2.Width
    mnuPopup.Visible = False
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = MouseButtonConstants.vbRightButton Then
        PopupMenu mnuOpt
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MApp.Doc.IsDataChanged Then
        Dim mbr As VbMsgBoxResult: mbr = MsgBox("Data changed, save it?", vbYesNoCancel)
        Select Case mbr
        Case vbYes:  mnuFileSave_Click
                     'Cancel = True
        Case vbNo:   'just quit
        Case vbCancel: Cancel = True
        End Select
    End If
End Sub
Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single

'    L = List1.Left: T = List1.Top: W = List1.Width: H = Me.ScaleHeight - T
'    If W > 0 And H > 0 Then List1.Move L, T, W, H
'    L = List2.Left: T = List2.Top: W = List2.Width ': H = Me.ScaleHeight - T
'    If W > 0 And H > 0 Then List2.Move L, T, W, H
'    L = TxtIPInfo.Left: T = TxtIPInfo.Top: W = Me.ScaleWidth - L
'    If W > 0 And H > 0 Then TxtIPInfo.Move L, T, W, H
    Dim brdr As Single: 'brdr = 8 * Screen.TwipsPerPixelX
    L = 0: T = Panel1.Top
    W = Me.ScaleWidth - L
    H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Panel1.Move L, T, W, H
    
End Sub

Public Sub UpdateView()
    Dim Doc As Document: Set Doc = MApp.Doc
    Me.TxtIPBase.Text = Doc.IPBase.ToStr
    Dim ip As IPAddressV4
    Dim i As Long
    For i = 1 To Doc.IPAddresses.Count
        Set ip = Doc.IPAddresses.ItemI(i)
        If ip.IsValid Then
            List2.AddItem ip.IPToStr
        Else
            List1.AddItem ip.IPToStr
        End If
    Next
    BtnScanNextXXIPs.Caption = "Scan next " & MApp.Doc.SearchNIPs & " IPs"
End Sub

' v ############################## v '  Menu mnuFile  ' v ############################## v '
Private Sub mnuFileNew_Click()
    'MApp.Doc.IPAddresses.Clear
    
    MApp.NewDoc
    
    TxtIPInfo.Text = ""
    List1.Clear
    List2.Clear
    
    'MApp.Doc.StartIPb4 = 0
    'StartIPb4 = 0
    'SearchNIPs = 50
End Sub
Private Sub mnuFileOpen_Click()
    'Me.FileDialog.FileName = "IPScan" & Now
    If MApp.DlgFileOpen_Show(Me.FileDialog) = vbCancel Then Exit Sub
    MApp.FileOpen Me.FileDialog.FileName
    MApp.Doc.IsDataChanged = False
    UpdateView 'MApp.Doc
End Sub
Private Sub mnuFileSave_Click()
    MApp.FileSave
    'TODO: we should first ask if filesave was True
    MApp.Doc.IsDataChanged = False
End Sub
Private Sub mnuFileSaveAs_Click()
'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/cc144156(v=vs.85)
    If MApp.DlgFileSave_Show(Me.FileDialog) = vbCancel Then Exit Sub
    MApp.FileSave Me.FileDialog.FileName
    'TODO: we should first ask if filesave was True
    MApp.Doc.IsDataChanged = False

End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub
' ^ ############################## ^ '  Menu mnuFile  ' ^ ############################## ^ '

' v ############################## v '  Menu mnuOpt   ' v ############################## v '
Private Sub mnuOptOption_Click()
    Dim s As String
    s = InputBox("n: ", "How many ips to search each time?", MApp.Doc.SearchNIPs)
    If s = vbNullString Then Exit Sub
    If Not IsNumeric(s) Then
        mnuOptOption_Click
        Exit Sub
    End If
    MApp.Doc.SearchNIPs = CLng(s)
    BtnScanNextXXIPs.Caption = "Scan next " & MApp.Doc.SearchNIPs & " IPs"
End Sub

Private Sub mnuPopAddPort_Click()
    Dim i As Long: i = mnuPopListBox.ListIndex
    If i < 0 Then Exit Sub
    Dim s As String: s = mnuPopListBox.List(i)
    Dim aIPV4 As IPAddressV4
    If MApp.Doc.IPAddresses.Contains(s) Then
        Set aIPV4 = MApp.Doc.IPAddresses.Item(s)
        'aIPV4.CallPing
        TxtIPInfo.Text = aIPV4.ToInfoStr
    End If
End Sub

Private Sub mnuOptPingIP_Click()
    Dim i As Long: i = mnuPopListBox.ListIndex
    If i < 0 Then Exit Sub
    Dim s As String: s = mnuPopListBox.List(i)
    Dim aIPV4 As IPAddressV4
    If MApp.Doc.IPAddresses.Contains(s) Then
        Set aIPV4 = MApp.Doc.IPAddresses.Item(s)
        aIPV4.CallPing
        TxtIPInfo.Text = aIPV4.ToInfoStr
    End If
End Sub
Private Sub mnuOptPingIPport_Click()
    Dim i As Long: i = mnuPopListBox.ListIndex
    If i < 0 Then Exit Sub
    Dim s As String: s = mnuPopListBox.List(i)
    Dim aIPV4 As IPAddressV4
    If MApp.Doc.IPAddresses.Contains(s) Then
        Set aIPV4 = MApp.Doc.IPAddresses.Item(s)
        Set aIPV4 = aIPV4.Clone
        Dim sp As String: sp = InputBox("Use port-number for IP: " & aIPV4.ToStr, "Port")
        If StrPtr(sp) = 0 Then Exit Sub
        aIPV4.Port = CLng(sp)
        aIPV4.CallPing
        MApp.Doc.IPAddresses_Add aIPV4
        mnuPopListBox.AddItem aIPV4.ToStr
        TxtIPInfo.Text = aIPV4.ToInfoStr
    End If
End Sub
Private Sub mnuOptNslookupIP_Click()
    MsgBox "n.y.i. todo: mnuOptNslookupIP_Click"
End Sub
Private Sub mnuOptNslookupIPport_Click()
    MsgBox "n.y.i. todo: mnuOptNslookupIPport_Click"
End Sub

' ^ ############################## ^ '  Menu mnuOpt   ' ^ ############################## ^ '

' v ############################## v '  Menu mnuInfo  ' v ############################## v '

Private Sub mnuHelpExternalIP_Click()
    Dim ipv4 As IPAddressV4: Set ipv4 = MNew.IPAddressV4(MApp.GetMyIP)
    Dim s As String: s = InputBox("My external IP:", , ipv4.ToStr)
End Sub

Private Sub mnuHelpInfo_Click()
    'MsgBox "MBO-Ing.com IPPingScanner " vbcrlf & App.CompanyName & " " & App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
    MsgBox App.CompanyName & " " & App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
' ^ ############################## ^ '  Menu mnuInfo  ' ^ ############################## ^ '

Private Sub TxtIPBase_LostFocus()
    Dim NewBaseIP As IPAddressV4: Set NewBaseIP = MNew.IPAddressV4(TxtIPBase.Text)
    TxtIPBase.Text = NewBaseIP.IPToStr
    Set MApp.Doc.IPBase = NewBaseIP
    MApp.Doc.StartIPb4 = MApp.Doc.IPBase.b1
    Set MApp.Doc.LastIP = MApp.Doc.IPBase.Clone
End Sub

Private Sub BtnScanNextXXIPs_Click()
    'MApp.DataChanged = True
    LblDTime.Caption = "Scanning..."
    DoEvents
    Dim dt As Single: dt = Timer
    Dim c As Long: c = MApp.Doc.SearchNIPs - 1
    IPPingScanner.Scan MApp.Doc.LastIP, MApp.Doc.StartIPb4, MApp.Doc.StartIPb4 + c
    MApp.Doc.StartIPb4 = MApp.Doc.StartIPb4 + c + 1
    MApp.Doc.LastIP.Add MApp.Doc.SearchNIPs
    dt = Timer - dt
    LblDTime.Caption = "Ready! time: " & Format(dt, "#.00") & " sec"
End Sub

Private Sub IPPingScanner_FoundIP(aIPV4 As IPAddressV4, out_Cancel As Boolean)
    If Not MApp.Doc.IPAddresses.Contains(aIPV4.IPToStr) Then
        MApp.Doc.IPAddresses.Add aIPV4
    Else
        Dim mbr As VbMsgBoxResult
        mbr = MsgBox("Already exists: " & aIPV4.ToStr, vbOKCancel)
        out_Cancel = mbr = vbCancel
    End If
    If aIPV4.IsValid Then
        List1.AddItem ""
        List2.AddItem aIPV4.IPToStr
    Else
        List1.AddItem aIPV4.IPToStr
    End If
    DoEvents
End Sub

Private Sub List1_Click()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then Exit Sub
    Dim aIPV4 As IPAddressV4
    If MApp.Doc.IPAddresses.Contains(List1.List(i)) Then
        Set aIPV4 = MApp.Doc.IPAddresses.Item(List1.List(i))
        TxtIPInfo.Text = aIPV4.ToInfoStr
    End If
End Sub
Private Sub List1_DblClick()
    LB_DblClick List1
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim btn As MouseButtonConstants: btn = Button
    If btn = vbRightButton Then
        Set mnuPopListBox = List1
        PopupMenu mnuPopup
    End If
End Sub

Private Sub List2_Click()
    Dim i As Long: i = List2.ListIndex
    Dim s As String: s = List2.List(i)
    If i < 0 Then Exit Sub
    Dim aIPV4 As IPAddressV4
    If MApp.Doc.IPAddresses.Contains(s) Then
        Set aIPV4 = MApp.Doc.IPAddresses.Item(s)
        Debug.Print aIPV4.IPToStr
        TxtIPInfo.Text = aIPV4.ToInfoStr
    End If
End Sub

Private Sub List2_DblClick()
    LB_DblClick List2
End Sub
Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim btn As MouseButtonConstants: btn = Button
    If btn = vbRightButton Then
        Set mnuPopListBox = List2
        PopupMenu mnuPopup
    End If
End Sub


Private Sub LB_DblClick(aLB As ListBox)
    Dim i As Long: i = aLB.ListIndex
    If i < 0 Then Exit Sub
    Dim s As String: s = aLB.List(i)
    Dim aIPV4 As IPAddressV4
    If MApp.Doc.IPAddresses.Contains(s) Then
        Set aIPV4 = MApp.Doc.IPAddresses.Item(s)
        aIPV4.CallPing
        TxtIPInfo.Text = aIPV4.ToInfoStr
    End If
End Sub

'Private Sub UpdateView(aDoc As Document)

    'TxtIPBase.Text = aDoc.

'End Sub
