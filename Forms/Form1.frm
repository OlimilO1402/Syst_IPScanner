VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmIPPingScanner 
   Caption         =   "IPPingScanner"
   ClientHeight    =   5775
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11655
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows-Standard
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   1680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
      Left            =   4080
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   480
      Width           =   7575
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
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
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
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton BtnScanNext50IPs 
      Caption         =   "Scan next 50 IPs"
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   45
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
      Top             =   80
      Width           =   1935
   End
   Begin VB.Label LblDTime 
      AutoSize        =   -1  'True
      Caption         =   "            "
      Height          =   195
      Left            =   8880
      TabIndex        =   8
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Valid IP-addresses:"
      Height          =   195
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Invalid IP-addresses:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1470
   End
   Begin VB.Label Label2 
      Caption         =   "IP-Base:"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   120
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
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "Info"
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

Private Sub Form_Load()
    Set IPPingScanner = New IPPingScanner
    'TxtIPBase.Text = MApp.IPBase.IPToStr
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
        Case Cancel: Cancel = True
        End Select
    End If
End Sub
Private Sub Form_Resize()
    Dim l As Single, t As Single, W As Single, H As Single
    l = List1.Left: t = List1.Top: W = List1.Width: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then List1.Move l, t, W, H
    l = List2.Left: t = List2.Top: W = List2.Width ': H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then List2.Move l, t, W, H
    l = TxtIPInfo.Left: t = TxtIPInfo.Top: W = Me.ScaleWidth - l
    If W > 0 And H > 0 Then TxtIPInfo.Move l, t, W, H
End Sub

Public Sub UpdateView()
    Dim Doc As Document: Set Doc = MApp.Doc
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

End Sub
Private Sub mnuFileSave_Click()
    MApp.FileSave
End Sub
Private Sub mnuFileSaveAs_Click()
'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/cc144156(v=vs.85)
    If MApp.DlgFileSave_Show(Me.FileDialog) = vbCancel Then Exit Sub
    
    MApp.FileSave Me.FileDialog.FileName
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
    BtnScanNext50IPs.Caption = "Scan next " & MApp.Doc.SearchNIPs & " IPs"
End Sub
' ^ ############################## ^ '  Menu mnuOpt   ' ^ ############################## ^ '

' v ############################## v '  Menu mnuInfo  ' v ############################## v '
Private Sub mnuHelpInfo_Click()
    MsgBox "MBO-Ing.com IPPingScanner 1.0"
End Sub
' ^ ############################## ^ '  Menu mnuInfo  ' ^ ############################## ^ '

Private Sub TxtIPBase_LostFocus()
    Dim NewBaseIP As IPAddressV4: Set NewBaseIP = MNew.IPAddressV4(TxtIPBase.Text)
    TxtIPBase.Text = NewBaseIP.IPToStr
    Set MApp.Doc.IPBase = NewBaseIP
    MApp.Doc.StartIPb4 = MApp.Doc.IPBase.b1
    Set MApp.Doc.LastIP = MApp.Doc.IPBase.Clone
End Sub

Private Sub BtnScanNext50IPs_Click()
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

Private Sub List2_Click()
    Dim i As Long: i = List2.ListIndex
    If i < 0 Then Exit Sub
    Dim aIPV4 As IPAddressV4
    If MApp.Doc.IPAddresses.Contains(List2.List(i)) Then
        Set aIPV4 = MApp.Doc.IPAddresses.Item(List2.List(i))
        TxtIPInfo.Text = aIPV4.ToInfoStr
    End If
End Sub

Private Sub List2_DblClick()
    LB_DblClick List2
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

