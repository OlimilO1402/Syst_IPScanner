VERSION 5.00
Begin VB.Form FrmIPPingScanner 
   Caption         =   "IPPingScanner"
   ClientHeight    =   5775
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   11655
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows-Standard
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
   Begin VB.CommandButton BtnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   10800
      TabIndex        =   6
      Top             =   80
      Width           =   855
   End
   Begin VB.CommandButton BtnGetNext50IPs 
      Caption         =   "Scan next 50 IPs"
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   80
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
      TabIndex        =   9
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Valid IP-addresses:"
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   120
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Invalid IP-addresses:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
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
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Visible         =   0   'False
      Begin VB.Menu mnuOption1 
         Caption         =   "Option"
      End
   End
End
Attribute VB_Name = "FrmIPPingScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_IPBase As IPAddressV4
Private m_LastIP As IPAddressV4
Private m_IPAddresses As IPAddresses
Private WithEvents IPPingScanner As IPPingScanner
Attribute IPPingScanner.VB_VarHelpID = -1
Private StartIPb4  As Long
Private SearchNIPs As Long
Private m_CancelFlag As VbMsgBoxResult

Private Sub Command1_Click()
    Dim IP As IPAddressV4: Set IP = MNew.IPAddressV4("0.0.0.0")
    'IP.OneUp
    MsgBox IP.IPToStr
    
    IP.Add 512
    MsgBox IP.IPToStr
    MsgBox IP.Address
    MsgBox Hex(IP.LAddress)
    
End Sub

Private Sub Form_Load()
    Set m_IPBase = MNew.IPAddressV4("192.168.178")
    TxtIPBase.Text = m_IPBase.IPToStr
    Set m_IPAddresses = New IPAddresses
    Set IPPingScanner = New IPPingScanner
    SearchNIPs = 50
    Set m_LastIP = m_IPBase.Clone
End Sub

Private Sub BtnGetNext50IPs_Click()
    LblDTime.Caption = "Scanning..."
    DoEvents
    Dim dt As Single: dt = Timer
    Dim c As Long: c = SearchNIPs - 1
    IPPingScanner.Scan m_LastIP, StartIPb4, StartIPb4 + c
    StartIPb4 = StartIPb4 + c + 1
    m_LastIP.Add SearchNIPs
    dt = Timer - dt
    LblDTime.Caption = "Ready! time: " & Format(dt, "#.00") & " sec"
End Sub

Private Sub BtnClear_Click()
    m_IPAddresses.Clear
    TxtIPInfo.Text = ""
    List1.Clear
    List2.Clear
    StartIPb4 = 0
    'SearchNIPs = 50
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbRightButton Then
        PopupMenu mnuOption
    End If
End Sub

Private Sub Form_Resize()
    Dim L As Single, t As Single, W As Single, H As Single
    L = List1.Left: t = List1.Top: W = List1.Width: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then List1.Move L, t, W, H
    L = List2.Left: t = List2.Top: W = List2.Width ': H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then List2.Move L, t, W, H
    L = TxtIPInfo.Left: t = TxtIPInfo.Top: W = Me.ScaleWidth - L
    If W > 0 And H > 0 Then TxtIPInfo.Move L, t, W, H
End Sub

Private Sub IPPingScanner_FoundIP(aIPV4 As IPAddressV4, out_Cancel As Boolean)
    If Not m_IPAddresses.Contains(aIPV4.IPToStr) Then
        m_IPAddresses.Add aIPV4
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
    If m_IPAddresses.Contains(List1.List(i)) Then
        Set aIPV4 = m_IPAddresses.Item(List1.List(i))
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
    If m_IPAddresses.Contains(List2.List(i)) Then
        Set aIPV4 = m_IPAddresses.Item(List2.List(i))
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
    If m_IPAddresses.Contains(s) Then
        Set aIPV4 = m_IPAddresses.Item(s)
        aIPV4.CallPing
        TxtIPInfo.Text = aIPV4.ToInfoStr
    End If
End Sub

Private Sub mnuOption1_Click()
    Dim s As String
    s = InputBox("n: ", "How many ips to search each time?", SearchNIPs)
    If s = vbNullString Then Exit Sub
    If Not IsNumeric(s) Then mnuOption1_Click
    SearchNIPs = CLng(s)
    BtnGetNext50IPs.Caption = "Scan next " & SearchNIPs & " IPs"
End Sub

Private Sub TxtIPBase_LostFocus()
    Dim NewBaseIP As IPAddressV4: Set NewBaseIP = MNew.IPAddressV4(TxtIPBase.Text)
    TxtIPBase.Text = NewBaseIP.IPToStr
    Set m_IPBase = NewBaseIP
    StartIPb4 = m_IPBase.b1
    Set m_LastIP = m_IPBase.Clone
End Sub
