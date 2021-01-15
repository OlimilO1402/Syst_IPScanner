VERSION 5.00
Begin VB.Form FrmIPPingScanner 
   Caption         =   "IPPingScanner"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12135
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      Height          =   5295
      Left            =   5040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   840
      Width           =   7095
   End
   Begin VB.ListBox List2 
      Height          =   5325
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   5325
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton BtnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnGetNext50IPs 
      Caption         =   "Scan next 50 IPs"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Valid IP-addresses:"
      Height          =   195
      Left            =   2640
      TabIndex        =   8
      Top             =   600
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Invalid IP-addresses:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1470
   End
   Begin VB.Label Label2 
      Caption         =   "IP-Base:"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmIPPingScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_IPBase As IPAddressV4
Private m_IPAddresses As IPAddresses
Private WithEvents IPPingScanner As IPPingScanner
Attribute IPPingScanner.VB_VarHelpID = -1
Private StartIPb4 As Integer

Private Sub Form_Load()
    Set m_IPBase = MNew.IPAddressV4("192.168.1")
    Text1.Text = m_IPBase.IPToStr
    Set m_IPAddresses = New IPAddresses
    Set IPPingScanner = New IPPingScanner
End Sub

Private Sub BtnGetNext50IPs_Click()
    Dim aIPBase As IPAddressV4: Set aIPBase = MNew.IPAddressV4(Text1.Text)
    Dim c As Byte: c = 49
    IPPingScanner.Scan aIPBase, StartIPb4, StartIPb4 + c
    StartIPb4 = StartIPb4 + c + 1
End Sub

Private Sub BtnClear_Click()
    m_IPAddresses.Clear
    Text2.Text = ""
    List1.Clear
    List2.Clear
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    L = List1.Left: T = List1.Top: W = List1.Width: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then List1.Move L, T, W, H
    L = List2.Left: T = List2.Top: W = List2.Width ': H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then List2.Move L, T, W, H
    L = Text2.Left: T = Text2.Top: W = Me.ScaleWidth - L
    If W > 0 And H > 0 Then Text2.Move L, T, W, H
End Sub

Private Sub IPPingScanner_FoundIP(aIPV4 As IPAddressV4)
    If Not m_IPAddresses.Contains(aIPV4.IPToStr) Then
        m_IPAddresses.Add aIPV4
    Else
        If MsgBox("Already contained: " & aIPV4.ToStr, vbOKCancel) = vbCancel Then Exit Sub
    End If
    If aIPV4.IsActive Then
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
        Text2.Text = aIPV4.ToInfoStr
    End If
End Sub

Private Sub List2_Click()
    Dim i As Long: i = List2.ListIndex
    If i < 0 Then Exit Sub
    Dim aIPV4 As IPAddressV4
    If m_IPAddresses.Contains(List2.List(i)) Then
        Set aIPV4 = m_IPAddresses.Item(List2.List(i))
        Text2.Text = aIPV4.ToInfoStr
    End If
End Sub

Private Sub Text1_LostFocus()
    Dim NewBaseIP As IPAddressV4: Set NewBaseIP = MNew.IPAddressV4(Text1.Text)
    Text1.Text = NewBaseIP.IPToStr
    StartIPb4 = 0
    Set m_IPBase = NewBaseIP
End Sub
