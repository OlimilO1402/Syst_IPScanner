VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "AV - Bremen"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   493
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   801
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnScan 
      Caption         =   "Scan"
      Height          =   375
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   1095
   End
   Begin VB.Timer tReadNew 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   900
      Top             =   5910
   End
   Begin MSComctlLib.ImageList ilClients 
      Left            =   210
      Top             =   5820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pgWait 
      Height          =   135
      Left            =   45
      TabIndex        =   1
      Top             =   480
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ListView lvClients 
      Height          =   6645
      Left            =   45
      TabIndex        =   0
      Top             =   720
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   11721
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilClients"
      SmallIcons      =   "ilClients"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Clientname"
         Object.Width           =   2910
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP Adresse"
         Object.Width           =   2910
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Benutzer"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Zentriert
      Height          =   225
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   7005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sDomain As String
Private tCount As Long
Private lTime As Long

Private Sub BtnScan_Click()
    
    lTime = 30 'Minuten

    'Call FormPosition_Get(Me)

    Dim objWshNet As Object
    
    Set objWshNet = CreateObject("Wscript.Network")
    
    sDomain = objWshNet.userdomain
    
    Set objWshNet = Nothing
    
    Me.Show
    MousePointer = vbHourglass
    DoEvents
    
    With lvClients
        .Sorted = True ' Sortierte Anzeige
        .SortKey = 0   ' Sortierung nach erster Spalte
        .SortOrder = lvwAscending ' Aufsteigende Sortierung
    End With
    
    If sDomain <> vbNullString Then

        'If ListDhcpServer = True Then
            
            If ListComputer = True Then

                If GetIPFromHost = True Then

                    If ClientPing = True Then

                        lblStatus.Caption = "Nächster scan in: " & CStr(lTime) & " Minuten."

                        tReadNew.Enabled = True

                        pgWait.Max = lTime
                        pgWait.Min = 0
                        pgWait.Value = 0

                    End If

                End If

            End If

        'End If

    End If

    MousePointer = vbDefault

End Sub

Private Function ListComputer() As Boolean
    
    Dim xItem As ListItem
    Dim sAttributes As String
    Dim sBase As String
    Dim oCommand As Object
    Dim oConnection As Object
    Dim sFilter As String
    Dim oRecordset As Object
    Dim sResult As String
    Dim oRootDSE As Object
    Dim sScope As String
    Dim sSort As String
    Dim sDomain As String
    Dim bRet As Boolean
    
    bRet = False
    
    Set oConnection = CreateObject("ADODB.Connection")
    Set oCommand = CreateObject("ADODB.Command")
Try: On Error GoTo Catch
    Set oRootDSE = GetObject("LDAP://RootDSE")
    
    sDomain = oRootDSE.Get("DefaultNamingContext")
 
    sBase = "<LDAP://" & sDomain & ">;"
    
    ' Filter, nur Clients mit H*-* auflisten!
    sFilter = "(&(objectCategory=computer)(Name=*));"
    
    ' alle Clients!
    'sFilter = "(&(objectCategory=computer));"
    
    sAttributes = "Name;"
    sScope = "SubTree "
    sSort = "Name"
 
    oConnection.Provider = "ADsDSOObject"
    oConnection.Open "Active Directory Provider"
    
    oCommand.ActiveConnection = oConnection
    oCommand.Properties("Sort On") = sSort
    oCommand.Properties("Page Size") = 1000
    oCommand.CommandText = sBase & sFilter & sAttributes & sScope
 
    Set oRecordset = oCommand.Execute
 
    If oRecordset.Recordcount < 1 Then
    
    Else
    
        oRecordset.MoveFirst
        
        Do Until oRecordset.EOF
        
            Set xItem = lvClients.ListItems.Add(, , oRecordset.Fields("Name").Value)
            xItem.SmallIcon = 4
            oRecordset.MoveNext
            
        Loop
        
    End If
 
    oRecordset.Close
    oConnection.Close
 
    Set oCommand = Nothing
    Set oConnection = Nothing
    Set oRecordset = Nothing
    Set oRootDSE = Nothing
    
    If lvClients.ListItems.Count > 0 Then
    
        bRet = True
    
    End If
    
    ListComputer = bRet
    Exit Function
Catch:
    MsgBox Err.Number & vbCrLf & Err.Description
End Function

Private Function GetIPFromHost() As Boolean

    Dim lItem As Long
    Dim lCount As Long
    Dim xItem As ListItem
    Dim sDNS As String
    Dim sDNS_IP As String
    Dim sIP As String
    
    sDNS = GetPDCName
    sDNS_IP = vbNullString
    
    If Len(sDNS) > 0 Then
        
        sDNS_IP = HostName2IP(Mid$(sDNS, 3))
    
    End If
    
    lCount = lvClients.ListItems.Count
    
    For lItem = 1 To lCount
    
        Set xItem = lvClients.ListItems(lItem)
        
        sIP = HostName2IP(xItem.Text, sDNS_IP)
        
        If Len(sIP) = 0 Then
            xItem.SmallIcon = 2
        Else
            xItem.SmallIcon = 1
        End If
        
        xItem.SubItems(1) = sIP
        'xItem.SubItems(2) = GetMacFromClient(sIP)
        
        DoEvents
    Next

    GetIPFromHost = True
    
End Function

Private Function ClientPing() As Boolean

    Dim lItem As Long
    Dim lCount As Long
    Dim xItem As ListItem
    Dim sUserName As String
    Dim sMessage As String
    Dim sIP As String
    
    lCount = lvClients.ListItems.Count
    pgWait.Max = lCount
    pgWait.Min = 0
    pgWait.Value = 0
    
    For lItem = 1 To lCount
    
        Set xItem = lvClients.ListItems(lItem)
        
        lblStatus.Caption = "Scanne: " & xItem.Text
        
            sIP = HostName2IP(xItem.Text)
            
            If Len(sIP) = 0 Then
                xItem.SmallIcon = 2
            Else
                'xItem.SmallIcon = 1
            End If
        
            xItem.SubItems(1) = sIP
            
            If Len(Trim$(xItem.SubItems(1))) > 0 Then
            
                If SimplePing(xItem.SubItems(1)) = True Then
                
                    'xItem.SubItems(2) = "On"
                    xItem.SmallIcon = 3
                
                    sUserName = GetUser(xItem.Text)
                        
                    If Len(sUserName) = 0 Then
                        sMessage = "No User is logged on"
                    Else
                        If InStr(1, sUserName, "$") Then
                            sMessage = "No User is logged on"
                        Else
                            Mid$(sUserName, 1, 1) = UCase$(sUserName)
                            sMessage = sUserName
                        End If
                    End If
                    
                    xItem.SubItems(2) = sMessage
                
                Else
                    'xItem.SubItems(2) = "Off"
                    xItem.SmallIcon = 4
                    xItem.SubItems(2) = vbNullString
                End If
            
            End If
        
        pgWait.Value = lItem

        DoEvents
    Next

    pgWait.Value = 0
    
    ClientPing = True
    
End Function

Function GetUser(ByVal sComputer As String) As String
    
    Dim strRet As String
    Dim lngItem As Long
    Dim lngCount As Long
    
    Call modMain.LoggedOnUser("\\" & sComputer)
    GetUser = modMain.Users(0).wkui1_username
    
End Function

Private Sub Form_Resize()
    
'    Const lSize As Long = 7515
'    'Const lSize As Long = 5715
'
'    If Me.Width > lSize Then
'        Me.Move Me.Left, Me.Top, lSize
'        Exit Sub
'    End If
'
'    If Me.Width < lSize Then
'        Me.Move Me.Left, Me.Top, lSize
'        Exit Sub
'    End If
'
'    lblStatus.Move 2, 2, Me.ScaleWidth - 2, 15
'    pgWait.Move 2, Me.ScaleHeight - 8, Me.ScaleWidth - 2, 8
'    lvClients.Move 2, lblStatus.Top + lblStatus.Height + 2, Me.ScaleWidth - 2, Me.ScaleHeight - 27
    Dim brdr As Single: brdr = 3
    Dim L As Single: L = brdr
    Dim T As Single: T = brdr
    Dim W As Single: W = BtnScan.Width
    Dim H As Single: H = BtnScan.Height
    If W > 0 And H > 0 Then Me.BtnScan.Move L, T, W, H
    
    T = T + H + brdr
    W = Me.ScaleWidth - L - brdr
    H = pgWait.Height
    If W > 0 And H > 0 Then Me.pgWait.Move L, T, W, H
    
    T = T + H + brdr
    H = Me.ScaleHeight - T - brdr
    If W > 0 And H > 0 Then Me.lvClients.Move L, T, W, H
End Sub

Private Sub ReRead()
    
    tReadNew.Enabled = False
    Call ClientPing
    lblStatus.Caption = "Nächster scann in: " & CStr(lTime) & " Minuten."
    DoEvents
    pgWait.Max = lTime
    pgWait.Min = 0
    pgWait.Value = 0
    tReadNew.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call FormPosition_Put(Me)
End Sub

Private Sub lblStatus_DblClick()
    Call ReRead
End Sub

Private Sub lvClients_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    With lvClients
        .Sorted = True ' Sortierte Anzeige
        .SortKey = ColumnHeader.Index - 1 ' Sortierung nach erster Spalte
        If .SortOrder = lvwAscending Then ' Aufsteigende Sortierung
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With

End Sub

Private Sub tReadNew_Timer()

    tCount = tCount + 1
    pgWait.Value = tCount
    
    If lTime - tCount > 1 Then
        lblStatus.Caption = "Nächster scann in: " & CStr(lTime - tCount) & " Minuten."
    Else
        lblStatus.Caption = "Nächster scann in: " & CStr(lTime - tCount) & " Minute."
    End If
    
    If tCount = lTime Then
        tCount = 0
        Call ReRead
    End If

End Sub
