VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FUserView 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "FUserView"
   ClientHeight    =   6240
   ClientLeft      =   795
   ClientTop       =   1080
   ClientWidth     =   8535
   Icon            =   "FUserView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8535
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "FUserView.frx":030A
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin MSComctlLib.ListView Users 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9763
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account Name"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Kommentar"
         Object.Width           =   7585
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8040
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FUserView.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FUserView.frx":05AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView SrvList 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6376
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Server Workstation"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Version"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FUserView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.
'
'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source!

Option Explicit
'Autor: Ruru, ruru@11mail.com

Dim OldUser As String
Dim Domain As String

Sub SeekComputer()
    
    Dim x As Integer, xItem As ListItem
    Dim ServerList As ListOfServer
    
    MousePointer = vbHourglass
    
    Me.SrvList.ListItems.Clear
    Me.SrvList.Enabled = False
    
    ServerList = EnumServer(SRV_TYPE_ALL)
    If ServerList.Init Then
        For x = 1 To UBound(ServerList.List)
            Set xItem = Me.SrvList.ListItems.Add(, , ServerList.List(x).ServerName)
            xItem.SubItems(1) = ServerList.List(x).Comment
            
            Select Case ServerList.List(x).Type
                Case Is >= 5
                    'xItem.Tag = "x"
                    xItem.SmallIcon = 1
                    xItem.SubItems(2) = " Workstation"
                Case Is = 4
                    xItem.SmallIcon = 2
                    xItem.SubItems(2) = " Server"
                Case Else
                    xItem.SmallIcon = 1
                    xItem.SubItems(2) = " Workstation"
            End Select
        Next
    End If
    
    Me.SrvList.Enabled = (Me.SrvList.ListItems.Count > 0)
    MousePointer = vbDefault
    
End Sub

Private Sub Command1_Click()
    SrvList.Visible = True
    Users.Visible = False
    Me.Caption = "Bitte Computer aus folgender Domäne wählen:   " & Domain
End Sub

Private Sub Form_Load()
    Dim x As Integer
    Dim xItem As ListItem
    Dim WksInfo As ServerInfo
    
    MousePointer = vbHourglass
    
    ' Einfaches lesen der Domäne über Wscript
    Dim objWshNet As Object
    Set objWshNet = CreateObject("Wscript.Network")
    
    Domain = objWshNet.userdomain
    Set objWshNet = Nothing
    
    If Domain <> "" Then
        Me.Caption = "Bitte Computer aus folgender Domäne wählen:   " & Domain
    Else
        MsgBox "Computer ist an keiner Domäne angeschlossen." & vbNewLine & _
            "Bitte Netzkabel und Netzverbindung überprüfen", vbExclamation, "Warnung"
        End
    End If
    
    MousePointer = vbDefault
    SeekComputer
    
    If CurrentServer <> "" Then
        Set xItem = Me.SrvList.FindItem(CurrentServer)
        If xItem Is Nothing Then
            Exit Sub
        Else
            xItem.EnsureVisible
            xItem.Selected = True
        End If
    End If
    
End Sub

Private Sub LoadAccountList(CurServer As String)
    Dim x As Integer
    Dim LocalUsers As ListOfUserExt
    Dim xItem As ListItem
    
    Me.Users.ListItems.Clear
    
    LocalUsers = LongEnumUsers(CurServer)
    If LocalUsers.Init Then
        For x = 1 To UBound(LocalUsers.List)
            Set xItem = Me.Users.ListItems.Add(, , LocalUsers.List(x).Name)
            xItem.SubItems(1) = LocalUsers.List(x).FullName
            xItem.SubItems(2) = LocalUsers.List(x).Comment
        Next
    End If
    
    Me.Users.Enabled = (Me.Users.ListItems.Count > 0)
    
    If Me.Users.Enabled Then
        Set xItem = Nothing
        If OldUser <> "" Then
            Set xItem = Me.Users.FindItem(OldUser)
        End If
        If xItem Is Nothing Then
            Set Me.Users.SelectedItem = Me.Users.ListItems(1)
        Else
            Set Me.Users.SelectedItem = xItem
        End If
        Me.Users.SelectedItem.EnsureVisible
    End If
    
    If LocalUsers.LastErr > 0 Then Unload Me
    
End Sub

Private Sub SrvList_DblClick()
    '    If Me.SrvList.SelectedItem.Tag = "x" Then
    '        Beep
    '        MsgBox "Bitte einen Server angeben"
    '    Else
    CurrentServer = Me.SrvList.SelectedItem.Text
    Me.Caption = "NT Userliste   " & CurrentServer
    SrvList.Visible = False
    Users.Visible = True
    LoadAccountList (CurrentServer)
    '    End If
End Sub

