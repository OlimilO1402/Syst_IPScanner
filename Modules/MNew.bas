Attribute VB_Name = "MNew"
Option Explicit

Sub Main()
    FrmIPPingScanner.Show
End Sub

Public Function IPAddressV4(StrLngBytesNewAddress, Optional aName As String) As IPAddressV4
    Set IPAddressV4 = New IPAddressV4: IPAddressV4.New_ StrLngBytesNewAddress, aName
End Function

Public Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function
    
