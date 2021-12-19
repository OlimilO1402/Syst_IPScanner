Attribute VB_Name = "MainModule"
Option Explicit

Sub Main()

    If App.PrevInstance Then
        Exit Sub
    
    End If
    
    frmMain.Show

End Sub
