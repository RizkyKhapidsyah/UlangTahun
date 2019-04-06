Attribute VB_Name = "Module1"
Option Explicit

Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset
Public Kalimat As String
Public X As Integer
Global Objek As Control

Public Sub Nyambung()
    If CN.State = adStateOpen Then CN.Close
    CN.CursorLocation = adUseClient
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data.rdb;Persist Security Info=False"
End Sub



