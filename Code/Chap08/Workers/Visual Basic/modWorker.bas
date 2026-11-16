Attribute VB_Name = "modWorker"
Option Explicit

Public Const SIZE_NAME As Integer = 20
Public Const SIZE_ADDRESS As Integer = 100
Public Const SIZE_PHONE As Integer = 20

Public Type Worker
    ID As Long
    Name As String * SIZE_NAME
    Address As String * SIZE_ADDRESS
    Phone As String * SIZE_PHONE
End Type

Public Sub SetWorkerData(ByRef pWorker As Worker, _
                         ByVal nID As Long, _
                         ByRef strName As String, _
                         ByRef strAddress As String, _
                         ByRef strPhone As String)
    pWorker.ID = nID
    pWorker.Name = strName
    pWorker.Address = strAddress
    pWorker.Phone = strPhone
End Sub

Public Sub PrintWorkerData(ByRef out As TextBox, _
                           ByRef pWorker() As Worker, _
                           ByVal nWorkers As Integer)
    Dim i As Integer
    For i = 0 To nWorkers - 1
        out.Text = out.Text & _
                   "ID: " & Str(pWorker(i).ID) & vbCrLf & _
                   "Name: " & pWorker(i).Name & vbCrLf & _
                   "Address: " & pWorker(i).Address & vbCrLf & _
                   "Phone: " & pWorker(i).Phone & vbCrLf & vbCrLf
    Next
End Sub
