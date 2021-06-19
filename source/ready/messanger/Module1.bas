Attribute VB_Name = "Module1"
Option Explicit

Public UserName As String
Public Function UpdateState() As Integer  'fun for controling sck.state

' print the state of sck
Select Case frmLogon.wsNet.State
    Case 0
        frmLogon.lblState.Caption = "Status: Not connected" ' "0 - sckClosed"
        frmOM.lblState.Caption = frmLogon.lblState.Caption
    Case 1
        frmLogon.lblState.Caption = "Status: Opening Connection" ' "1 - sckOpen"
        frmOM.lblState.Caption = frmLogon.lblState.Caption
    Case 2
        frmLogon.lblState.Caption = "Status: Connected to port.... waiting for remote user for responce." '"2 - sckListening"
        frmOM.lblState.Caption = frmLogon.lblState.Caption
    Case 3
        frmLogon.lblState.Caption = "Status: Connection pending." '"3 - sckConnectionPending"
        frmOM.lblState.Caption = frmLogon.lblState.Caption
    Case 4
        frmLogon.lblState.Caption = "Status: Resolving Host" '"4 - sckResolvingHost"
        frmOM.lblState.Caption = frmLogon.lblState.Caption
    Case 5
        frmLogon.lblState.Caption = "Status: Host Resolved" '5 - sckHostResolved"
        frmOM.lblState.Caption = frmLogon.lblState.Caption
    Case 6
        frmLogon.lblState.Caption = "Status: Connecting to port ..." '"6 - sckConnecting"
        frmOM.lblState.Caption = frmLogon.lblState.Caption
    Case 7
        frmLogon.lblState.Caption = "Status: Connected..." '7 - sckConnected"
        frmOM.lblState.Caption = frmLogon.lblState.Caption
    Case 8
        frmLogon.lblState.Caption = "Status: Connection closed..." '8 - sckClosing"
        frmOM.lblState.Caption = frmLogon.lblState.Caption
    Case 9
        frmLogon.lblState.Caption = "Status: An error occured while connecting.." '"9 - sckError"
        frmOM.lblState.Caption = frmLogon.lblState.Caption
End Select
 UpdateState = frmLogon.wsNet.State
End Function
