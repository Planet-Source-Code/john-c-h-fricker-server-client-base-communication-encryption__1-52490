VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl ServerCtl 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ServerCtl.ctx":0000
   ScaleHeight     =   420
   ScaleWidth      =   420
   ToolboxBitmap   =   "ServerCtl.ctx":0972
   Begin MSWinsockLib.Winsock sck 
      Index           =   0
      Left            =   2880
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "ServerCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

''''''''''''''''''''''
''  SERVER CONTROL  ''
''         by       ''
''                  ''
''John C. H. Fricker''
''''''''''''''''''''''

''NOTE:  SendDataEncrypt is an OPTIONAL event for encrypting outgoing data... this is so for things like BROADCASTDATA, stuff can still be encrypted for individual clients.  Of course can totally be ignored
''TelnetServer lets you wait for vbCrlf (enter) before counting the data as received

Public Event ServerStopped()
Public Event ServerStarted()
Public Event ServerStartError(ByVal Number As Long, ByVal Description As String)
Public Event ClientConnected(ByVal IDX As Long, ByVal RemoteIP As String)
Public Event ClientDisconnected(ByVal IDX As Long, ByVal RemoteIP As String)
Public Event ClientSendData(ByVal IDX As Long, ByVal RemoteIP As String, ByVal Data As String, ByVal BytesTotal As Long)
Public Event ClientSendComplete(ByVal IDX As Long)
Public Event ClientSendProgress(ByVal IDX As Long, ByVal bytesRemaining As Long, ByVal bytesSent As Long)
Public Event SendDataEncrypt(ByVal IDX As Long, ByRef Data As String)
Dim TelnetSer As Boolean

Public Sub StopServer()
On Error Resume Next
sck(0).Close                                                    'Close Server
Dim I As Long                                                   'Define I
For I = 1 To sck.ubound                                         'Close all client's connections and unload.
    CloseConnection I
Next
RaiseEvent ServerStopped
End Sub

Public Function CloseConnection(ByVal IDX As Long)
    If IDX = 0 Then Exit Function                               'Cannot close Server sock
    On Error Resume Next                                        'Just in case
    Call sck_Close(CInt(IDX))                                   'Close the connection
End Function

Public Function StartServer(ByVal LocalPort As Long, Optional ByVal TelnetServer As Boolean)
On Error GoTo err

    Call StopServer                                             'Disconnect any clients
    sck(0).LocalPort = LocalPort                                'Set listen port
    sck(0).Listen                                               'Listen
    
    TelnetSer = TelnetServer                                    'See the notes at the top of the file
    RaiseEvent ServerStarted                                    'Raise Started event

    Exit Function
err:
    RaiseEvent ServerStartError(err.Number, err.Description)    'Error!
End Function

Private Sub sck_Close(Index As Integer)
    
    RaiseEvent ClientDisconnected(Index, sck(Index).RemoteHostIP) 'Client Disconnected
    sck(Index).Close
    Unload sck(Index)
    
End Sub

Private Sub sck_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    Load sck(sck.ubound + 1)                                    'Open New Control to Accept connection, accept and raise
    sck(sck.ubound).Accept requestID                            'connect event
    RaiseEvent ClientConnected(sck.ubound, sck(sck.ubound).RemoteHostIP)
    
End Sub

Private Sub sck_DataArrival(Index As Integer, ByVal BytesTotal As Long)

    Dim TMPStr As String
    If TelnetSer = True Then                                    'Keep data in winsock buffer until vbCrlf if TELNETSERVER
        sck(Index).PeekData TMPStr
        If VBA.Right(TMPStr, 2) = vbCrLf Then GoTo getallnow    'Treat as normal as have now received vbCrlf
    Else
getallnow:
        sck(Index).GetData TMPStr                               'Get Data & Call Event
        RaiseEvent ClientSendData(Index, sck(Index).RemoteHostIP, TMPStr, BytesTotal)
    End If
End Sub

Private Sub sck_SendComplete(Index As Integer)
    RaiseEvent ClientSendComplete(Index)                        'Sent Data
End Sub

Private Sub sck_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    RaiseEvent ClientSendProgress(Index, bytesRemaining, bytesSent)     'Sending Data
End Sub

Public Function SendData(ByVal IDX As Long, ByVal Data As Variant)
On Error GoTo err
Dim TMP As String
TMP = CStr(Data)                                                        'Make data variant into string
RaiseEvent SendDataEncrypt(IDX, TMP)                                    'See notes
sck(IDX).SendData TMP                                                   'Send TMP

Exit Function
err:                                                                    'Just in case
End Function

Public Function BroadcastData(ByVal Data As Variant)
    Dim I As Long                                                       'Define I
    For I = 1 To sck.ubound
        SendData CLng(I), Data                                          'Send data to each client
    Next
End Function

Private Sub UserControl_Resize()
On Error GoTo err
UserControl.Width = 420
UserControl.Height = 420
err:
End Sub

Private Sub UserControl_Terminate()

    Call StopServer                                                     'Close server automatically
    
End Sub

Public Property Get ServerPort() As Long
    ServerPort = sck(0).LocalPort                                       'Server port
End Property

Public Property Get ServerHostName() As String
    ServerHost = sck(0).LocalHostName                                   'Name of the computer server is on
End Property

Public Function GetClientRemoteIP(ByVal IDX As Long) As String
    On Error Resume Next
    GetClientRemoteIP = sck(IDX).RemoteHostIP                           'Get ip of a client
End Function

Public Function GetClientRemotePort(ByVal IDX As Long) As Long
    On Error Resume Next
    GetClientRemotePort = sck(IDX).RemotePort                           'Get remote port of client
End Function

Public Function GetClientRemoteHost(ByVal IDX As Long) As String
    On Error Resume Next
    GetClientRemoteHost = sck(IDX).RemoteHost                           'Get remote host
End Function

Public Function GetClientLocalPort(ByVal IDX As Long) As Long
    On Error Resume Next
    GetClientLocalPort = sck(IDX).LocalPort                             'Get local port of client (also server port)
End Function
