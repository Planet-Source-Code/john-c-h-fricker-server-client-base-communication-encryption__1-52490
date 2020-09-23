VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Project1.ServerCtl svr 
      Left            =   1800
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    svr.StartServer 5410
    EncryptionKey = "testing2k4"
End Sub

Public Function SendData(ByVal Data As Variant, ByVal IDX As Long)
    svr.SendData IDX, Encrypt(CStr(Data), EncryptionKey)            'SEND DATA TO A CLIENT
End Function

Public Function BroadcastData(ByVal Data As Variant)
    svr.BroadcastData Encrypt(CStr(Data), EncryptionKey)            'SEND DATA TO ALL CLIENTS
End Function

Private Sub svr_ClientSendData(ByVal IDX As Long, ByVal RemoteIP As String, ByVal Data As String, ByVal BytesTotal As Long)
On Error GoTo err
    Dim vT As Variant, vData As String
    vT = SplitDataStreams(Data)                                     'SPLIT DATA INTO STREAMS

    For I = 0 To UBound(vT) - 1
        vData = CStr(Decrypt(CStr(vT(I)), EncryptionKey))           'DECRYPT
        HandleClientData IDX, RemoteIP, vData, Len(vData)           'PARSE
    Next
err:
End Sub

Private Sub HandleClientData(ByVal IDX As Long, ByVal RemoteIP As String, ByVal Data As String, ByVal BytesTotal As Long)
On Error GoTo err
    ''PARSE FUNCTION
    MsgBox Data
    SendData "RETURN DATA: " & Data, IDX
err:
End Sub
