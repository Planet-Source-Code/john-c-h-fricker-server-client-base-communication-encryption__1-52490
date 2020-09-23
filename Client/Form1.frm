VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
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
   Begin MSWinsockLib.Winsock sck 
      Left            =   2520
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    sck.Connect "localhost", 5410
    EncryptionKey = "testing2k4"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sck.Close
    End
End Sub

Private Sub sck_Close()
    sck.Close
    End
End Sub

Public Function SendData(ByVal Data As Variant)
    sck.SendData Encrypt(CStr(Data), EncryptionKey)
End Function

Private Sub sck_Connect()
    SendData "Hiya"
    SendData "Cya"
End Sub

Private Sub sck_DataArrival(ByVal BytesTotal As Long)
On Error GoTo err
    Dim vT As Variant, vData As String, Data As String
    sck.GetData Data
    vT = SplitDataStreams(Data)
    
    For I = 0 To UBound(vT)
        vData = CStr(Decrypt(vT(I), EncryptionKey))
        HandleServerData vData, Len(vData)
    Next
err:
End Sub

Private Sub HandleServerData(ByVal Data As String, ByVal BytesTotal As Long)
On Error GoTo err
    MsgBox Data
err:
End Sub

Private Sub sck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sck.Close
    End
End Sub
