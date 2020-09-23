Attribute VB_Name = "modEncrypt"
Global EncryptionKey As String

Public Function SplitDataStreams(ByVal Data As String) As Variant
Dim Tv() As String, C As Long
Do
DoEvents

    Dim TMP As String
    TMP = VBA.Left(Data, 4 + Val(VBA.Left(Data, 4)))
    TMP = VBA.Mid(TMP, 5)
    Data = VBA.Mid(Data, Len(TMP) + 5)
    
    ReDim Preserve Tv(1 To C + 1) As String
    C = C + 1
    Tv(C) = TMP
    

Loop Until Len(Data) = 0

Dim n As Variant
n = Split(String(C, "-"), "-")

For I = 0 To UBound(n) - 1
    n(I) = Tv(I + 1)
Next

SplitDataStreams = n
End Function

Public Function Encrypt(ByVal Info As String, Optional ByVal Key As String) As String
    Dim TMP As String
    Dim A As Long, B As Long, C As Long, vCur As Long
    Randomize
    A = Int(Rnd * 2555) And 255
    B = Int(Rnd * 2555) And 255
    C = Int(Rnd * 2555) And 255
        
    TMP = Chr(A) & Chr(B) & Chr(C)
    
    Encrypt = Format(Len(Info) + 3, "0000") & TMP & CryptData(Info, A, B, C, Key)
End Function

Public Function Decrypt(ByVal Info As String, Optional ByVal Key As String) As String
    Dim TMP As String
    Dim A As Long, B As Long, C As Long, vCur As Long
    Randomize
       
    A = Asc(VBA.Mid(Info, 1, 1))
    B = Asc(VBA.Mid(Info, 2, 1))
    C = Asc(VBA.Mid(Info, 3, 1))
    Info = VBA.Mid(Info, 4)
    
    Decrypt = CryptData(Info, A, B, C, Key)
End Function

Private Function CryptData(ByVal Info As String, ByVal A As Long, ByVal B As Long, ByVal C As Long, Optional ByVal Key As String) As String
    Dim TMP As String, P As Long
    
    For I = 1 To Len(Info)
        vCur = Asc(VBA.Mid(Info, I, 1))
        vCur = vCur Xor A
        vCur = vCur Xor B
        vCur = vCur Xor C
        
        A = A Xor B
        B = B + 20 And 255
        B = B Xor C
        C = C + 20 And 255
        C = C Xor A
        
        If Key <> "" Then
        
            P = (A + B + C) And Len(Key)
            If P = 0 Then P = 1
            P = Asc(VBA.Mid(Key, P, 1))
            vCur = vCur Xor P
        
        End If
        
        TMP = TMP & Chr(vCur)
    Next
    
    CryptData = TMP
End Function
