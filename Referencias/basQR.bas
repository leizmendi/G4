Attribute VB_Name = "basQR"
Option Compare Database
Option Explicit


Public Function GeneraQR(sCodigo As String) As Boolean
    On Error GoTo HandleError
    Dim sFileQR As String
    sFileQR = GetCarpetaQR() & NombreFicheroValido(sCodigo) & ".gif"
    qrcodeCreateImage sFileQR, sCodigo
    GeneraQR = Dir(sFileQR) <> ""
HandleExit:
    Exit Function
HandleError:
    MsgBox Err.Description
    Resume HandleExit
End Function


Public Function GetCarpetaQR() As String
    GetCarpetaQR = Nz(GetParam("CarpetaImagenesQR", False), "")
End Function
