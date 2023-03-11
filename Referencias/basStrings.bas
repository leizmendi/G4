Attribute VB_Name = "basStrings"
Option Compare Database
Option Explicit

Public Function RecDerTop(ByVal str As String, intRecorteDerecha As Integer, intMaxLen As Integer) As Variant
    On Error GoTo Error_RecDerTop
    Dim l As Long
    l = Len(str)
    If l <= intRecorteDerecha Then
        RecDerTop = Null
        Exit Function
    End If
    str = Left(str, l - intRecorteDerecha)
    l = Len(str)
    If intMaxLen = 0 Then
    Else
        If l > intMaxLen Then
            str = Left(str, intMaxLen)
        End If
    End If
    RecDerTop = str
Salir_RecDerTop:
    Exit Function
Error_RecDerTop:
    Select Case Err
        Case Else
            MsgBox "error nº " & Err & " en RecDerTop" & vbCrLf & Err.Description
            Resume Salir_RecDerTop
            Resume Next
    End Select
End Function

Public Function CambiaChar(vstrTexto As Variant, strCharBusca As String, strCharReemplaza As String) As Variant
    If IsNull(vstrTexto) Then
        CambiaChar = Null
        Exit Function
    End If
    
    Dim i As Integer, j As Integer, strTxt As String, strR As String, str1 As String
    strTxt = CStr(vstrTexto)
    For i = 1 To Len(strTxt)
        str1 = Mid(strTxt, i, 1)
        If str1 = strCharBusca Then
            strR = strR & strCharReemplaza
        Else
            strR = strR & str1
        End If
    Next i
    CambiaChar = strR
End Function

Public Function DameHHMM(ByRef lngHH As Variant, ByRef lngMM As Variant) As String
    ' pasándole un total de HH y de MM, devuelve en lngHH y lngMM los totales corregidos y
    ' la función devuelve el string
    lngHH = CLng(Nz(lngHH, 0))
    lngMM = CLng(Nz(lngMM, 0))
    lngHH = lngHH + lngMM \ 60
    lngMM = lngMM Mod 60
    DameHHMM = CStr(lngHH) & "h. " & CStr(lngMM) & "m."
End Function

Public Function NombreFicheroValido(ByVal strName As String) As String
    Dim i As Integer
    For i = 1 To Len(strName)
        If InStr("\/:*?<>|", Mid(strName, i, 1)) > 0 Then
            Mid(strName, i) = "_"
        End If
    Next i
    NombreFicheroValido = strName
End Function


Public Function BuscaCampoInStr(strTexto As String, Optional ByRef intStart As Integer = 1, Optional strReplace As String = "") As String
'Busca en el texto la cadena <<NombreCampo>> y devuelve NombreCampo y posición de la primera < de <<NombreCampo>>
    On Error GoTo HandleError
    Dim i As Integer, j As Integer, strR As String
    i = InStr(intStart, strTexto, "<<")
    If i > 0 Then
        j = InStr(i, strTexto, ">>")
        If j > i + 2 Then
            strR = Mid(strTexto, i + 2, j - 1 - (i + 2))
        End If
    End If
    If strReplace <> "" Then
    End If
    BuscaCampoInStr = strR
    intStart = i
HandleExit:
    Exit Function
HandleError:
    MsgBox Err.Description
    Resume HandleExit
End Function

Public Function ExtraeLin(varTexto As Variant, intLin As Integer, Optional intCRLF As Integer = 1) As String
    Dim i As Integer, j As Integer, k As Integer, intDesde As Integer, intHasta As Integer
    Dim strTexto As String
On Error GoTo Error_ExtraeLin
    If IsNull(varTexto) Then
        ExtraeLin = "!"
        Exit Function
    Else
        strTexto = CStr(varTexto)
    End If
    For i = 1 To intLin
        If j = 0 Then
            intDesde = 1
        Else
            intDesde = k 'j + IIf(intCRLF = 1, 2, 1)
        End If
        If intCRLF = 1 Then
            j = InStr(intDesde, strTexto, vbCrLf)
            k = j + 2
        Else
            j = InStr(intDesde, strTexto, vbCrLf)
            If j = 0 Then
                j = InStr(intDesde, strTexto, vbLf)
                k = j + 1
            Else
                k = j + 2
            End If
        End If
        If j = 0 Then
            Exit For
        End If
        intHasta = j - 1
    Next i
    If j = 0 Then
        If i = intLin Then
            If Len(strTexto) >= intDesde Then
                ExtraeLin = Mid(strTexto, intDesde)
            Else
                ExtraeLin = ""
            End If
        Else
            ExtraeLin = ""
        End If
    Else
        ExtraeLin = Mid(strTexto, intDesde, intHasta - intDesde + 1)
    End If

Salir_ExtraeLin:
    Exit Function
Error_ExtraeLin:
    Select Case Err
        Case Else
            MsgBox "Error nº " & Err & vbCrLf & Err.Description & vbCrLf & "En ExtraeLin"
            
    End Select
    
    Resume Salir_ExtraeLin

End Function

Public Function Llena(strTexto As String, intHasta As Integer) As String
    If Len(strTexto) >= intHasta Then
        Llena = Left(strTexto, intHasta)
    Else
        Llena = strTexto & Space(intHasta - Len(strTexto))
    End If
End Function

Public Function LlenaAutoExtVarLin(strTexto As String, intHasta As Integer, intDesde As Integer, ByRef strR_lin2 As String) As String
    Dim strR As String, intEspacio As Integer, i As Integer
Ini:
    If Len(strTexto) = intHasta Then
        strR = strR & strTexto
    ElseIf Len(strTexto) > intHasta Then
        'busca un espacio par aplicar el salto de linea
        intEspacio = False
        For i = intHasta + 1 To 2 Step -1
            If Mid(strTexto, i, 1) = " " Then
                strR = strR & Left(strTexto, i - 1) & Space(intHasta - (i - 1)) & vbCrLf & Space(intDesde)
                strTexto = Mid(strTexto, i + 1)
                intEspacio = True
                Exit For
            End If
        Next i
        If intEspacio = False Then
            strR = strR & Left(strTexto, intHasta) & vbCrLf & Space(intDesde)
            strTexto = Mid(strTexto, intHasta + 1)
        End If
        GoTo Ini
    Else
        strR = strR & strTexto & Space(intHasta - Len(strTexto))
    End If
    If Len(strR) = intHasta Then
        LlenaAutoExtVarLin = strR
        strR_lin2 = ""
    Else
        LlenaAutoExtVarLin = Mid(strR, 1, intHasta)
        strR_lin2 = Mid(strR, intHasta + 1)
    End If
End Function

Public Function LlenaIzq(strTexto As String, intHasta As Integer) As String
    If Len(strTexto) >= intHasta Then
        LlenaIzq = Right(strTexto, intHasta)
    Else
        LlenaIzq = Space(intHasta - Len(strTexto)) & strTexto
    End If
End Function

Public Function Centra(strTexto As String, intAncho As Integer) As String
    If Len(strTexto) >= intAncho Then
        Centra = Left(strTexto, intAncho)
    Else
        Centra = Llena(Space((intAncho - Len(strTexto)) \ 2) & strTexto, intAncho)
    End If
End Function

Public Function CentraChar(strTexto As String, intAncho As Integer, str1 As String, Optional intEspaciosTambien As Integer = False) As String
    Dim i As Integer, r As Integer, strR As String
    If Len(strTexto) >= intAncho Then
        CentraChar = Left(strTexto, intAncho)
    Else
        i = (intAncho - Len(strTexto)) \ 2
        r = intAncho - Len(strTexto) - i
        strR = Replace(LlenaChar("·", i) & strTexto & LlenaChar("·", r), "·", str1)
        If intEspaciosTambien Then strR = Replace(strR, " ", str1)
        CentraChar = strR
    End If
End Function

Public Function LlenaChar(str1 As String, intCant As Integer) As String
    Dim i As Integer
    If Len(str1) < 1 Then
        LlenaChar = ""
        Exit Function
    End If
    For i = 1 To intCant
        LlenaChar = LlenaChar & Left(str1, 1)
    Next i
End Function


Public Function LlenaIzqChar(strTexto As String, intHasta As Integer, str1 As String, Optional intMoverSignoMenos As Integer = False) As String
    If Len(strTexto) >= intHasta Then
        LlenaIzqChar = Right(strTexto, intHasta)
    Else
        If Len(strTexto) > 0 Then
            If Left(strTexto, 1) = "-" And intMoverSignoMenos Then
                LlenaIzqChar = "-" & LlenaChar(str1, intHasta - Len(strTexto)) & Mid(strTexto, 2)
                Exit Function
            End If
        End If
        LlenaIzqChar = LlenaChar(str1, intHasta - Len(strTexto)) & strTexto
    End If
End Function



Public Sub ImprimirTodosLosChar(bytDesde As Byte)
'Para ver la pag. de códigos de la impresora
    Dim strgLin As String
    Dim i As Integer
    For i = bytDesde To 255
        strgLin = strgLin & Format(i, "000") & "-" & Chr(i) & ", "
    Next i
    Open DameValorParam("AplicacionCurrentUd") & "lpt1" For Output As #1
    Print #1, strgLin
    Close
End Sub


Public Function LineasDeStr(varTexto As Variant, Optional intCRLF As Integer = 1) As Integer
    Dim i As Integer, j As Integer, intLin As Integer
    Dim strTexto As String
On Error GoTo Error_LineasDeStr
    If IsNull(varTexto) Then
        LineasDeStr = 0
        Exit Function
    Else
        strTexto = CStr(varTexto)
    End If
    j = 1
    While True
        If intCRLF = 1 Then
            i = InStr(j, strTexto, vbCrLf)
        Else
            i = InStr(j, strTexto, vbLf)
        End If
            
        intLin = intLin + 1
        If i = 0 Then
            GoTo sigue
        Else
            j = i + 1
        End If
    Wend
sigue:
    LineasDeStr = intLin

Salir_LineasDeStr:
    Exit Function
Error_LineasDeStr:
    Select Case Err
        Case Else
            MsgBox "Error nº " & Err & vbCrLf & Err.Description & vbCrLf & "En LineasDeStr"
            
    End Select
    
    Resume Salir_LineasDeStr

End Function

Public Sub PrimerEspacioDcha(strTexto As String, intStart As Integer, ByRef intAncho As Integer, Optional ByRef intvbCrLfSN As Integer)
    'Se le pasa una cadena y a partir de que carácter se examina intStart) y el ancho máximo de
    'la línea, en intAncho devuelve el primer espacio que encuentra para saltar de línea
    'Si encuentra un salto delínea lo devuelve
    'Si no encuentra espacio, devuelve el mismo intAncho
    On Error GoTo Error_PrimerEspacioDcha
    Dim i As Integer
    i = intStart
    While i <= intStart + intAncho
        If Mid(strTexto, i, 2) = vbCrLf Then
            intAncho = i
            intvbCrLfSN = True
            Exit Sub
        End If
        i = i + 1
    Wend
    intvbCrLfSN = False
    i = intStart + intAncho
    If Len(strTexto) < i Then Exit Sub
    While i > intStart
        If Mid(strTexto, i, 1) = " " Then
            intAncho = i
            Exit Sub
        End If
        i = i - 1
    Wend
Salir_PrimerEspacioDcha:
    Exit Sub
Error_PrimerEspacioDcha:
    Select Case Err
        Case Else
            MsgBox "Error nº " & Err & " en PrimerEspacioDcha" & vbCrLf & Err.Description
            Resume Salir_PrimerEspacioDcha
    End Select

End Sub

Public Function DimeTextoEnLineasDeAncho(strTexto As String, intAncho As Integer) As String
    On Error GoTo Error_DimeTextoEnLineasDeAncho
    Dim strR As String, intStart As Integer, intvbCrLf As Integer, j As Integer
Bucle:
    j = intAncho
    If Len(strTexto) > intAncho Then
        PrimerEspacioDcha strTexto, 1, j
        strR = strR & Mid(strTexto, 1, j) & vbCrLf
        strTexto = Mid(strTexto, j + 1)
        GoTo Bucle
    Else
        strR = strR & strTexto
    End If
    
    DimeTextoEnLineasDeAncho = strR
Salir_DimeTextoEnLineasDeAncho:
    
    Exit Function
Error_DimeTextoEnLineasDeAncho:
    Select Case Err
        Case Else
            MsgBox "error nº " & Err & " en DimeTextoEnLineasDeAncho" & vbCrLf & Err.Description
            Resume Salir_DimeTextoEnLineasDeAncho
    End Select

End Function


Public Function SinExtension(strFile As String) As String
    On Error GoTo HandleError
    Dim i As Integer, strR As String
    strR = strFile
    i = InStrRev(strFile, ".")
    If i > 0 Then
        strR = Left(strFile, i - 1)
    End If
    SinExtension = strR
HandleExit:
    Exit Function
HandleError:
    MsgBox Err.Description
    Resume HandleExit
End Function

Public Function SpNz(v, alter)
    On Error GoTo Error_SpNz
    v = IIf(v = "", Null, v)
    SpNz = Nz(v, alter)
Salir_SpNz:
    Exit Function
Error_SpNz:
    Select Case Err
        Case Else
            'MsgBox "error nº " & Err & " en SpNz" & vbCrLf & Err.Description
            Resume Salir_SpNz
    End Select
End Function

Function ExtensionDeFile(strFile As String, Optional intMasUltimosCaracteres As Integer = 0) As String
    On Error GoTo HandleError
    Dim i As Integer, strR As String
    i = InStrRev(strFile, ".")
    If i > 0 Then
        If intMasUltimosCaracteres < i Then
            i = i - intMasUltimosCaracteres
        End If
        strR = Mid(strFile, i)
    End If
    ExtensionDeFile = strR
HandleExit:
    Exit Function
HandleError:
    MsgBox Err.Description
    Resume HandleExit
End Function

Public Function NoSemiColon(strT As String, Optional strPor As String = ",") As String
    On Error GoTo HandleError
    NoSemiColon = Replace(strT, ";", strPor)
HandleExit:
    Exit Function
HandleError:
    MsgBox Err.Description
    Resume HandleExit

End Function

Public Function SoloNum(strT As String) As String
    On Error GoTo HandleError
    Dim strR As String
    If IsNumeric(strT) Then
        strR = strT
    Else
        Dim i As Integer
        While i < Len(strT)
            i = i + 1
            If IsNumeric(Mid(strT, i, 1)) Then
                strR = strR & Mid(strT, i, 1)
            End If
        Wend
    End If
    SoloNum = strR
HandleExit:
    Exit Function
HandleError:
    MsgBox Err.Description
    Resume HandleExit

End Function


