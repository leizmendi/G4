Attribute VB_Name = "basXML"

Option Compare Database
Option Explicit


Public Sub PonXML_old(ByRef strXML As String, strNombreTag As String, vValor As Variant, Optional intSaltoLinea As Integer = True)
    strXML = strXML & "<" & strNombreTag & ">" & CStr(Nz(vValor, "")) & "</" & strNombreTag & ">" & IIf(intSaltoLinea, vbCrLf, "")
End Sub

Public Function DimeXML_old(strXML As String, strNombreTag As String, Optional intInicio As Integer = 1) As Variant
    Dim i As Integer, j As Integer, v As Variant
    i = InStr(intInicio, strXML, "<" & strNombreTag & ">")
    j = InStr(intInicio, strXML, "</" & strNombreTag & ">")
    If j = 0 Then j = InStr(intInicio, strXML, "<" & strNombreTag & "/>")
    If i = 0 Or j = 0 Or (i + Len(strNombreTag) + 2) > j Then
        DimeXML_old = Null
        Exit Function
    End If
    i = i + Len(strNombreTag) + 2
    v = Mid(strXML, i, j - i)
    DimeXML_old = v
End Function


Public Sub PonXML(ByRef strXML As String, strNombreTag As String, vValor As Variant, Optional intSustituye As Integer = False)
    Dim i As Integer, j As Integer
    If intSustituye = True Then
        i = InStr(1, strXML, "<" & strNombreTag & ">")
        j = InStr(1, strXML, "</" & strNombreTag & ">")
        If i = 0 Or j = 0 Or (i + Len(strNombreTag) + 2) > j Then GoTo AñadirTag
        strXML = Left(strXML, i + Len(strNombreTag) + 1) & CStr(Nz(vValor, "")) & Mid(strXML, j)
        Exit Sub
    End If
AñadirTag:
    strXML = strXML & "<" & strNombreTag & ">" & CStr(Nz(vValor, "")) & "</" & strNombreTag & ">"
End Sub

Public Function DimeXML(strXML As String, strNombreTag As String, Optional intInicio As Integer = 1, Optional intPosicion As Integer = 1) As Variant
    Dim i As Integer, j As Integer, v As Variant, p As Integer
    If intPosicion > 1 Then
        For p = 1 To intPosicion - 1
            i = InStr(intInicio, strXML, "<" & strNombreTag & ">", vbBinaryCompare)
            j = InStr(intInicio, strXML, "</" & strNombreTag & ">", vbBinaryCompare)
            If i = 0 Or j = 0 Or (i + Len(strNombreTag) + 2) > j Then
                DimeXML = Null
                Exit Function
            End If
            intInicio = j + 1
        Next p
    End If
    i = InStr(intInicio, strXML, "<" & strNombreTag & ">", vbBinaryCompare)
    j = InStr(intInicio, strXML, "</" & strNombreTag & ">", vbBinaryCompare)
    If i = 0 Or j = 0 Or (i + Len(strNombreTag) + 2) > j Then
        DimeXML = Null
        Exit Function
    End If
    i = i + Len(strNombreTag) + 2
    v = Mid(strXML, i, j - i)
    DimeXML = v
End Function

Public Function ReemplazaChars(strTxt As String, strBusca As String, strReemplazaPor As String) As String
    Dim strR As String, i As Integer, j As Integer
    j = 1
Bucle:
    i = InStr(j, strTxt, strBusca)
    While i > 0
        strR = strR & Mid(strTxt, j, i - j) & strReemplazaPor
        j = i + Len(strBusca)
        GoTo Bucle
    Wend
    strR = strR & Mid(strTxt, j)
    ReemplazaChars = strR
    
    
End Function

Public Function QuitaTagsXML(strXML As String, Optional intInicio As Integer = 1) As Variant
    Dim i As Integer, j As Integer, k As Integer, v As Variant
    Dim strR As String, strTag As String, strAux As String
    strR = strXML
    i = InStr(intInicio, strR, "<")
    If i > 0 Then
        j = InStr(i, strR, ">")
        If j > 0 Then
            strTag = Mid(strR, i + 1, j - (i + 1))
            k = InStr(j + 1, strR, "</" & strTag & ">")
            If k > 0 Then
                If i > 1 Then
                    'strR = Mid(strR, 1, i - 1) & Mid(strR, i + Len(strTag) + 2, k - (i + Len(strTag) + 2)) & Mid(strR, k + Len(strTag) + 3)
                    strR = Mid(strR, 1, i - 1) & Mid(strR, k + Len(strTag) + 3)
                Else
                    'strR = Mid(strR, i + Len(strTag) + 2, j - (i + Len(strTag) + 2)) & Mid(strR, k + Len(strTag) + 3)
                    strR = Mid(strR, k + Len(strTag) + 3)
                End If
            End If
        End If
        QuitaTagsXML = QuitaTagsXML(strR, i)
    Else
        QuitaTagsXML = strR
    End If
    
End Function

Public Function ComaPunto(ByVal vStr As Variant) As Variant
    ComaPunto = Replace(vStr, ",", ".")
End Function

