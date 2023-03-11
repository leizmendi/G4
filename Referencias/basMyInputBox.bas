Attribute VB_Name = "basMyInputBox"
Option Compare Database
Option Explicit


Public Function MyInputBox(strTitulo As String, Optional strPrompt As String = "" _
                                            , Optional strDefault As String _
                                            , Optional strInputMask As String _
                                            , Optional intTextoEnriquecido As Integer = False, Optional strAdjuntarFicheros As String = "NO")
    On Error GoTo Error_MyInputBox
    Dim strXML As String
    PonXML strXML, "Titulo", strTitulo
    PonXML strXML, "Prompt", strPrompt
    PonXML strXML, "Default", strDefault
    PonXML strXML, "InputMask", strInputMask
    PonXML strXML, "TextoEnriquecido", IIf(intTextoEnriquecido, "S", "N")
    PonXML strXML, "AdjuntarFicheros", strAdjuntarFicheros
    DoCmd.OpenForm "frmMyInputBox", , , , , acDialog, strXML
    If Not IsOpenForm("frmMyInputBox") Then
        MyInputBox = Null
        GoTo Salir_MyInputBox
    End If
    MyInputBox = Nz(Forms("frmMyInputBox")("txtIntro"), "")
    strAdjuntarFicheros = Forms("frmMyInputBox")("lstAdjuntarFicheros").RowSource
    DoCmd.Close acForm, "frmMyInputBox"
Salir_MyInputBox:
    Exit Function
Error_MyInputBox:
    Select Case Err
        Case Else
            MsgBox "Error nº " & Err & " en MyInputBox" & vbCrLf & Err.Description
            Resume Salir_MyInputBox
    End Select

End Function

Public Function IsOpenForm(strForm As String) As Boolean
    On Error Resume Next
    Dim Foo As Form
    Set Foo = Forms(strForm)
    IsOpenForm = Err = 0
End Function

Public Function MyInput2List(strDisponiblesRowSource As String, _
                             strAsignadasRowSource As String, _
                             strSqlAdd As String, _
                             strSqlQuit As String, _
                             strTitulo As String, _
                    Optional strlblDisponibles As String = "Disponibles", _
                    Optional strlblAsignadas As String = "Asignadas", _
                    Optional strlblDescripcion As String, _
                    Optional intVerOkCancel As Integer = False) As Integer
    On Error GoTo HandleError
    Dim strXML As String
    PonXML strXML, "DisponiblesRowSource", strDisponiblesRowSource
    PonXML strXML, "AsignadasRowSource", strAsignadasRowSource
    PonXML strXML, "SqlAdd", strSqlAdd
    PonXML strXML, "SqlQuit", strSqlQuit
    PonXML strXML, "Titulo", strTitulo
    PonXML strXML, "lblDisponibles", strlblDisponibles
    PonXML strXML, "lblAsignadas", strlblAsignadas
    PonXML strXML, "lblDescripcion", strlblDescripcion
    PonXML strXML, "VerOKCancel", IIf(intVerOkCancel, "S", "N")
    DoCmd.OpenForm "frmMyInput2List", , , , , acDialog, strXML
    MyInput2List = True
    If intVerOkCancel = True Then
        If IsOpenForm("frmMyInput2List") Then
            DoCmd.Close acForm, "frmMyInput2List"
        Else
            MyInput2List = False
        End If
    End If
    
HandleExit:
    Exit Function
HandleError:
    Select Case Err
        Case Else
            MsgBox "Error nº " & Err & " en MyInput2List" & vbCrLf & Err.Description
            Resume HandleExit
    End Select

End Function


Public Sub My2ListGrupoClientes(lngIdCliente As Long)
    On Error GoTo HandleError
    Dim strSQL As String, strSqlAdd As String, strSqlQuit As String
    strSQL = "SELECT tbCliente_Grupo.Id, tbClientesGrupos.GrupoClientes, tbCliente_Grupo.IdGrupoClientes"
    strSQL = strSQL & " FROM tbClientesGrupos INNER JOIN tbCliente_Grupo ON tbClientesGrupos.IdGrupoClientes = tbCliente_Grupo.IdGrupoClientes"
    strSQL = strSQL & " WHERE tbCliente_Grupo.IdCliente=" & lngIdCliente
    CurrentDb.QueryDefs("qryAsignadas").SQL = strSQL
    strSQL = "SELECT tbClientesGrupos.IdGrupoClientes, tbClientesGrupos.GrupoClientes"
    strSQL = strSQL & " FROM tbClientesGrupos LEFT JOIN qryAsignadas ON tbClientesGrupos.IdGrupoClientes = qryAsignadas.IdGrupoClientes"
    strSQL = strSQL & " WHERE (((qryAsignadas.IdGrupoClientes) Is Null))"
    strSQL = strSQL & " ORDER BY tbClientesGrupos.Orden;"
    CurrentDb.QueryDefs("qryDisponibles").SQL = strSQL
    strSQL = "INSERT INTO tbCliente_Grupo(IdGrupoClientes, IdCliente)"
    strSQL = strSQL & " SELECT <<ItemData>> as IdGrupo, " & lngIdCliente & " as IdEmp"
    strSqlAdd = strSQL
    strSQL = "DELETE * FROM tbCliente_Grupo"
    strSQL = strSQL & " WHERE Id = <<ItemData>>"
    strSqlQuit = strSQL
    MyInput2List "qryDisponibles", "qryAsignadas", strSqlAdd, strSqlQuit, "Grupos del Cliente", , , "Grupos de " & DimeCliente(lngIdCliente)
    PonGruposCliente lngIdCliente
    
HandleExit:
    Exit Sub
HandleError:
    MsgBox Err.Description
    Resume HandleExit
End Sub




Public Sub MyListInfo(strTitulo As String, strRowSource As String, Optional strRowSourceType As String = "Tabla/Consulta", Optional intColumnCount As Integer = 1, Optional strColumnWidths As String)
    On Error GoTo HandleError
    Dim strXML As String
    PonXML strXML, "Titulo", strTitulo
    PonXML strXML, "RowSource", strRowSource
    PonXML strXML, "RowSourceType", strRowSourceType
    PonXML strXML, "ColumnCount", intColumnCount
    PonXML strXML, "ColumnWidths", strColumnWidths
    DoCmd.OpenForm "frmMyInfo", , , , , acDialog, strXML
HandleExit:
    Exit Sub
HandleError:
    MsgBox Err.Description
    Resume HandleExit
End Sub

