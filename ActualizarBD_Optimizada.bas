Attribute VB_Name = "ActualizarBD_Optimizada"
'Attribute VB_Name = "ActualizarBD_Optimizada"
Option Explicit

Const TAMANO_LOTE As Long = 20

Public Sub ActualizarBD_Optimizada_Batch()
    Dim campos(1 To 105) As String
    CargarCamposActualizarBD campos

    Dim lo As ListObject
    Dim datos As Range
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cmdUpdate As ADODB.Command
    Dim cmdInsert As ADODB.Command
    Dim cmdDelete As ADODB.Command
    Dim cmdAux As ADODB.Command
    Dim i As Long, fila As Long
    Dim colCampos(1 To 105) As Long
    Dim tipoCampos(1 To 105) As Long
    Dim tamCampos(1 To 105) As Long
    Dim precisionCampos(1 To 105) As Byte
    Dim escalaCampos(1 To 105) As Byte
    Dim colAccion As Long, colIdExcel As Long, colIdBD As Long, colFecha As Long
    Dim tipoId1 As Long, tamId1 As Long, precisionId1 As Byte, escalaId1 As Byte
    Dim tipoId As Long, tamId As Long, precisionId As Byte, escalaId As Byte
    Dim sql As String, columnaError As String, accion As String
    Dim idAsignado As Variant, valor As Variant
    Dim fechaHoy As Date, tiempoInicio As Date
    Dim filasMarcadas As Long, filasProcesadas As Long, filasEnLote As Long
    Dim transaccionIniciada As Boolean
    Dim calcAnterior As XlCalculation
    Dim pantallaAnterior As Boolean, eventosAnteriores As Boolean
    Dim estadoAplicacionGuardado As Boolean
    Dim numeroError As Long, descripcionError As String, filaError As Long
    Dim loteActual As Long

    Dim filasLote() As Long
    Dim accionesLote() As String
    Dim idsLote() As Variant

    On Error GoTo ManejarError

    Set lo = ActiveCell.ListObject
    If lo Is Nothing Then Err.Raise vbObjectError + 1000, , "Selecciona una celda dentro de la tabla de datos."

    Set datos = lo.DataBodyRange
    If datos Is Nothing Then
        MsgBox "La tabla no contiene filas para procesar.", vbInformation
        Exit Sub
    End If

    colAccion = ColumnaTabla(lo, "AccionEnBD")
    colIdExcel = ColumnaTabla(lo, "Id")
    colIdBD = ColumnaTabla(lo, "ID_ActualizarBD")
    colFecha = ColumnaTabla(lo, "Fecha_Actualizacion")

    If colAccion = 0 Or colIdExcel = 0 Or colIdBD = 0 Or colFecha = 0 Then
        MsgBox "Falta una columna requerida: AccionEnBD, Id, ID_ActualizarBD o Fecha_Actualizacion.", vbExclamation
        Exit Sub
    End If

    For i = LBound(campos) To UBound(campos)
        colCampos(i) = ColumnaTabla(lo, campos(i))
        If colCampos(i) = 0 Then columnaError = columnaError & vbCrLf & "- " & campos(i)
    Next i

    If Len(columnaError) > 0 Then
        MsgBox "No encontré estas columnas en la tabla de Excel:" & columnaError, vbExclamation
        Exit Sub
    End If

    Set cn = New ADODB.Connection
    AbreConexion cn

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM Bd_Ppto WHERE 1=0", cn, adOpenForwardOnly, adLockReadOnly

    For i = LBound(campos) To UBound(campos)
        With rs.Fields(campos(i))
            tipoCampos(i) = .Type
            tamCampos(i) = CLng(.DefinedSize)
            precisionCampos(i) = .Precision
            escalaCampos(i) = .NumericScale
        End With
    Next i

    With rs.Fields("Id1")
        tipoId1 = .Type
        tamId1 = CLng(.DefinedSize)
        precisionId1 = .Precision
        escalaId1 = .NumericScale
    End With

    With rs.Fields("Id")
        tipoId = .Type
        tamId = CLng(.DefinedSize)
        precisionId = .Precision
        escalaId = .NumericScale
    End With
    rs.Close
    Set rs = Nothing

    fechaHoy = Now
    tiempoInicio = Now
    calcAnterior = Application.Calculation
    pantallaAnterior = Application.ScreenUpdating
    eventosAnteriores = Application.EnableEvents
    estadoAplicacionGuardado = True
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For fila = 1 To datos.Rows.Count
        accion = CStr(datos.Cells(fila, colAccion).Value)
        If EsAccionProcesable(accion) Then
            filasMarcadas = filasMarcadas + 1
            datos.Cells(fila, colFecha).Value = fechaHoy
        End If
    Next fila

    If filasMarcadas = 0 Then
        MsgBox "No hay filas marcadas para Actualizar, Insertar o Eliminar.", vbInformation
        GoTo Salida
    End If

    datos.Calculate

    Set cmdUpdate = New ADODB.Command
    With cmdUpdate
        .ActiveConnection = cn
        .CommandType = adCmdText
        sql = "UPDATE Bd_Ppto SET "
        For i = LBound(campos) To UBound(campos)
            If i > LBound(campos) Then sql = sql & ", "
            sql = sql & "[" & Replace(campos(i), "]", "]]" ) & "]= ?"
        Next i
        .CommandText = sql & " WHERE [Id1]=?"
        For i = LBound(campos) To UBound(campos)
            AgregarParametro cmdUpdate, tipoCampos(i), tamCampos(i), precisionCampos(i), escalaCampos(i)
        Next i
        AgregarParametro cmdUpdate, tipoId1, tamId1, precisionId1, escalaId1
    End With

    Set cmdInsert = New ADODB.Command
    With cmdInsert
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = "INSERT INTO Bd_Ppto ([Id]) VALUES (?)"
        AgregarParametro cmdInsert, tipoId, tamId, precisionId, escalaId
    End With

    Set cmdDelete = New ADODB.Command
    With cmdDelete
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = "DELETE FROM Bd_Ppto WHERE [Id1]=?"
        AgregarParametro cmdDelete, tipoId1, tamId1, precisionId1, escalaId1
    End With

    Set cmdAux = New ADODB.Command
    cmdAux.ActiveConnection = cn
    cmdAux.CommandType = adCmdText

    loteActual = 1
    filasEnLote = 0
    ReDim filasLote(1 To TAMANO_LOTE)
    ReDim accionesLote(1 To TAMANO_LOTE)
    ReDim idsLote(1 To TAMANO_LOTE)

    For fila = 1 To datos.Rows.Count
        accion = CStr(datos.Cells(fila, colAccion).Value)

        If EsAccionProcesable(accion) Then
            filaError = datos.Rows(fila).Row
            Application.StatusBar = "Procesando lote " & loteActual & " - " & (filasProcesadas + 1) & "/" & filasMarcadas

            If Not transaccionIniciada Then
                cn.BeginTrans
                transaccionIniciada = True
                filasEnLote = 0
                ReDim filasLote(1 To TAMANO_LOTE)
                ReDim accionesLote(1 To TAMANO_LOTE)
                ReDim idsLote(1 To TAMANO_LOTE)
            End If

            If StrComp(accion, "Eliminar", vbTextCompare) = 0 Then
                idAsignado = datos.Cells(fila, colIdBD).Value
                If IsError(idAsignado) Or IsNull(idAsignado) Or Len(CStr(idAsignado)) = 0 Then
                    Err.Raise vbObjectError + 1001, , "Falta ID_ActualizarBD para eliminar el registro."
                End If

                On Error GoTo ErrorParametroIdentificador
                cmdDelete.Parameters(0).Value = ConvertirValorParametro(idAsignado, tipoId1)
                On Error GoTo ManejarError

            Else
                If StrComp(accion, "Insertar", vbTextCompare) = 0 Then
                    valor = ValorCelda(datos.Cells(fila, colIdExcel), tipoId)
                    On Error GoTo ErrorParametroIdentificador
                    cmdInsert.Parameters(0).Value = ConvertirValorParametro(valor, tipoId)
                    On Error GoTo ManejarError
                    cmdInsert.Execute , , adExecuteNoRecords

                    cmdAux.CommandText = "SELECT @@Identity"
                    Set rs = cmdAux.Execute
                    If rs.EOF Or IsNull(rs.Fields(0).Value) Then
                        Err.Raise vbObjectError + 1002, , "Access no devolvió el Id1 del registro recién insertado."
                    End If
                    idAsignado = rs.Fields(0).Value
                    rs.Close
                    Set rs = Nothing
                Else
                    idAsignado = datos.Cells(fila, colIdBD).Value
                    If IsError(idAsignado) Or IsNull(idAsignado) Or Len(CStr(idAsignado)) = 0 Then
                        Err.Raise vbObjectError + 1003, , "Falta ID_ActualizarBD para actualizar el registro."
                    End If
                End If

                For i = LBound(campos) To UBound(campos)
                    valor = ValorCelda(datos.Cells(fila, colCampos(i)), tipoCampos(i))
                    AjustarTamParametro cmdUpdate.Parameters(i - 1), valor, tipoCampos(i), tamCampos(i)

                    On Error GoTo ErrorParametroCampo
                    cmdUpdate.Parameters(i - 1).Value = ConvertirValorParametro(valor, tipoCampos(i))
                    On Error GoTo ManejarError
                Next i

                On Error GoTo ErrorParametroIdentificador
                cmdUpdate.Parameters(UBound(campos)).Value = ConvertirValorParametro(idAsignado, tipoId1)
                On Error GoTo ManejarError
                cmdUpdate.Execute , , adExecuteNoRecords
            End If

            filasEnLote = filasEnLote + 1
            filasLote(filasEnLote) = fila
            accionesLote(filasEnLote) = accion
            idsLote(filasEnLote) = idAsignado

            If filasEnLote >= TAMANO_LOTE Then
                cn.CommitTrans
                transaccionIniciada = False

                For i = 1 To filasEnLote
                    If StrComp(accionesLote(i), "Eliminar", vbTextCompare) = 0 Then
                        datos.Cells(filasLote(i), colAccion).Value = "Este renglón fue eliminado de la BD"
                    ElseIf StrComp(accionesLote(i), "Insertar", vbTextCompare) = 0 Then
                        datos.Cells(filasLote(i), colIdBD).Value = idsLote(i)
                        datos.Cells(filasLote(i), colAccion).Value = "Actualizar-"
                    Else
                        datos.Cells(filasLote(i), colAccion).Value = "Actualizar-"
                    End If
                    filasProcesadas = filasProcesadas + 1
                Next i

                loteActual = loteActual + 1
                filasEnLote = 0
            End If
        End If
    Next fila

    If transaccionIniciada Then
        cn.CommitTrans
        transaccionIniciada = False

        For i = 1 To filasEnLote
            If StrComp(accionesLote(i), "Eliminar", vbTextCompare) = 0 Then
                datos.Cells(filasLote(i), colAccion).Value = "Este renglón fue eliminado de la BD"
            ElseIf StrComp(accionesLote(i), "Insertar", vbTextCompare) = 0 Then
                datos.Cells(filasLote(i), colIdBD).Value = idsLote(i)
                datos.Cells(filasLote(i), colAccion).Value = "Actualizar-"
            Else
                datos.Cells(filasLote(i), colAccion).Value = "Actualizar-"
            End If
            filasProcesadas = filasProcesadas + 1
        Next i
    End If

    cmdAux.CommandText = "UPDATE Bd_Ppto SET [Variacion]=NULL"
    cmdAux.Execute , , adExecuteNoRecords
    cmdAux.CommandText = "UPDATE Bd_Ppto SET [Vencidos]='En Ejecucion'"
    cmdAux.Execute , , adExecuteNoRecords
    cmdAux.CommandText = "UPDATE Bd_Ppto SET [Vencidos]='Vencidos' WHERE [Concepto_Corto]='Ped S/R' AND [Fecha Termino Pedido] < Now()"
    cmdAux.Execute , , adExecuteNoRecords
    cmdAux.CommandText = "UPDATE Bd_Ppto SET [AnalisiForm_Azul]=NULL"
    cmdAux.Execute , , adExecuteNoRecords
    cmdAux.CommandType = adCmdStoredProc
    cmdAux.CommandText = "Stp_2026_Actualiza_AnalizaForm_Azul"
    cmdAux.Execute , , adExecuteNoRecords

    MsgBox "Actualización terminada. Renglones procesados: " & filasProcesadas & ". Tiempo: " & Format(Now - tiempoInicio, "hh:mm:ss"), vbInformation
    GoTo Salida

ErrorParametroCampo:
    numeroError = Err.Number
    descripcionError = "Campo: " & campos(i) & "; tipo ADO: " & tipoCampos(i) & "; valor: " & DescribirValorParametro(valor) & "; " & Err.Description
    filaError = datos.Rows(fila).Row
    Err.Raise vbObjectError + 1020, , descripcionError

ErrorParametroIdentificador:
    numeroError = Err.Number
    descripcionError = "Identificador Id1/Id; tipo ADO: " & IIf(accion = "Insertar", tipoId, tipoId1) & "; valor: " & DescribirValorParametro(idAsignado) & "; " & Err.Description
    filaError = datos.Rows(fila).Row
    Err.Raise vbObjectError + 1021, , descripcionError

ManejarError:
    numeroError = Err.Number
    descripcionError = Err.Description

    RegistrarLog "Error en fila " & filaError & " - " & descripcionError & " (Número: " & numeroError & ") - Lote: " & loteActual & " - Tiempo: " & Now

    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> adStateClosed Then rs.Close
    If transaccionIniciada Then cn.RollbackTrans

    If estadoAplicacionGuardado Then
        Application.Calculation = calcAnterior
        Application.ScreenUpdating = pantallaAnterior
        Application.EnableEvents = eventosAnteriores
    End If

    Application.StatusBar = False
    MsgBox "Error al procesar ActualizarBD" & IIf(filaError > 0, " en la fila de Excel " & filaError, "") & ": " & descripcionError & " (" & numeroError & ")." & vbCrLf & "Se revirtieron los cambios del lote actual.", vbCritical
    GoTo CerrarObjetos

Salida:
    On Error Resume Next
    If estadoAplicacionGuardado Then
        Application.Calculation = calcAnterior
        Application.ScreenUpdating = pantallaAnterior
        Application.EnableEvents = eventosAnteriores
        datos.Calculate
    End If
    Application.StatusBar = False

CerrarObjetos:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> adStateClosed Then rs.Close
    If Not cn Is Nothing Then If cn.State = adStateOpen Then cn.Close
    Set cmdUpdate = Nothing
    Set cmdInsert = Nothing
    Set cmdDelete = Nothing
    Set cmdAux = Nothing
    Set rs = Nothing
    Set cn = Nothing
    Set datos = Nothing
    Set lo = Nothing
End Sub

Private Sub RegistrarLog(ByVal texto As String)
    Dim rutaLog As String
    Dim nArchivo As Integer

    On Error Resume Next
    If Len(ThisWorkbook.Path) = 0 Then Exit Sub

    rutaLog = ThisWorkbook.Path & "\ActualizarBD_Log.txt"
    nArchivo = FreeFile
    Open rutaLog For Append As #nArchivo
    Print #nArchivo, Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & texto
    Close #nArchivo
End Sub

Private Sub CargarCamposActualizarBD(ByRef campos() As String)
    campos(1) = "Id"
    campos(2) = "Version"
    campos(3) = "Fecha Integracion"
    campos(4) = "Identificador GMM"
    campos(5) = "POT"
    campos(6) = "VERSION2"
    campos(7) = "Techo_Adicional"
    campos(8) = "TOTAL"
    campos(9) = "Concepto"
    campos(10) = "Ejecutor"
    campos(11) = "Centro_Gestor"
    campos(12) = "Fondo"
    campos(13) = "Programa_Presupuestario"
    campos(14) = "E_PEP_6nivel"
    campos(15) = "Desc_ProgPresup"
    campos(16) = "Posicion Presupuestaria"
    campos(17) = "Actividaad GAE"
    campos(18) = "Moneda"
    campos(19) = "ContratoOrigen"
    campos(20) = "Contrato"
    campos(21) = "MODO CONTRAT"
    campos(22) = "Estado Contrato"
    campos(23) = "Politica_de_Pago"
    campos(24) = "Fecha de inicio Contrato"
    campos(25) = "Fecha de termino Contrato"
    campos(26) = "Desc_Contrato"
    campos(27) = "Compania"
    campos(28) = "CIA_ABREV"
    campos(29) = "Reserva_Pptal (Contrato)"
    campos(30) = "Folio_SIT"
    campos(31) = "Folio_RAF_AF"
    campos(32) = "FAO (F_Azul)"
    campos(33) = "ID_AsignadoFA"
    campos(34) = "Reserva_Pos_Asignada"
    campos(35) = "SolPed"
    campos(36) = "Pedido"
    campos(37) = "Fecha Asig FA"
    campos(38) = "FA (F_Azul)"
    campos(39) = "Fecha Inicio Pedido"
    campos(40) = "Fecha Termino Pedido"
    campos(41) = "ARES anuencia"
    campos(42) = "PREINICIATIVA/ SPYA"
    campos(43) = "Descripción_de_la_actividad"
    campos(44) = "Estatus de la actividad formato K"
    campos(45) = "Tipo de mantenimiento"
    campos(46) = "Tipo de equipo"
    campos(47) = "EQUIPO"
    campos(48) = "ENTIDAD"
    campos(49) = "INSTALACION"
    campos(50) = "COMPLEJO"
    campos(51) = "Supervisor Técnico"
    campos(52) = "Nombre_Corto"
    campos(53) = "Supervisor Administrativo"
    campos(54) = "DENE"
    campos(55) = "DFEB"
    campos(56) = "DMAR"
    campos(57) = "DABR"
    campos(58) = "DMAY"
    campos(59) = "DJUN"
    campos(60) = "DJUL"
    campos(61) = "DAGO"
    campos(62) = "DSEP"
    campos(63) = "DOCT"
    campos(64) = "DNOV"
    campos(65) = "DDIC"
    campos(66) = "DTotal"
    campos(67) = "Costo de la actividad"
    campos(68) = "OBSERVACIONES"
    campos(69) = "OBSERVACIONS ADICIONALES"
    campos(70) = "Comentarios"
    campos(71) = "Grupo_Concepto"
    campos(72) = "Concepto_Corto"
    campos(73) = "Orden_Grafica"
    campos(74) = "TIPO_MANTTO"
    campos(75) = "Tipo_Actividad"
    campos(76) = "INSTALACN"
    campos(77) = "Superintendencia"
    campos(78) = "Concepto_Cto"
    campos(79) = "Año"
    campos(80) = "Area Administradora del Contrato"
    campos(81) = "Fecha_Recepcion"
    campos(82) = "Orden_Grupo"
    campos(83) = "Familia"
    campos(84) = "TIPO_MANTTO2"
    campos(85) = "Plazo_Entrega"
    campos(86) = "Capas"
    campos(87) = "Grupo"
    campos(88) = "SubGrupo"
    campos(89) = "Tipo_Equipo"
    campos(90) = "Fecha_Facturacion"
    campos(91) = "EventoFAO"
    campos(92) = "EventoFaAsig"
    campos(93) = "SemanaCOPADE"
    campos(94) = "SemanaFactura"
    campos(95) = "EventoFaSol"
    campos(96) = "Concepto_Resumen"
    campos(97) = "Form_Adecuado"
    campos(98) = "Responsable"
    campos(99) = "Fecha_Actualizacion"
    campos(100) = "ActividadCorta"
    campos(101) = "Obs_FAO"
    campos(102) = "Obs_FA"
    campos(103) = "Obs_Vencidos"
    campos(104) = "Avance"
    campos(105) = "Estatus_Actividad"
End Sub

Private Function ColumnaTabla(ByVal tabla As ListObject, ByVal nombre As String) As Long
    Dim lc As ListColumn
    For Each lc In tabla.ListColumns
        If StrComp(lc.Name, nombre, vbTextCompare) = 0 Then
            ColumnaTabla = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Function EsAccionProcesable(ByVal accion As String) As Boolean
    EsAccionProcesable = (StrComp(accion, "Actualizar", vbTextCompare) = 0 Or _
                          StrComp(accion, "Insertar", vbTextCompare) = 0 Or _
                          StrComp(accion, "Eliminar", vbTextCompare) = 0)
End Function

Private Sub AgregarParametro(ByVal cmd As ADODB.Command, ByVal tipo As Long, _
                             ByVal tamano As Long, ByVal precision As Byte, ByVal escala As Byte)
    Dim p As ADODB.Parameter
    Dim longitud As Long

    If EsTipoVariable(tipo) Then
        longitud = tamano
        If longitud < 1 Then longitud = 1
        If tipo = adLongVarWChar Or tipo = adLongVarChar Or tipo = adLongVarBinary Then
            If longitud < 65535 Then longitud = 65535
        End If
        Set p = cmd.CreateParameter("", tipo, adParamInput, longitud)
    Else
        Set p = cmd.CreateParameter("", tipo, adParamInput)
    End If

    If tipo = adNumeric Or tipo = adDecimal Then
        If precision > 0 Then p.Precision = precision
        p.NumericScale = escala
    End If

    cmd.Parameters.Append p
End Sub

Private Function EsTipoVariable(ByVal tipo As Long) As Boolean
    Select Case tipo
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, _
             adBinary, adVarBinary, adLongVarBinary
            EsTipoVariable = True
    End Select
End Function

Private Sub AjustarTamParametro(ByVal p As ADODB.Parameter, ByVal valor As Variant, _
                                ByVal tipo As Long, ByVal tamanoDefinido As Long)
    Dim requerido As Long

    If EsTipoVariable(tipo) Then
        requerido = tamanoDefinido
        If requerido < 1 Then requerido = 1

        If Not IsNull(valor) And Not IsEmpty(valor) Then
            If VarType(valor) = vbString Then
                If Len(CStr(valor)) > requerido Then requerido = Len(CStr(valor))
            End If
        End If

        p.Size = requerido
    End If
End Sub

Private Function ConvertirValorParametro(ByVal valor As Variant, ByVal tipo As Long) As Variant
    If IsError(valor) Or IsNull(valor) Or IsEmpty(valor) Then
        ConvertirValorParametro = Null
        Exit Function
    End If

    If VarType(valor) = vbString And Len(CStr(valor)) = 0 Then
        Select Case tipo
            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar
                ConvertirValorParametro = ""
            Case Else
                ConvertirValorParametro = Null
        End Select
        Exit Function
    End If

    Select Case tipo
        Case adTinyInt, adSmallInt, adInteger, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt
            ConvertirValorParametro = CLng(valor)
        Case adBigInt, adUnsignedBigInt
            ConvertirValorParametro = CDec(valor)
        Case adNumeric, adDecimal
            ConvertirValorParametro = CDec(valor)
        Case adCurrency
            ConvertirValorParametro = CCur(valor)
        Case adDouble
            ConvertirValorParametro = CDbl(valor)
        Case adSingle
            ConvertirValorParametro = CSng(valor)
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            ConvertirValorParametro = CDate(valor)
        Case adBoolean
            ConvertirValorParametro = CBool(valor)
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar
            ConvertirValorParametro = CStr(valor)
        Case Else
            ConvertirValorParametro = valor
    End Select
End Function

Private Function DescribirValorParametro(ByVal valor As Variant) As String
    If IsError(valor) Then
        DescribirValorParametro = "<Error de Excel>"
    ElseIf IsNull(valor) Then
        DescribirValorParametro = "<Null>"
    ElseIf IsEmpty(valor) Then
        DescribirValorParametro = "<Empty>"
    Else
        DescribirValorParametro = "[" & CStr(valor) & "] (VarType=" & CStr(VarType(valor)) & ")"
    End If
End Function

Private Function ValorCelda(ByVal celda As Range, ByVal tipo As Long) As Variant
    Dim v As Variant
    v = celda.Value

    If IsError(v) Then Err.Raise vbObjectError + 1010, , "La celda " & celda.Address(False, False) & " contiene un error de Excel."

    If IsEmpty(v) Or (VarType(v) = vbString And Len(v) = 0) Then
        If tipo = adChar Or tipo = adVarChar Or tipo = adLongVarChar Or _
           tipo = adWChar Or tipo = adVarWChar Or tipo = adLongVarWChar Then
            ValorCelda = ""
        Else
            ValorCelda = Null
        End If
    Else
        ValorCelda = v
    End If
End Function
