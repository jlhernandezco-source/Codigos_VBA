Attribute VB_Name = "ActualizarBD_Optimizada"
Option Explicit

Private Const TAMANO_LOTE As Long = 20

Public Sub ActualizarBD_Optimizada_Batch()
    Dim campos As Variant
    campos = CamposActualizarBD()

    Dim lo As ListObject
    Dim datos As Range
    Dim valores As Variant
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cmdUpdate As ADODB.Command
    Dim cmdInsert As ADODB.Command
    Dim cmdDelete As ADODB.Command
    Dim cmdAux As ADODB.Command
    Dim i As Long, fila As Long, indiceTrabajo As Long
    Dim nCampos As Long
    Dim colCampos() As Long
    Dim tipoCampos() As Long
    Dim tamCampos() As Long
    Dim precisionCampos() As Byte
    Dim escalaCampos() As Byte
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
    Dim filasTrabajo() As Long
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

    nCampos = UBound(campos) - LBound(campos) + 1
    ReDim colCampos(LBound(campos) To UBound(campos))
    ReDim tipoCampos(LBound(campos) To UBound(campos))
    ReDim tamCampos(LBound(campos) To UBound(campos))
    ReDim precisionCampos(LBound(campos) To UBound(campos))
    ReDim escalaCampos(LBound(campos) To UBound(campos))

    colAccion = ColumnaTabla(lo, "AccionEnBD")
    colIdExcel = ColumnaTabla(lo, "Id")
    colIdBD = ColumnaTabla(lo, "ID_ActualizarBD")
    colFecha = ColumnaTabla(lo, "Fecha_Actualizacion")

    If colAccion = 0 Or colIdExcel = 0 Or colIdBD = 0 Or colFecha = 0 Then
        MsgBox "Falta una columna requerida: AccionEnBD, Id, ID_ActualizarBD o Fecha_Actualizacion.", vbExclamation
        Exit Sub
    End If

    For i = LBound(campos) To UBound(campos)
        colCampos(i) = ColumnaTabla(lo, CStr(campos(i)))
        If colCampos(i) = 0 Then columnaError = columnaError & vbCrLf & "- " & CStr(campos(i))
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
        With rs.Fields(CStr(campos(i)))
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

    ' Una sola lectura de la tabla a memoria para contar las acciones.
    valores = datos.Value2
    For fila = 1 To UBound(valores, 1)
        accion = AccionEnValor(valores(fila, colAccion))
        If EsAccionProcesable(accion) Then
            filasMarcadas = filasMarcadas + 1
            datos.Cells(fila, colFecha).Value = fechaHoy
        End If
    Next fila

    If filasMarcadas = 0 Then
        MsgBox "No hay filas marcadas para Actualizar, Insertar o Eliminar.", vbInformation
        GoTo Salida
    End If

    ' Recalcular y volver a cargar el bloque completo, ya con las fechas asignadas.
    datos.Calculate
    valores = datos.Value2

    ' Se conservan únicamente los índices de filas procesables; no dependemos de filtros
    ' visuales ni de rangos visibles discontinuos.
    filasMarcadas = 0
    For fila = 1 To UBound(valores, 1)
        If EsAccionProcesable(AccionEnValor(valores(fila, colAccion))) Then
            filasMarcadas = filasMarcadas + 1
        End If
    Next fila
    If filasMarcadas = 0 Then GoTo Salida

    ReDim filasTrabajo(1 To filasMarcadas)
    indiceTrabajo = 0
    For fila = 1 To UBound(valores, 1)
        If EsAccionProcesable(AccionEnValor(valores(fila, colAccion))) Then
            indiceTrabajo = indiceTrabajo + 1
            filasTrabajo(indiceTrabajo) = fila
        End If
    Next fila

    Set cmdUpdate = New ADODB.Command
    With cmdUpdate
        .ActiveConnection = cn
        .CommandType = adCmdText
        sql = "UPDATE Bd_Ppto SET "
        For i = LBound(campos) To UBound(campos)
            If i > LBound(campos) Then sql = sql & ", "
            sql = sql & "[" & Replace(CStr(campos(i)), "]", "]]" ) & "]= ?"
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

    For indiceTrabajo = 1 To filasMarcadas
        fila = filasTrabajo(indiceTrabajo)
        accion = AccionEnValor(valores(fila, colAccion))
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
            idAsignado = valores(fila, colIdBD)
            If IsError(idAsignado) Or IsNull(idAsignado) Or IsEmpty(idAsignado) Or Len(CStr(idAsignado)) = 0 Then
                Err.Raise vbObjectError + 1001, , "Falta ID_ActualizarBD para eliminar el registro."
            End If
            On Error GoTo ErrorParametroIdentificador
            cmdDelete.Parameters(0).Value = ConvertirValorParametro(idAsignado, tipoId1)
            On Error GoTo ManejarError
            cmdDelete.Execute , , adExecuteNoRecords
        Else
            If StrComp(accion, "Insertar", vbTextCompare) = 0 Then
                valor = ValorArreglo(valores(fila, colIdExcel), tipoId, lo.ListColumns(colIdExcel).Name)
                On Error GoTo ErrorParametroIdentificador
                cmdInsert.Parameters(0).Value = ConvertirValorParametro(valor, tipoId)
                On Error GoTo ManejarError
                cmdInsert.Execute , , adExecuteNoRecords

                cmdAux.CommandType = adCmdText
                cmdAux.CommandText = "SELECT @@Identity"
                Set rs = cmdAux.Execute
                If rs.EOF Or IsNull(rs.Fields(0).Value) Then
                    Err.Raise vbObjectError + 1002, , "Access no devolvió el Id1 del registro recién insertado."
                End If
                idAsignado = rs.Fields(0).Value
                rs.Close
                Set rs = Nothing
            Else
                idAsignado = valores(fila, colIdBD)
                If IsError(idAsignado) Or IsNull(idAsignado) Or IsEmpty(idAsignado) Or Len(CStr(idAsignado)) = 0 Then
                    Err.Raise vbObjectError + 1003, , "Falta ID_ActualizarBD para actualizar el registro."
                End If
            End If

            For i = LBound(campos) To UBound(campos)
                valor = ValorArreglo(valores(fila, colCampos(i)), tipoCampos(i), CStr(campos(i)))
                AjustarTamParametro cmdUpdate.Parameters(i - LBound(campos)), valor, tipoCampos(i), tamCampos(i)
                On Error GoTo ErrorParametroCampo
                cmdUpdate.Parameters(i - LBound(campos)).Value = ConvertirValorParametro(valor, tipoCampos(i))
                On Error GoTo ManejarError
            Next i

            On Error GoTo ErrorParametroIdentificador
            cmdUpdate.Parameters(nCampos).Value = ConvertirValorParametro(idAsignado, tipoId1)
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
            ConfirmarLote datos, filasLote, accionesLote, idsLote, filasEnLote, colAccion, colIdBD, filasProcesadas
            loteActual = loteActual + 1
            filasEnLote = 0
        End If
    Next indiceTrabajo

    If transaccionIniciada Then
        cn.CommitTrans
        transaccionIniciada = False
        ConfirmarLote datos, filasLote, accionesLote, idsLote, filasEnLote, colAccion, colIdBD, filasProcesadas
    End If

    cmdAux.CommandType = adCmdText
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
    descripcionError = "Campo: " & CStr(campos(i)) & "; tipo ADO: " & tipoCampos(i) & "; valor: " & DescribirValorParametro(valor) & "; " & Err.Description
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
    MsgBox "Error al procesar ActualizarBD" & IIf(filaError > 0, " en la fila de Excel " & filaError, "") & ": " & descripcionError & " (" & numeroError & ")." & vbCrLf & "Se revirtió el lote actual; los lotes confirmados anteriormente no se revierten.", vbCritical
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

Private Sub ConfirmarLote(ByVal datos As Range, ByRef filasLote() As Long, _
                          ByRef accionesLote() As String, ByRef idsLote() As Variant, _
                          ByVal filasEnLote As Long, ByVal colAccion As Long, _
                          ByVal colIdBD As Long, ByRef filasProcesadas As Long)
    Dim i As Long
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

Private Function CamposActualizarBD() As Variant
    Dim campos As String

    campos = "Id|Version|Fecha Integracion|Identificador GMM|POT|VERSION2|Techo_Adicional|"
    campos = campos & "TOTAL|Concepto|Ejecutor|Centro_Gestor|Fondo|Programa_Presupuestario|"
    campos = campos & "E_PEP_6nivel|Desc_ProgPresup|Posicion Presupuestaria|Actividaad GAE|"
    campos = campos & "Moneda|ContratoOrigen|Contrato|MODO CONTRAT|Estado Contrato|"
    campos = campos & "Politica_de_Pago|Fecha de inicio Contrato|Fecha de termino Contrato|"
    campos = campos & "Desc_Contrato|Compania|CIA_ABREV|Reserva_Pptal (Contrato)|Folio_SIT|"
    campos = campos & "Folio_RAF_AF|FAO (F_Azul)|ID_AsignadoFA|Reserva_Pos_Asignada|SolPed|"
    campos = campos & "Pedido|Fecha Asig FA|FA (F_Azul)|Fecha Inicio Pedido|Fecha Termino Pedido|"
    campos = campos & "ARES anuencia|PREINICIATIVA/ SPYA|Descripción_de_la_actividad|"
    campos = campos & "Estatus de la actividad formato K|Tipo de mantenimiento|Tipo de equipo|"
    campos = campos & "EQUIPO|ENTIDAD|INSTALACION|COMPLEJO|Supervisor Técnico|Nombre_Corto|"
    campos = campos & "Supervisor Administrativo|DENE|DFEB|DMAR|DABR|DMAY|DJUN|DJUL|DAGO|"
    campos = campos & "DSEP|DOCT|DNOV|DDIC|DTotal|Costo de la actividad|OBSERVACIONES|"
    campos = campos & "OBSERVACIONS ADICIONALES|Comentarios|Grupo_Concepto|Concepto_Corto|"
    campos = campos & "Orden_Grafica|TIPO_MANTTO|Tipo_Actividad|INSTALACN|Superintendencia|"
    campos = campos & "Concepto_Cto|Año|Area Administradora del Contrato|Fecha_Recepcion|"
    campos = campos & "Orden_Grupo|Familia|TIPO_MANTTO2|Plazo_Entrega|Capas|Grupo|SubGrupo|"
    campos = campos & "Tipo_Equipo|Fecha_Facturacion|EventoFAO|EventoFaAsig|SemanaCOPADE|"
    campos = campos & "SemanaFactura|EventoFaSol|Concepto_Resumen|Form_Adecuado|Responsable|"
    campos = campos & "Fecha_Actualizacion|ActividadCorta|Obs_FAO|Obs_FA|Obs_Vencidos|Avance|"
    campos = campos & "Estatus_Actividad"

    CamposActualizarBD = Split(campos, "|")
End Function

Private Function ColumnaTabla(ByVal tabla As ListObject, ByVal nombre As String) As Long
    Dim lc As ListColumn
    For Each lc In tabla.ListColumns
        If StrComp(lc.Name, nombre, vbTextCompare) = 0 Then
            ColumnaTabla = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Function AccionEnValor(ByVal valor As Variant) As String
    If IsError(valor) Or IsNull(valor) Or IsEmpty(valor) Then Exit Function
    AccionEnValor = CStr(valor)
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
        Case adBigInt, adUnsignedBigInt, adNumeric, adDecimal
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

Private Function ValorArreglo(ByVal valor As Variant, ByVal tipo As Long, _
                              ByVal nombreCampo As String) As Variant
    If IsError(valor) Then
        Err.Raise vbObjectError + 1010, , "El valor del campo " & nombreCampo & " contiene un error de Excel."
    End If
    If IsEmpty(valor) Or (VarType(valor) = vbString And Len(valor) = 0) Then
        If tipo = adChar Or tipo = adVarChar Or tipo = adLongVarChar Or _
           tipo = adWChar Or tipo = adVarWChar Or tipo = adLongVarWChar Then
            ValorArreglo = ""
        Else
            ValorArreglo = Null
        End If
    Else
        ValorArreglo = valor
    End If
End Function
