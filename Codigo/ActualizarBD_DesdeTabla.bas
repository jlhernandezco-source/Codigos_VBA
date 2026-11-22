Attribute VB_Name = "ActualizarBD_DesdeTabla"
Option Explicit

Private Const TAMANO_LOTE As Long = 20

Public Sub ActualizarBD_DesdeTabla()
    Dim lo As ListObject
    Dim datos As Range
    Dim valores As Variant
    Dim campos As Variant
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim cmdActualizar As ADODB.Command
    Dim cmdInsertar As ADODB.Command
    Dim cmdEliminar As ADODB.Command
    Dim cmdIdentidad As ADODB.Command
    Dim indicesCampo() As Long
    Dim columnasExcel() As Long
    Dim tiposCampo() As Long
    Dim tamanosCampo() As Long
    Dim precisionesCampo() As Byte
    Dim escalasCampo() As Byte
    Dim filasLote() As Long
    Dim accionesLote() As String
    Dim idsNuevos() As Variant
    Dim fechasLote() As Date
    Dim colAccion As Long, colIdBD As Long, colFecha As Long
    Dim tipoId As Long, tamanoId As Long
    Dim precisionId As Byte, escalaId As Byte
    Dim fila As Long, i As Long, indiceCampo As Long
    Dim cantidadAcciones As Long, cantidadProcesada As Long, filasEnLote As Long
    Dim filaError As Long, loteActual As Long
    Dim accion As String, campoError As String, sql As String
    Dim idRegistro As Variant, idNuevo As Variant, valor As Variant
    Dim fechaOperacion As Date
    Dim transaccionIniciada As Boolean
    Dim loteConfirmado As Boolean, loteRevertido As Boolean
    Dim calcAnterior As XlCalculation
    Dim pantallaAnterior As Boolean, eventosAnteriores As Boolean
    Dim statusAnterior As Variant, estadoAplicacionGuardado As Boolean
    Dim numeroError As Long, detalleError As String
    Dim filasTotales As Long, indiceId As Long, registrosAfectados As Long
    Dim columnasFaltantes As String

    On Error GoTo ManejarError

    On Error Resume Next
    Set lo = ActiveCell.ListObject
    On Error GoTo ManejarError
    If lo Is Nothing Then
        MsgBox "Selecciona primero una celda dentro de la tabla de Excel que deseas sincronizar.", vbExclamation
        Exit Sub
    End If

    Set datos = lo.DataBodyRange
    If datos Is Nothing Then
        MsgBox "La tabla de Excel no contiene filas para procesar.", vbInformation
        Exit Sub
    End If

    colAccion = ColumnaTabla(lo, "AccionEnBD")
    colIdBD = ColumnaTabla(lo, "ID_ActualizarBD")
    colFecha = ColumnaTabla(lo, "Fecha_Actualizacion")
    If colAccion = 0 Or colIdBD = 0 Or colFecha = 0 Then
        MsgBox "A la tabla le falta una columna requerida: AccionEnBD, ID_ActualizarBD o Fecha_Actualizacion.", vbExclamation
        Exit Sub
    End If

    valores = datos.Value2
    For fila = 1 To UBound(valores, 1)
        If EsAccionProcesable(AccionEnValor(valores(fila, colAccion))) Then
            cantidadAcciones = cantidadAcciones + 1
        End If
    Next fila

    If cantidadAcciones = 0 Then
        MsgBox "No hay filas marcadas como Actualizar, Insertar o Eliminar.", vbInformation
        Exit Sub
    End If

    campos = CamposActualizarBD()
    ReDim columnasExcel(LBound(campos) To UBound(campos))
    For i = LBound(campos) To UBound(campos)
        columnasExcel(i) = ColumnaTabla(lo, CStr(campos(i)))
        If columnasExcel(i) = 0 Then
            columnasFaltantes = columnasFaltantes & vbCrLf & "- " & CStr(campos(i))
        End If
    Next i
    If Len(columnasFaltantes) > 0 Then
        MsgBox "Faltan estas columnas de datos en la tabla de Excel:" & _
               columnasFaltantes, vbExclamation
        Exit Sub
    End If

    ' Se reutiliza el inicializador de conexión definido por el proyecto original.
    Set cn = New ADODB.Connection
    AbreConexion cn

    ReDim indicesCampo(LBound(campos) To UBound(campos))
    ReDim tiposCampo(LBound(campos) To UBound(campos))
    ReDim tamanosCampo(LBound(campos) To UBound(campos))
    ReDim precisionesCampo(LBound(campos) To UBound(campos))
    ReDim escalasCampo(LBound(campos) To UBound(campos))

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM [Bd_Ppto] WHERE 1=0", cn, adOpenForwardOnly, adLockReadOnly

    With rs.Fields("Id1")
        tipoId = .Type
        tamanoId = CLng(.DefinedSize)
        precisionId = .Precision
        escalaId = .NumericScale
    End With

    indiceCampo = 0
    For i = LBound(campos) To UBound(campos)
        indiceCampo = indiceCampo + 1
        indicesCampo(indiceCampo) = i
        With rs.Fields(CStr(campos(i)))
            tiposCampo(indiceCampo) = .Type
            tamanosCampo(indiceCampo) = CLng(.DefinedSize)
            precisionesCampo(indiceCampo) = .Precision
            escalasCampo(indiceCampo) = .NumericScale
        End With
    Next i
    filasTotales = indiceCampo
    rs.Close
    Set rs = Nothing

    Set cmdActualizar = New ADODB.Command
    With cmdActualizar
        .ActiveConnection = cn
        .CommandType = adCmdText
        sql = "UPDATE [Bd_Ppto] SET "
        For i = 1 To filasTotales
            If i > 1 Then sql = sql & ", "
            sql = sql & "[" & Replace(CStr(campos(indicesCampo(i))), "]", "]]") & "]=?"
        Next i
        .CommandText = sql & " WHERE [Id1]=?"
        For i = 1 To filasTotales
            AgregarParametro cmdActualizar, tiposCampo(i), tamanosCampo(i), _
                             precisionesCampo(i), escalasCampo(i)
        Next i
        AgregarParametro cmdActualizar, tipoId, tamanoId, precisionId, escalaId
    End With

    Set cmdInsertar = New ADODB.Command
    With cmdInsertar
        .ActiveConnection = cn
        .CommandType = adCmdText
        sql = "INSERT INTO [Bd_Ppto] ("
        For i = 1 To filasTotales
            If i > 1 Then sql = sql & ", "
            sql = sql & "[" & Replace(CStr(campos(indicesCampo(i))), "]", "]]") & "]"
        Next i
        sql = sql & ") VALUES ("
        For i = 1 To filasTotales
            If i > 1 Then sql = sql & ", "
            sql = sql & "?"
        Next i
        .CommandText = sql & ")"
        For i = 1 To filasTotales
            AgregarParametro cmdInsertar, tiposCampo(i), tamanosCampo(i), _
                             precisionesCampo(i), escalasCampo(i)
        Next i
    End With

    Set cmdEliminar = New ADODB.Command
    With cmdEliminar
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = "DELETE FROM [Bd_Ppto] WHERE [Id1]=?"
        AgregarParametro cmdEliminar, tipoId, tamanoId, precisionId, escalaId
    End With

    Set cmdIdentidad = New ADODB.Command
    With cmdIdentidad
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = "SELECT @@Identity"
    End With

    ReDim filasLote(1 To TAMANO_LOTE)
    ReDim accionesLote(1 To TAMANO_LOTE)
    ReDim idsNuevos(1 To TAMANO_LOTE)
    ReDim fechasLote(1 To TAMANO_LOTE)

    calcAnterior = Application.Calculation
    pantallaAnterior = Application.ScreenUpdating
    eventosAnteriores = Application.EnableEvents
    statusAnterior = Application.StatusBar
    estadoAplicacionGuardado = True
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    loteActual = 1
    For fila = 1 To UBound(valores, 1)
        accion = AccionEnValor(valores(fila, colAccion))
        If EsAccionProcesable(accion) Then
            filaError = datos.Rows(fila).Row
            Application.StatusBar = "Procesando registro " & (cantidadProcesada + 1) & _
                                    " de " & cantidadAcciones & " (fila " & filaError & ")..."
            DoEvents

            If Not transaccionIniciada Then
                cn.BeginTrans
                transaccionIniciada = True
                filasEnLote = 0
            End If

            fechaOperacion = Now
            idNuevo = Empty
            If StrComp(accion, "Insertar", vbTextCompare) = 0 Then
                For i = 1 To filasTotales
                    indiceId = indicesCampo(i)
                    campoError = CStr(campos(indiceId))
                    valor = ValorCampo(valores(fila, columnasExcel(indiceId)), _
                                       tiposCampo(i), campoError)
                    If StrComp(campoError, "Fecha_Actualizacion", vbTextCompare) = 0 Then
                        valor = fechaOperacion
                    End If
                    AjustarTamParametro cmdInsertar.Parameters(i - 1), valor, _
                                        tiposCampo(i), tamanosCampo(i)
                    cmdInsertar.Parameters(i - 1).Value = ConvertirValorParametro(valor, tiposCampo(i))
                Next i
                cmdInsertar.Execute , , adExecuteNoRecords

                Set rs = cmdIdentidad.Execute
                If rs.EOF Or IsNull(rs.Fields(0).Value) Then
                    Err.Raise vbObjectError + 1100, , _
                              "Access no devolvió el identificador generado para el nuevo registro."
                End If
                idNuevo = rs.Fields(0).Value
                rs.Close
                Set rs = Nothing
            Else
                idRegistro = valores(fila, colIdBD)
                If IsError(idRegistro) Then
                    Err.Raise vbObjectError + 1101, , _
                              "ID_ActualizarBD contiene un error de Excel."
                End If
                If IsNull(idRegistro) Or IsEmpty(idRegistro) Then
                    Err.Raise vbObjectError + 1101, , _
                              "Falta un valor válido en ID_ActualizarBD para esta operación."
                End If
                If VarType(idRegistro) = vbString Then
                    If Len(Trim$(CStr(idRegistro))) = 0 Then
                        Err.Raise vbObjectError + 1101, , _
                                  "Falta un valor válido en ID_ActualizarBD para esta operación."
                    End If
                End If

                If StrComp(accion, "Eliminar", vbTextCompare) = 0 Then
                    cmdEliminar.Parameters(0).Value = ConvertirValorParametro(idRegistro, tipoId)
                    cmdEliminar.Execute registrosAfectados, , adExecuteNoRecords
                    If registrosAfectados = 0 Then
                        Err.Raise vbObjectError + 1102, , _
                                  "No se encontró en Access el Id indicado para eliminar."
                    End If
                Else
                    For i = 1 To filasTotales
                        indiceId = indicesCampo(i)
                        campoError = CStr(campos(indiceId))
                        valor = ValorCampo(valores(fila, columnasExcel(indiceId)), _
                                           tiposCampo(i), campoError)
                        If StrComp(campoError, "Fecha_Actualizacion", vbTextCompare) = 0 Then
                            valor = fechaOperacion
                        End If
                        AjustarTamParametro cmdActualizar.Parameters(i - 1), valor, _
                                            tiposCampo(i), tamanosCampo(i)
                        cmdActualizar.Parameters(i - 1).Value = ConvertirValorParametro(valor, tiposCampo(i))
                    Next i
                    cmdActualizar.Parameters(filasTotales).Value = _
                        ConvertirValorParametro(idRegistro, tipoId)
                    cmdActualizar.Execute registrosAfectados, , adExecuteNoRecords
                    If registrosAfectados = 0 Then
                        Err.Raise vbObjectError + 1103, , _
                                  "No se encontró en Access el Id indicado para actualizar."
                    End If
                End If
            End If

            filasEnLote = filasEnLote + 1
            filasLote(filasEnLote) = fila
            accionesLote(filasEnLote) = accion
            idsNuevos(filasEnLote) = idNuevo
            fechasLote(filasEnLote) = fechaOperacion
            cantidadProcesada = cantidadProcesada + 1

            If filasEnLote = TAMANO_LOTE Then
                cn.CommitTrans
                transaccionIniciada = False
                loteConfirmado = True
                ConfirmarLote datos, filasLote, accionesLote, idsNuevos, fechasLote, _
                              filasEnLote, colAccion, colIdBD, colFecha
                loteConfirmado = False
                loteActual = loteActual + 1
            End If
        End If
    Next fila

    If transaccionIniciada Then
        cn.CommitTrans
        transaccionIniciada = False
        loteConfirmado = True
        ConfirmarLote datos, filasLote, accionesLote, idsNuevos, fechasLote, _
                      filasEnLote, colAccion, colIdBD, colFecha
        loteConfirmado = False
    End If

    MsgBox "Proceso terminado correctamente. Registros procesados: " & cantidadProcesada & ".", vbInformation
    GoTo Salida

ManejarError:
    numeroError = Err.Number
    detalleError = Err.Description
    RegistrarLog "Error en fila " & filaError & ", acción " & accion & _
                 ", campo " & campoError & ", lote " & loteActual & _
                 ", número " & numeroError & ": " & detalleError

    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> adStateClosed Then rs.Close
    If transaccionIniciada Then
        loteRevertido = True
        cn.RollbackTrans
        transaccionIniciada = False
    End If
    If estadoAplicacionGuardado Then
        Application.Calculation = calcAnterior
        Application.ScreenUpdating = pantallaAnterior
        Application.EnableEvents = eventosAnteriores
        Application.StatusBar = statusAnterior
    End If
    If filaError > 0 Then
        MsgBox "No se pudo " & AccionEnFrase(accion) & " el registro de la fila " & filaError & _
               " de Excel." & IIf(Len(campoError) > 0, " Revisa el campo '" & campoError & "'.", "") & _
               vbCrLf & "Detalle: " & detalleError & vbCrLf & _
               IIf(loteRevertido, _
                   "Se deshizo el lote actual; los lotes confirmados anteriormente se conservaron.", _
                   IIf(loteConfirmado, _
                       "El lote quedó guardado en Access, pero no se pudo actualizar la confirmación en Excel.", _
                       "Los lotes confirmados anteriormente se conservaron.")), vbCritical
    Else
        MsgBox "No se pudo iniciar o preparar la actualización de la tabla." & vbCrLf & _
               "Detalle: " & detalleError, vbCritical
    End If
    GoTo CerrarObjetos

Salida:
    On Error Resume Next
    If estadoAplicacionGuardado Then
        Application.Calculation = calcAnterior
        Application.ScreenUpdating = pantallaAnterior
        Application.EnableEvents = eventosAnteriores
        Application.StatusBar = statusAnterior
    End If

CerrarObjetos:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> adStateClosed Then rs.Close
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
    Set cmdActualizar = Nothing
    Set cmdInsertar = Nothing
    Set cmdEliminar = Nothing
    Set cmdIdentidad = Nothing
    Set rs = Nothing
    Set cn = Nothing
    Set datos = Nothing
    Set lo = Nothing
End Sub

Private Sub ConfirmarLote(ByVal datos As Range, ByRef filasLote() As Long, _
                          ByRef accionesLote() As String, ByRef idsNuevos() As Variant, _
                          ByRef fechasLote() As Date, ByVal filasEnLote As Long, _
                          ByVal colAccion As Long, ByVal colIdBD As Long, ByVal colFecha As Long)
    Dim i As Long
    For i = 1 To filasEnLote
        datos.Cells(filasLote(i), colFecha).Value = fechasLote(i)
        If StrComp(accionesLote(i), "Eliminar", vbTextCompare) = 0 Then
            datos.Cells(filasLote(i), colAccion).Value = "Este renglón fue eliminado de la BD"
        Else
            If StrComp(accionesLote(i), "Insertar", vbTextCompare) = 0 Then
                datos.Cells(filasLote(i), colIdBD).Value = idsNuevos(i)
            End If
            datos.Cells(filasLote(i), colAccion).Value = "Actualizar-"
        End If
    Next i
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
    Dim columna As ListColumn
    For Each columna In tabla.ListColumns
        If StrComp(columna.Name, nombre, vbTextCompare) = 0 Then
            ColumnaTabla = columna.Index
            Exit Function
        End If
    Next columna
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

Private Function ValorCampo(ByVal valor As Variant, ByVal tipo As Long, _
                            ByVal nombreCampo As String) As Variant
    If IsError(valor) Then
        Err.Raise vbObjectError + 1110, , _
                  "El campo '" & nombreCampo & "' contiene un error de Excel."
    End If
    If IsNull(valor) Or IsEmpty(valor) Or _
       (VarType(valor) = vbString And Len(valor) = 0) Then
        If EsTipoTexto(tipo) Then
            ValorCampo = ""
        Else
            ValorCampo = Null
        End If
    Else
        ValorCampo = valor
    End If
End Function

Private Sub AgregarParametro(ByVal cmd As ADODB.Command, ByVal tipo As Long, _
                             ByVal tamano As Long, ByVal precision As Byte, ByVal escala As Byte)
    Dim parametro As ADODB.Parameter
    Dim longitud As Long
    If EsTipoVariable(tipo) Then
        longitud = tamano
        If longitud < 1 Then longitud = 1
        If tipo = adLongVarWChar Or tipo = adLongVarChar Or tipo = adLongVarBinary Then
            If longitud < 65535 Then longitud = 65535
        End If
        Set parametro = cmd.CreateParameter("", tipo, adParamInput, longitud)
    Else
        Set parametro = cmd.CreateParameter("", tipo, adParamInput)
    End If
    If tipo = adNumeric Or tipo = adDecimal Then
        If precision > 0 Then parametro.Precision = precision
        parametro.NumericScale = escala
    End If
    cmd.Parameters.Append parametro
End Sub

Private Sub AjustarTamParametro(ByVal parametro As ADODB.Parameter, ByVal valor As Variant, _
                                ByVal tipo As Long, ByVal tamanoDefinido As Long)
    Dim requerido As Long
    If EsTipoVariable(tipo) Then
        requerido = tamanoDefinido
        If requerido < 1 Then requerido = 1
        If Not IsNull(valor) And Not IsEmpty(valor) And VarType(valor) = vbString Then
            If Len(CStr(valor)) > requerido Then requerido = Len(CStr(valor))
        End If
        parametro.Size = requerido
    End If
End Sub

Private Function ConvertirValorParametro(ByVal valor As Variant, ByVal tipo As Long) As Variant
    If IsNull(valor) Or IsEmpty(valor) Then
        ConvertirValorParametro = Null
        Exit Function
    End If
    If VarType(valor) = vbString And Len(CStr(valor)) = 0 Then
        If EsTipoTexto(tipo) Then
            ConvertirValorParametro = ""
        Else
            ConvertirValorParametro = Null
        End If
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

Private Function EsTipoTexto(ByVal tipo As Long) As Boolean
    Select Case tipo
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar
            EsTipoTexto = True
    End Select
End Function

Private Function EsTipoVariable(ByVal tipo As Long) As Boolean
    Select Case tipo
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, _
             adBinary, adVarBinary, adLongVarBinary
            EsTipoVariable = True
    End Select
End Function

Private Function AccionEnFrase(ByVal accion As String) As String
    Select Case LCase$(accion)
        Case "actualizar": AccionEnFrase = "actualizar"
        Case "insertar": AccionEnFrase = "insertar"
        Case "eliminar": AccionEnFrase = "eliminar"
        Case Else: AccionEnFrase = "procesar"
    End Select
End Function

Private Sub RegistrarLog(ByVal texto As String)
    Dim rutaLog As String
    Dim numeroArchivo As Integer
    On Error Resume Next
    If Len(ThisWorkbook.Path) = 0 Then Exit Sub
    rutaLog = ThisWorkbook.Path & Application.PathSeparator & "ActualizarBD_Log.txt"
    numeroArchivo = FreeFile
    Open rutaLog For Append As #numeroArchivo
    Print #numeroArchivo, Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & texto
    Close #numeroArchivo
End Sub
