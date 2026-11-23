Attribute VB_Name = "Formato_2026"
Option Compare Database

Type Tipo_Soportes
    Pos_Graf As Integer
    Nombre_X As String
    Nombre_Btn As String
End Type

Type ProcesoConsulta
    Consulta As String
    Proc_Almacenado As String
    Estructura As String
    EncabezadoVariable As String
    EncabezdoDiap As String
End Type

Sub IntegraFormato_2026(A_Integrar As String)
     Dim Tiempo As Date
    Tiempo = VBA.Now
    
    Set Db = CurrentDb
    
    Dim Pp As New PowerPoint.Application
    Dim Power As PowerPoint.Presentation
    
    Dim Nombre_PP As String
    Dim Num_Diap As Integer
    'Nombre_PP = "Formatito.potx"
    'Nombre_PP = "Formatito2024.potx"
    Nombre_PP = "Plantilla_2026.potx"
        
    Set Power = Pp.Presentations.Open(CurrentProject.Path & "\" & Nombre_PP, msoFalse, msoCTrue)
    
    Power.SaveAs (CurrentProject.Path & "\" & "Informe_2026_GMM-AELT_" & Day(Now()) & "-" & Format(Month(Now()), "00") & "-" & Year(Now()) & "_" & Hour(Now()) & "-" & Minute(Now()) & ".pptx")
   ' Power.Application.WindowState = ppWindowMinimized
    DoEvents
'
    If Variables.QueHago(1) Then
        Call Integra_ResumenAdecuado_2026(Power)
        Call Soporte_Variacion(Power)
        Call Integra_VariacionPrograma_Real(Power)
    End If
    If Variables.QueHago(2) Then Call FormatoAzul_FAO_FA_UnaSolaGraf(Power)
    If Variables.QueHago(3) Then Call Pedidos_Calendarizadosv2(Power)
    
    
    
'    Call Integra_Iniciativa(Power)
'    Call Integra_Resumen(Power)
 '   Call Integra_FormatosAzules_2026(Power)
    
    
    MsgBox "Listo Jorgito" & vbNewLine & " Jalale porque para eso te pagan (a veces)....." & vbNewLine & "Te integre la presentacion en " & Format(Now() - Tiempo, "hh:mm:ss") & vbNewLine & "Tu sabras si valoras lo que hago por ti...."
    
    Power.Application.WindowState = ppWindowMaximized
    Power.Application.Activate
End Sub


Sub Integra_ResumenAdecuado_2026(Pwr As PowerPoint.Presentation)
    Dim Grafica As PowerPoint.Chart
    Dim gChartData As PowerPoint.ChartData
    Dim gWorkBook As excel.Workbook
    Dim gWorkSheet As New excel.Worksheet
    Dim Diapositiva As PowerPoint.Slide
    Dim NumDiapRetorno As Integer
    
    Dim Qdf As DAO.QueryDef
    Dim Rs As DAO.Recordset
    Set Qdf = Db.QueryDefs("Cons_2026_GraficoResumen")
    Qdf.Parameters("Id_Adecuado") = Id_Adecuado
    Qdf.Parameters("Fecha") = Variables.Periodo_Fin
    
    Set Rs = Qdf.OpenRecordset
    
    Set Diapositiva = Pwr.Slides(Variables.No_DiapGrafAdec)
    
    Diapositiva.Shapes("Txt_Periodo").TextFrame.TextRange.Text = "Periodo " & Variables.PeriodoaEvaluar
    Set Grafica = Diapositiva.Shapes("GraficoPrincipal").Chart
    Set gChartData = Grafica.ChartData
    Set gWorkBook = gChartData.Workbook
        
    Dim sql As String
    
    Db.Execute ("UPDATE Bd_Ppto SET SumaPeriodo = " & Meses(1) & ";")
    
    sql = "SELECT Sum([DTotal])/1000000 AS Importe From Bd_Ppto WHERE (((Bd_Ppto.Id)=" & Id_Adecuado & ") AND ((Bd_Ppto.Techo_Adicional)=""Inercial""));"
    
    Dim Progr As Double
    Dim RsProgr As Recordset
    Dim Inicio_Renglon As Integer
    
    Inicio_Renglon = gWorkBook.Worksheets(1).Range("InicioRenglon").Value 'Worksheet.Cells(1, 13).Value
    
    Dim i As Integer
    i = Inicio_Renglon
    
    Set RsProgr = Db.OpenRecordset(sql)
    Progr = RsProgr(0).Value
    'gWorkBook.Worksheets(1).Range("A" & i - 1).Value = "Programado" ' & RetornaMes4(MesFinal) 'Month(Now()))
    
    gWorkBook.Worksheets(1).Range("C" & i).Value = Progr
    gWorkBook.Worksheets(1).Range("C" & i).NumberFormat = "#,##0.00"
    gWorkBook.Worksheets(1).Range("B" & i).Value = 0
    gWorkBook.Worksheets(1).Range("A" & i).Value = "Total"
    Set RsProgr = Nothing
    
    i = i + 1
    'Sql = "SELECT Sum(" & Meses(1) & ")/1000000 AS Programado From Bd_Ppto WHERE (((Bd_Ppto.Id)=" & Id_AdecuadoOficial & ") AND ((Bd_Ppto.Techo_Adicional)=""Inercial""));"
    sql = "SELECT Sum([DTotal]/1000000) FROM Bd_Ppto WHERE ((Id = " & Id_Adecuado & ") AND (Grupo_Concepto = ""Consumo"") AND ((Bd_Ppto.[Fecha Termino Pedido]) <= #" & Variables.Periodo_Fin & "#));"
    Dim Total As Double
        
    Set RsProgr = Db.OpenRecordset(sql)
    Total = RsProgr(0).Value
    
    gWorkBook.Worksheets(1).Range("C" & i).Value = Total
    gWorkBook.Worksheets(1).Range("C" & i).NumberFormat = "#,##0.00"
    gWorkBook.Worksheets(1).Range("B" & i).Value = 0
    gWorkBook.Worksheets(1).Range("A" & i).Value = "Progr. Periodo"
    i = i + 1
    
    gWorkBook.Worksheets(1).Range("C" & i).Formula = "=C3-C5"
    gWorkBook.Worksheets(1).Range("C" & i).NumberFormat = "#,##0.00"
    gWorkBook.Worksheets(1).Range("B" & i).Formula = "=C5"
    gWorkBook.Worksheets(1).Range("A" & i).Value = "Variacion"
    
    sql = "SELECT Sum(Total)/1000000 As TOT From Bd_Ppto WHERE ((Id = " & Id_AdecuadoOficial & ") AND (Concepto_Corto = ""Ejercicio"") AND (Fecha_Facturacion<= #" & Variables.Periodo_Fin & "#));"
    Set RsProgr = Db.OpenRecordset(sql)
    Total = RsProgr(0).Value
    i = i + 1
     gWorkBook.Worksheets(1).Range("C" & i).Value = Total
    gWorkBook.Worksheets(1).Range("C" & i).NumberFormat = "#,##0.00"
    gWorkBook.Worksheets(1).Range("B" & i).Formula = 0
    gWorkBook.Worksheets(1).Range("A" & i).Value = "Facturado Periodo"
    
'        Sql = "SELECT Sum(Total)/1000000 As TOT From Bd_Ppto WHERE ((Id = " & Id_AdecuadoOficial & ") AND (Concepto_Corto = ""Ejercicio""));"
'    Set RsProgr = Db.OpenRecordset(Sql)
'    Dim FactAdic As Double
'    FactAdic = RsProgr(0).Value
'
'     i = i + 1
'     gWorkBook.Worksheets(1).Range("C" & i).Value = Abs(FactAdic - Total)
'    gWorkBook.Worksheets(1).Range("C" & i).NumberFormat = "#,##0.00"
'    gWorkBook.Worksheets(1).Range("B" & i).Formula = "=C5"
'    gWorkBook.Worksheets(1).Range("A" & i).Value = "Facturacion Adicional"

    
    Dim Para_Soportes() As Tipo_Soportes
    ReDim Preserve Para_Soportes(Rs.RecordCount + 1)
    Rs.MoveFirst
     Dim ContReg As Integer
    ContReg = 0
    
    While Not Rs.EOF
        i = i + 1
        ContReg = ContReg + 1
        Para_Soportes(ContReg).Pos_Graf = i - 1
        Para_Soportes(ContReg).Nombre_X = Rs("Concepto_Corto").Value
        gWorkBook.Worksheets(1).Range("A" & i).Value = Rs("Concepto_Corto").Value
        gWorkBook.Worksheets(1).Range("C" & i).Value = Format(Rs("Total").Value, "#,##0.00")
        gWorkBook.Worksheets(1).Range("C" & i).NumberFormat = "#,##0.00"
        If ContReg > 1 Then gWorkBook.Worksheets(1).Range("b" & i).Formula = "=B" & i - 1 & "+c" & i - 1
        
        Rs.MoveNext
    Wend
            
    Grafica.SetSourceData ("=" & gWorkBook.Worksheets(1).Name & "!$A$1:$C$" & i)
    Grafica.Refresh
    
    gWorkBook.Close
    
    For i = 1 To Grafica.SeriesCollection(2).DataLabels.Count
        Grafica.SeriesCollection(2).DataLabels(i).Top = Grafica.SeriesCollection(2).Points(i).Top - (Grafica.SeriesCollection(2).DataLabels(i).Height + Grafica.SeriesCollection(2).DataLabels(i).Height / 2) '- Grafica.Axes(1).Height
'        Grafica.SeriesCollection(2).DataLabels(i).Left = Grafica.SeriesCollection(2).Points(i).Left
        Grafica.SeriesCollection(2).DataLabels(i).Left = (Grafica.SeriesCollection(2).Points(i).Left + (Grafica.SeriesCollection(2).Points(i).Width / 2)) - (Grafica.SeriesCollection(2).DataLabels(i).Width / 2) - (Grafica.SeriesCollection(2).DataLabels(i).Width * 0.13)
    Next
    
    Dim Estruc(3) As Proceso.Tipo_Tabla
    
    Estruc(1).Alineacion = 4
    Estruc(1).Encabezado = "Acreedores"
    Estruc(1).Size_Fuente = 9
    Estruc(1).Ancho = CentimetrosAPuntos(8.1)
    Estruc(1).Fuente = FuenteOficial
    Estruc(1).SeSuma = False
    Estruc(1).Tipo = "Cadena"
    Estruc(1).Total = 0
    
    Estruc(2).Alineacion = 4
    Estruc(2).Encabezado = "Supervisor Tecnico"
    Estruc(2).Size_Fuente = 9
    Estruc(2).Ancho = CentimetrosAPuntos(9)
    Estruc(2).Fuente = FuenteOficial
    Estruc(2).SeSuma = False
    Estruc(2).Tipo = "Cadena"
    Estruc(2).Total = 0
    
    Estruc(3).Alineacion = 3
    Estruc(3).Encabezado = "Monto"
    Estruc(3).Size_Fuente = 9
    Estruc(3).Ancho = CentimetrosAPuntos(2.5)
    Estruc(3).Fuente = FuenteOficial
    Estruc(3).SeSuma = True
    Estruc(3).Tipo = "Numero"
    Estruc(3).Total = 0
    
    If Variables.Soportes Then
        Set Qdf = Db.QueryDefs("Cons_2026_Soporte_GrafPrincipal")
        Qdf.Parameters("Id_Adecuado") = Id_Adecuado
        
        
        Dim Btn_Soporte As PowerPoint.Shape
        
        Dim Num_DiapRetorno As Integer
        Dim TotalDiap As Integer
        
        Dim Link As PowerPoint.ActionSetting
        Dim Qdf2 As DAO.QueryDef
        
        For i = 1 To UBound(Para_Soportes)
            If Para_Soportes(i).Nombre_X = "Facturacion Adicional" Then
                 Set Qdf2 = Db.QueryDefs("Cons_2026_Soporte_GrafPrincipalFactAdic")
                Qdf2.Parameters("Id_Adecuado") = Id_Adecuado
                Qdf2.Parameters("Fecha") = Variables.Periodo_Fin
                Set Rs = Qdf2.OpenRecordset
            Else
                Qdf.Parameters("Grupo_Concepto") = Para_Soportes(i).Nombre_X
                Set Rs = Qdf.OpenRecordset
            End If
            
            
            If Rs.RecordCount > 0 Then
                Diapositiva.Shapes("Btn_Variacion").Duplicate
                Set Btn_Soporte = Diapositiva.Shapes("Btn_Variacion")
                Btn_Soporte.Name = "Btn_" & Para_Soportes(i).Nombre_X
                Btn_Soporte.Top = Grafica.SeriesCollection(2).DataLabels(Para_Soportes(i).Pos_Graf).Top + Diapositiva.Shapes("GraficoPrincipal").Top
                Btn_Soporte.Left = Grafica.SeriesCollection(2).DataLabels(Para_Soportes(i).Pos_Graf).Left + Grafica.SeriesCollection(2).DataLabels(Para_Soportes(i).Pos_Graf).Width + Diapositiva.Shapes("GraficoPrincipal").Left + Btn_Soporte.Width
                
                Num_DiapRetorno = Crea_LaTablaJorgito(Pwr, Estruc, Rs, Variables.No_DiapResumenAcreedores, Diapositiva.SlideNumber, Para_Soportes(i).Nombre_X, False, True, False, 1, "Tbl_Resumen", CentimetrosAPuntos(0.55), CentimetrosAPuntos(2.84), CentimetrosAPuntos(0.52), Variables.Color_Encabezado, Variables.Color_Rng1, Variables.Color_Rng2, Variables.Color_Grupo, TotalDiap, Variables.Fecha_Consumo)
                Set Link = Btn_Soporte.ActionSettings(ppMouseClick)
                
                Link.Action = ppActionHyperlink
                Link.Hyperlink.SubAddress = Num_DiapRetorno
    
                
            End If
        Next
    End If

    Diapositiva.Shapes("Btn_Variacion").Delete
    Set Comandito = Nothing
    Set Rs_TotVar = Nothing
    Set Grafica = Nothing
    Set gChartData = Nothing
    Set gWorkBook = Nothing
    Set gWorkSheet = Nothing
    Set Rs_Acreed = Nothing
    Set Rs = Nothing
    Set Qdf = Nothing
    'Diapositiva.Shapes("Btn_").Delete
    Set Diapositiva = Nothing
End Sub

Sub Integra_Resumen(Pwr As PowerPoint.Presentation)
    Dim Diapositiva As PowerPoint.Slide
    Dim NumDiapRetorno As Integer
    
    Dim Qdf As DAO.QueryDef
    Dim Rs As DAO.Recordset
    Dim RsAux As DAO.Recordset
    Set Qdf = Db.QueryDefs("Cons_2026_ListaAcreedores")
    Qdf.Parameters("Id_Adecuado") = Id_Adecuado
    
    Dim sql As String
    sql = "SELECT Max(Consecutivo) As MáxDeConsecutivo From Tbl_Tmp_Acreed_Seg;"
    
    Set RsAux = Db.OpenRecordset(sql)
    Dim Consecutivo As Integer
    'Format(Valor, "@") = ""
    If Format(RsAux(0).Value, "@") = "" Then
        Consecutivo = 1
    Else
        Consecutivo = RsAux(0).Value + 1
    End If
    
    Set Rs = Qdf.OpenRecordset
    
    Dim Fechita As Date
    Fechita = Now
    
    Db.Execute ("DELETE * From Tbl_Temporal_ConsentraResumen;")
    
    Dim QdfAux As DAO.QueryDef
    Set QdfAux = Db.QueryDefs("Cons_2026_LlenaTablaConceptosResumen")
    QdfAux.Parameters("Id_Adecuado") = Id_Adecuado
    QdfAux.Execute
    
    
    Set QdfAux = Db.QueryDefs("Cons_2026_ParaLlenarTablaPowerPoint")
    'QdfAux.Parameters("Id_Adecuado") = Id_Adecuado
    Dim Primer As Boolean
    Dim TieneProyeccion As Boolean
    
    While Not Rs.EOF
        sql = "INSERT INTO Tbl_Tmp_Acreed_Seg (Id_Adecuado,Fecha,Consecutivo,Nombre_Reporte,Acreedor) SELECT " & Id_Adecuado & ", '" & Fechita & "', " & Consecutivo & ", '" & Pwr.Name & "', '" & Rs("CIA_ABREV") & "';"
        Db.Execute (sql)
        QdfAux.Parameters("Acreedor") = Rs("CIA_ABREV")
        Set RsAux = QdfAux.OpenRecordset
        
        'pruebas
        
        'fin de las pruebas
        
        sql = "UPDATE Tbl_Tmp_Acreed_Seg SET "
        Dim Sumita As Double
        Dim Largo As Long
        Largo = Len(sql)
        Sumita = 0
        TieneProyeccion = False
        While Not RsAux.EOF
            
            Select Case RsAux("Concepto_Resumen")
                Case "Facturado"
                    sql = sql & "Facturado = " & RsAux("Total") & ", "
                    
                Case "Recepcion"
                    sql = sql & "Recepcion = " & RsAux("Total") & ", "
                    
                Case "COPADE"
                    sql = sql & "Copade = " & RsAux("Total") & ", "
                    
                Case "FA"
                    sql = sql & "FA = " & RsAux("Total") & ", "
                    
                Case "SIT"
                    sql = sql & "SIT = " & RsAux("Total") & ", "
                    
                Case "Pedido S/R"
                    Select Case RsAux("Formalizado")
                        Case "Adecuado I"
                            sql = sql & "AdecI = " & RsAux("Total") & ", "
                        Case "Adecuado II"
                            sql = sql & "AdecII = " & RsAux("Total") & ", "
                        Case "Adecuado III"
                            sql = sql & "AdecIII = " & RsAux("Total") & ", "
                        Case "Adecuado IV"
                            sql = sql & "AdecIV = " & RsAux("Total") & ", "
                    End Select
                    
                Case "Terminados"
                    sql = sql & "Terminado = " & RsAux("Total") & ", "
                    
                Case "Proyección"
                    sql = sql & "Proyeccion = " & RsAux("Total") & ", "
                    Sumita = RsAux("Total")
                    TieneProyeccion = True
                Case "SolPed"
                    sql = sql & "SolP = " & RsAux("Total") & ", "
                    
                Case "RAF"
                    sql = sql & "RAF = " & RsAux("Total") & ", "
                
                Case "PTE"
                    If TieneProyeccion Then
                        Dim BuscaPr As Integer
                        BuscaPr = InStr(1, sql, "Proyeccion")
                        sql = Mid(sql, 1, BuscaPr - 1)
                        sql = sql & "Proyeccion = " & RsAux("Total") + Sumita & ", "
                    Else
                        sql = sql & "Proyeccion = " & RsAux("Total") & ", "
                    End If
                    
                                   
                Case "Convenio"
                    sql = sql & "Convenio = " & RsAux("Total") & ", "
                    
                Case "Ped.Sin Recep. Form. Adec I"
                    sql = sql & "AdecI = " & RsAux("Total") & ", "
                    
                Case "Ped. Sin Recep. Form. Adec II"
                    sql = sql & "AdecII = " & RsAux("Total") & ", "
                    
                Case "Ped. Sin Recep. Form. Adec III"
                    sql = sql & "AdecIII = " & RsAux("Total") & ", "
                    
                Case "Ped. Sin Recep. Form. Adec IV"
                    sql = sql & "AdecIV = " & RsAux("Total") & ", "
                    
                Case "FAO"
                    sql = sql & "FAO = " & RsAux("Total") & ", "
                    
                Case "Ped. Sin Recep. Form. 2024"
                    
                    
                Case "Ped. Sin Recep. Form. 2025"
                    Select Case RsAux("Formalizado")
                        Case "Adecuado I"
                            sql = sql & "AdecI = " & RsAux("Total") & ", "
                        Case "Adecuado II"
                            sql = sql & "AdecII = " & RsAux("Total") & ", "
                        Case "Adecuado III"
                            sql = sql & "AdecIII = " & RsAux("Total") & ", "
                        Case "Adecuado IV"
                            sql = sql & "AdecIV = " & RsAux("Total") & ", "
                    End Select
                    

            End Select
            RsAux.MoveNext
        Wend
        
        sql = Mid(sql, 1, Len(sql) - 2)
        sql = sql & " where Id_Adecuado=" & Id_Adecuado & " and Consecutivo = " & Consecutivo & " and Acreedor = '" & Rs("CIA_ABREV") & "';"
        Db.Execute (sql)
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    Set RsAux = Nothing
    
    Set QdfAux = Nothing
    Set Qdf = Nothing
    
    Set Qdf = Db.QueryDefs("Cons_2026_Tbl_ResumenAcreedor")
    Qdf.Parameters("Id_Adecuado") = Id_Adecuado
    Qdf.Parameters("Consecutivo") = Consecutivo
    
    Set Rs = Qdf.OpenRecordset
   
    
    Dim ColumnasTbl(17) As Proceso.Tipo_Tabla
    ColumnasTbl(1).Alineacion = 4
    ColumnasTbl(1).Encabezado = "Acreedores"
    ColumnasTbl(1).Size_Fuente = 9
    ColumnasTbl(1).Ancho = CentimetrosAPuntos(4.1)
    ColumnasTbl(1).Fuente = FuenteOficial
    ColumnasTbl(2).SeSuma = False
    
    
    ColumnasTbl(2).Alineacion = 2
    ColumnasTbl(2).Encabezado = "Facturado"
    ColumnasTbl(2).Size_Fuente = 9
    ColumnasTbl(2).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(2).Fuente = FuenteOficial
    ColumnasTbl(2).SeSuma = True
    ColumnasTbl(2).Tipo = "Numero"
    ColumnasTbl(2).Total = 0
    
    ColumnasTbl(3).Alineacion = 2
    ColumnasTbl(3).Encabezado = "COPADE"
    ColumnasTbl(3).Size_Fuente = 9
    ColumnasTbl(3).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(3).Fuente = FuenteOficial
    ColumnasTbl(3).SeSuma = True
    ColumnasTbl(3).Tipo = "Numero"
    ColumnasTbl(3).Total = 0
    
    ColumnasTbl(4).Alineacion = 2
    ColumnasTbl(4).Encabezado = "Recepción"
    ColumnasTbl(4).Size_Fuente = 9
    ColumnasTbl(4).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(4).Fuente = FuenteOficial
    ColumnasTbl(4).SeSuma = True
    ColumnasTbl(4).Tipo = "Numero"
    ColumnasTbl(4).Total = 0
    
    ColumnasTbl(5).Alineacion = 2
    ColumnasTbl(5).Encabezado = "FA"
    ColumnasTbl(5).Size_Fuente = 9
    ColumnasTbl(5).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(5).Fuente = FuenteOficial
    ColumnasTbl(5).SeSuma = True
    ColumnasTbl(5).Tipo = "Numero"
    ColumnasTbl(5).Total = 0
    
    ColumnasTbl(6).Alineacion = 2
    ColumnasTbl(6).Encabezado = "Terminado"
    ColumnasTbl(6).Size_Fuente = 9
    ColumnasTbl(6).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(6).Fuente = FuenteOficial
    ColumnasTbl(6).SeSuma = True
    ColumnasTbl(6).Tipo = "Numero"
    ColumnasTbl(6).Total = 0
    
    ColumnasTbl(7).Alineacion = 2
    ColumnasTbl(7).Encabezado = "Adec I"
    ColumnasTbl(7).Size_Fuente = 9
    ColumnasTbl(7).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(7).Fuente = FuenteOficial
    ColumnasTbl(7).SeSuma = True
    ColumnasTbl(7).Tipo = "Numero"
    ColumnasTbl(7).Total = 0
    
    ColumnasTbl(8).Alineacion = 2
    ColumnasTbl(8).Encabezado = "Adec II"
    ColumnasTbl(8).Size_Fuente = 9
    ColumnasTbl(8).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(8).Fuente = FuenteOficial
    ColumnasTbl(8).SeSuma = True
    ColumnasTbl(8).Tipo = "Numero"
    ColumnasTbl(8).Total = 0
    
    ColumnasTbl(9).Alineacion = 2
    ColumnasTbl(9).Encabezado = "Adec III"
    ColumnasTbl(9).Size_Fuente = 9
    ColumnasTbl(9).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(9).Fuente = FuenteOficial
    ColumnasTbl(9).SeSuma = True
    ColumnasTbl(9).Tipo = "Numero"
    ColumnasTbl(9).Total = 0
    
    ColumnasTbl(10).Alineacion = 2
    ColumnasTbl(10).Encabezado = "Adec IV"
    ColumnasTbl(10).Size_Fuente = 9
    ColumnasTbl(10).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(10).Fuente = FuenteOficial
    ColumnasTbl(10).SeSuma = True
    ColumnasTbl(10).Tipo = "Numero"
    ColumnasTbl(10).Total = 0
    
    ColumnasTbl(11).Alineacion = 2
    ColumnasTbl(11).Encabezado = "SolP"
    ColumnasTbl(11).Size_Fuente = 9
    ColumnasTbl(11).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(11).Fuente = FuenteOficial
    ColumnasTbl(11).SeSuma = True
    ColumnasTbl(11).Tipo = "Numero"
    ColumnasTbl(11).Total = 0
    
    ColumnasTbl(12).Alineacion = 2
    ColumnasTbl(12).Encabezado = "FAO"
    ColumnasTbl(12).Size_Fuente = 9
    ColumnasTbl(12).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(12).Fuente = FuenteOficial
    ColumnasTbl(12).SeSuma = True
    ColumnasTbl(12).Tipo = "Numero"
    ColumnasTbl(12).Total = 0
    
    ColumnasTbl(13).Alineacion = 2
    ColumnasTbl(13).Encabezado = "SIT"
    ColumnasTbl(13).Size_Fuente = 9
    ColumnasTbl(13).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(13).Fuente = FuenteOficial
    ColumnasTbl(13).SeSuma = True
    ColumnasTbl(13).Tipo = "Numero"
    ColumnasTbl(13).Total = 0
    
    ColumnasTbl(14).Alineacion = 2
    ColumnasTbl(14).Encabezado = "RAF"
    ColumnasTbl(14).Size_Fuente = 9
    ColumnasTbl(14).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(14).Fuente = FuenteOficial
    ColumnasTbl(14).SeSuma = True
    ColumnasTbl(14).Tipo = "Numero"
    ColumnasTbl(14).Total = 0
    
    ColumnasTbl(15).Alineacion = 2
    ColumnasTbl(15).Encabezado = "Proyeccion"
    ColumnasTbl(15).Size_Fuente = 9
    ColumnasTbl(15).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(15).Fuente = FuenteOficial
    ColumnasTbl(15).SeSuma = True
    ColumnasTbl(15).Tipo = "Numero"
    ColumnasTbl(15).Total = 0
    
    ColumnasTbl(16).Alineacion = 2
    ColumnasTbl(16).Encabezado = "Convenio"
    ColumnasTbl(16).Size_Fuente = 9
    ColumnasTbl(16).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(16).Fuente = FuenteOficial
    ColumnasTbl(16).SeSuma = True
    ColumnasTbl(16).Tipo = "Numero"
    ColumnasTbl(16).Total = 0
    
    ColumnasTbl(17).Alineacion = 3
    ColumnasTbl(17).Encabezado = "Total"
    ColumnasTbl(17).Size_Fuente = 9
    ColumnasTbl(17).Ancho = CentimetrosAPuntos(1.8)
    ColumnasTbl(17).Fuente = FuenteOficial
    ColumnasTbl(17).SeSuma = True
    ColumnasTbl(17).Tipo = "Numero"
    ColumnasTbl(17).Total = 0
    
    
    'Set Tabla = Proceso.Crea_Tabla(Diapositiva, CentimetrosAPuntos(1.15), CentimetrosAPuntos(3.53), ColumnasTbl)
    Dim Num_DiapRetorno As Integer
    Dim TotalDiap As Integer
    Num_DiapRetorno = Crea_LaTablaJorgito(Pwr, ColumnasTbl, Rs, Variables.No_DiapResumenAcreedores, 5, "Desglose Por Acreedor", False, True, False, 1, "Tbl_Resumen", CentimetrosAPuntos(0.55), CentimetrosAPuntos(2.48), CentimetrosAPuntos(0.52), Variables.Color_Encabezado, Variables.Color_Rng1, Variables.Color_Rng2, Variables.Color_Grupo, TotalDiap, Variables.Fecha_Consumo)
    
    
    Dim pptTable As PowerPoint.Table
    Dim k As Integer
    Dim Tot1, Tot2, Tot3 As Double
    Dim j As Integer
    
    Tot1 = Val(Replace(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Cell(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text, ",", "")) + _
        Val(Replace(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Cell(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Rows.Count, 3).Shape.TextFrame.TextRange.Text, ",", "")) + _
        Val(Replace(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Cell(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Rows.Count, 4).Shape.TextFrame.TextRange.Text, ",", "")) + _
        Val(Replace(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Cell(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Rows.Count, 5).Shape.TextFrame.TextRange.Text, ",", ""))
    
    Tot2 = Val(Replace(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Cell(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Rows.Count, 6).Shape.TextFrame.TextRange.Text, ",", "")) + _
        Val(Replace(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Cell(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Rows.Count, 7).Shape.TextFrame.TextRange.Text, ",", "")) + _
        Val(Replace(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Cell(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Rows.Count, 8).Shape.TextFrame.TextRange.Text, ",", "")) + _
        Val(Replace(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Cell(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Rows.Count, 9).Shape.TextFrame.TextRange.Text, ",", "")) + _
        Val(Replace(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Cell(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Rows.Count, 10).Shape.TextFrame.TextRange.Text, ",", ""))
        
    
    Tot3 = Val(Replace(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Cell(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Rows.Count, 11).Shape.TextFrame.TextRange.Text, ",", "")) + _
        Val(Replace(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Cell(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Rows.Count, 12).Shape.TextFrame.TextRange.Text, ",", "")) + _
        Val(Replace(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Cell(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Rows.Count, 13).Shape.TextFrame.TextRange.Text, ",", "")) + _
        Val(Replace(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Cell(Pwr.Slides(Num_DiapRetorno + TotalDiap - 1).Shapes("Tbl_Resumen").Table.Rows.Count, 14).Shape.TextFrame.TextRange.Text, ",", ""))
    
    For k = Num_DiapRetorno To Num_DiapRetorno + TotalDiap - 1
        Set Diapositiva = Pwr.Slides(k)
        Set pptTable = Diapositiva.Shapes("Tbl_Resumen").Table
        
        pptTable.Rows.Add (1)
        pptTable.Cell(1, 2).Merge mergeto:=pptTable.Cell(1, 5)
        pptTable.Cell(1, 6).Merge mergeto:=pptTable.Cell(1, 10)
        pptTable.Cell(1, 11).Merge mergeto:=pptTable.Cell(1, 14)
        pptTable.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Activ. Concluidas " & Format(Tot1, "#,##0")
        pptTable.Cell(1, 6).Shape.TextFrame.TextRange.Text = "Activ. Contratadas " & Format(Tot2, "#,##0")
        pptTable.Cell(1, 11).Shape.TextFrame.TextRange.Text = "En Contratación " & Format(Tot3, "#,##0")
        
        For j = 1 To pptTable.Columns.Count
            
            pptTable.Cell(2, j).Shape.TextFrame.TextRange.Paragraphs.Font.Color.RGB = RGB(255, 255, 255)
            
        Next
        
    Next
   
    
    
    Dim ControlProceso(14) As ProcesoConsulta
    ControlProceso(1).Proc_Almacenado = "Cons_2026_Activ_Facturadas"
    ControlProceso(1).Consulta = "Facturado"
    ControlProceso(1).Estructura = "Estructura_Consumo"
    ControlProceso(1).EncabezadoVariable = "Fecha Real Facturacion"
    ControlProceso(1).EncabezdoDiap = "Actividades Facturadas de "
    
    ControlProceso(2).Proc_Almacenado = "Cons_2026_Activ_COPADE"
    ControlProceso(2).Consulta = "COPADE"
    ControlProceso(2).Estructura = "Estructura_Consumo"
    ControlProceso(2).EncabezadoVariable = "Semana de Elaboracion COPADE"
    ControlProceso(2).EncabezdoDiap = "Actividades con COPADE de "
    
    ControlProceso(3).Proc_Almacenado = "Cons_2026_Activ_Recepcion"
    ControlProceso(3).Consulta = "Recepcion"
    ControlProceso(3).Estructura = "Estructura_Consumo"
    ControlProceso(3).EncabezadoVariable = "Fecha Recepción"
    ControlProceso(3).EncabezdoDiap = "Actividadaes que se encuentran recepcionadas de "
    
    ControlProceso(4).Proc_Almacenado = "Cons_2026_Activ_FA"
    ControlProceso(4).Consulta = "FA"
    ControlProceso(4).Estructura = "Estructura_Consumo"
    ControlProceso(4).EncabezadoVariable = "Fecha Solicitud FA/Pago"
    ControlProceso(4).EncabezdoDiap = "Actividades solicitadas en Formato Azul Corto para pago de "
    
    ControlProceso(5).Proc_Almacenado = "Cons_2026_Activ_Ejecucion"
    ControlProceso(5).Consulta = "Terminados"
    ControlProceso(5).Estructura = "Ejecucion"
    ControlProceso(5).EncabezadoVariable = ""
    ControlProceso(5).EncabezdoDiap = "Actividades terminadas, en documentación para solicitar el FA (pago) de "
    
    ControlProceso(6).Proc_Almacenado = "Cons_2026_Activ_EjecucionAdec"
    ControlProceso(6).Consulta = "Adecuado I"
    ControlProceso(6).Estructura = "Ejecucion"
    ControlProceso(6).EncabezadoVariable = ""
    ControlProceso(6).EncabezdoDiap = "Actividades en ejecución que fueron formalizadas en el Adec I de "
    
    ControlProceso(7).Proc_Almacenado = "Cons_2026_Activ_EjecucionAdec"
    ControlProceso(7).Consulta = "Adecuado II"
    ControlProceso(7).Estructura = "Ejecucion"
    ControlProceso(7).EncabezadoVariable = ""
    ControlProceso(7).EncabezdoDiap = "Actividades en ejecución que fueron formalizadas en el Adec II de "
    
    ControlProceso(8).Proc_Almacenado = "Cons_2026_Activ_EjecucionAdec"
    ControlProceso(8).Consulta = "Adecuado III"
    ControlProceso(8).Estructura = "Ejecucion"
    ControlProceso(8).EncabezadoVariable = ""
    ControlProceso(8).EncabezdoDiap = "Actividades en ejecución que fueron formalizadas en el Adec III de "
    
    ControlProceso(9).Proc_Almacenado = "Cons_2026_Activ_EjecucionAdec"
    ControlProceso(9).Consulta = "Adecuado IV"
    ControlProceso(9).Estructura = "Ejecucion"
    ControlProceso(9).EncabezadoVariable = ""
    ControlProceso(9).EncabezdoDiap = "Actividades en ejecución que fueron formalizadas en el Adec IV de "
    
    ControlProceso(10).Proc_Almacenado = "Cons_2026_Activ_EjecucionSolP"
    ControlProceso(10).Consulta = "SolPed"
    ControlProceso(10).Estructura = "Ejecucion"
    ControlProceso(10).EncabezadoVariable = "SolP"
    ControlProceso(10).EncabezdoDiap = "Actividades en Solicitud de Pedido, en proceso de contratación de "
    
    ControlProceso(11).Proc_Almacenado = "Cons_2026_Activ_Contratacion"
    ControlProceso(11).Consulta = "FAO"
    ControlProceso(11).Estructura = "Contratacion"
    ControlProceso(11).EncabezadoVariable = ""
    ControlProceso(11).EncabezdoDiap = "Actividades solicitadas en Formato Azul Largo (Recursos para Operar) de "
    
    ControlProceso(12).Proc_Almacenado = "Cons_2026_Activ_Contratacion"
    ControlProceso(12).Consulta = "SIT"
    ControlProceso(12).Estructura = "Contratacion"
    ControlProceso(12).EncabezadoVariable = ""
    ControlProceso(12).EncabezdoDiap = "Actividades que su proceso de contratación están en RAF de "
    
    ControlProceso(13).Proc_Almacenado = "Cons_2026_Activ_Contratacion"
    ControlProceso(13).Consulta = "RAF"
    ControlProceso(13).Estructura = "Contratacion"
    ControlProceso(13).EncabezadoVariable = ""
    ControlProceso(13).EncabezdoDiap = "Actividades que su proceso de contratación están en SIT de "
    
    ControlProceso(14).Proc_Almacenado = "Cons_2026_Activ_Proyeccion"
    ControlProceso(14).Consulta = "Proyeccion"
    ControlProceso(14).Estructura = "Proyeccion"
    ControlProceso(14).EncabezadoVariable = ""
    ControlProceso(14).EncabezdoDiap = "Actividades autorizadas en el adecuado pero no han iniciado su contratacion de "
    
    
    
    Dim Columnas As Integer
    Dim Renglones As Integer
    
    Dim Acreedor As String
    Dim NombreColumna As String
    
    
     Dim Estructura_Consumo(6) As Proceso.Tipo_Tabla
     Dim Estructura_Ejecucion(6) As Proceso.Tipo_Tabla
     Dim Estructura_Terminada(5) As Proceso.Tipo_Tabla
     Dim Estructura_Contratacion(7) As Proceso.Tipo_Tabla
     Dim Estructura_Proyeccion(4) As Proceso.Tipo_Tabla
     
    Estructura_Consumo(1).Alineacion = 4
    Estructura_Consumo(1).Encabezado = "Responsable"
    Estructura_Consumo(1).Size_Fuente = 9
    Estructura_Consumo(1).Ancho = CentimetrosAPuntos(4.09)
    Estructura_Consumo(1).Fuente = FuenteOficial
    Estructura_Consumo(1).SeSuma = False
    
    Estructura_Consumo(2).Alineacion = 4
    Estructura_Consumo(2).Encabezado = "Descripcion de la Actividad"
    Estructura_Consumo(2).Size_Fuente = 9
    Estructura_Consumo(2).Ancho = CentimetrosAPuntos(19.4)
    Estructura_Consumo(2).Fuente = FuenteOficial
    Estructura_Consumo(2).SeSuma = False
    
    Estructura_Consumo(3).Alineacion = 2
    Estructura_Consumo(3).Encabezado = "No. Pedido"
    Estructura_Consumo(3).Size_Fuente = 9
    Estructura_Consumo(3).Ancho = CentimetrosAPuntos(2.4)
    Estructura_Consumo(3).Fuente = FuenteOficial
    Estructura_Consumo(3).SeSuma = False
    
    Estructura_Consumo(4).Alineacion = 2
    Estructura_Consumo(4).Encabezado = "Fecha Progr. Entrega"
    Estructura_Consumo(4).Size_Fuente = 9
    Estructura_Consumo(4).Ancho = CentimetrosAPuntos(2.4)
    Estructura_Consumo(4).Fuente = FuenteOficial
    Estructura_Consumo(4).SeSuma = False
    
    Estructura_Consumo(5).Alineacion = 2
    Estructura_Consumo(5).Encabezado = "Fecha Real Entrega"
    Estructura_Consumo(5).Size_Fuente = 9
    Estructura_Consumo(5).Ancho = CentimetrosAPuntos(2.4)
    Estructura_Consumo(5).Fuente = FuenteOficial
    Estructura_Consumo(5).SeSuma = False
    
    Estructura_Consumo(6).Alineacion = 3
    Estructura_Consumo(6).Encabezado = "Monto"
    Estructura_Consumo(6).Size_Fuente = 9
    Estructura_Consumo(6).Ancho = CentimetrosAPuntos(2.4)
    Estructura_Consumo(6).Fuente = FuenteOficial
    Estructura_Consumo(6).SeSuma = True
    Estructura_Consumo(6).Tipo = "Numero"
    Estructura_Consumo(6).Total = 0
    
    Estructura_Terminada(1).Alineacion = 4
    Estructura_Terminada(1).Encabezado = "Responsable"
    Estructura_Terminada(1).Size_Fuente = 9
    Estructura_Terminada(1).Ancho = CentimetrosAPuntos(4.09)
    Estructura_Terminada(1).Fuente = FuenteOficial
    Estructura_Terminada(1).SeSuma = False
    
    Estructura_Terminada(2).Alineacion = 4
    Estructura_Terminada(2).Encabezado = "Descripcion de la Actividad"
    Estructura_Terminada(2).Size_Fuente = 9
    Estructura_Terminada(2).Ancho = CentimetrosAPuntos(19.4)
    Estructura_Terminada(2).Fuente = FuenteOficial
    Estructura_Terminada(2).SeSuma = False
    
    Estructura_Terminada(3).Alineacion = 2
    Estructura_Terminada(3).Encabezado = "No. Pedido"
    Estructura_Terminada(3).Size_Fuente = 9
    Estructura_Terminada(3).Ancho = CentimetrosAPuntos(2.4)
    Estructura_Terminada(3).Fuente = FuenteOficial
    Estructura_Terminada(3).SeSuma = False
    
    Estructura_Terminada(4).Alineacion = 2
    Estructura_Terminada(4).Encabezado = "Fecha Progr. Entrega"
    Estructura_Terminada(4).Size_Fuente = 9
    Estructura_Terminada(4).Ancho = CentimetrosAPuntos(2.4)
    Estructura_Terminada(4).Fuente = FuenteOficial
    Estructura_Terminada(4).SeSuma = False
            
    Estructura_Terminada(5).Alineacion = 3
    Estructura_Terminada(5).Encabezado = "Monto"
    Estructura_Terminada(5).Size_Fuente = 9
    Estructura_Terminada(5).Ancho = CentimetrosAPuntos(1.8)
    Estructura_Terminada(5).Fuente = FuenteOficial
    Estructura_Terminada(5).SeSuma = True
    Estructura_Terminada(5).Tipo = "Numero"
    Estructura_Terminada(5).Total = 0

    
       
    Estructura_Ejecucion(1).Alineacion = 4
    Estructura_Ejecucion(1).Encabezado = "Responsable"
    Estructura_Ejecucion(1).Size_Fuente = 9
    Estructura_Ejecucion(1).Ancho = CentimetrosAPuntos(4.09)
    Estructura_Ejecucion(1).Fuente = FuenteOficial
    Estructura_Ejecucion(1).SeSuma = False
    
    Estructura_Ejecucion(2).Alineacion = 4
    Estructura_Ejecucion(2).Encabezado = "Descripcion de la Actividad"
    Estructura_Ejecucion(2).Size_Fuente = 9
    Estructura_Ejecucion(2).Ancho = CentimetrosAPuntos(19.4)
    Estructura_Ejecucion(2).Fuente = FuenteOficial
    Estructura_Ejecucion(2).SeSuma = False
    
    Estructura_Ejecucion(3).Alineacion = 2
    Estructura_Ejecucion(3).Encabezado = "No. Pedido"
    Estructura_Ejecucion(3).Size_Fuente = 9
    Estructura_Ejecucion(3).Ancho = CentimetrosAPuntos(2.4)
    Estructura_Ejecucion(3).Fuente = FuenteOficial
    Estructura_Ejecucion(3).SeSuma = False
    
    Estructura_Ejecucion(4).Alineacion = 2
    Estructura_Ejecucion(4).Encabezado = "Fecha Progr. Entrega"
    Estructura_Ejecucion(4).Size_Fuente = 9
    Estructura_Ejecucion(4).Ancho = CentimetrosAPuntos(2.4)
    Estructura_Ejecucion(4).Fuente = FuenteOficial
    Estructura_Ejecucion(4).SeSuma = False
    
    Estructura_Ejecucion(5).Alineacion = 2
    Estructura_Ejecucion(5).Encabezado = "Estatus"
    Estructura_Ejecucion(5).Size_Fuente = 9
    Estructura_Ejecucion(5).Ancho = CentimetrosAPuntos(3.1)
    Estructura_Ejecucion(5).Fuente = FuenteOficial
    Estructura_Ejecucion(5).SeSuma = False
        
    Estructura_Ejecucion(6).Alineacion = 3
    Estructura_Ejecucion(6).Encabezado = "Monto"
    Estructura_Ejecucion(6).Size_Fuente = 9
    Estructura_Ejecucion(6).Ancho = CentimetrosAPuntos(1.8)
    Estructura_Ejecucion(6).Fuente = FuenteOficial
    Estructura_Ejecucion(6).SeSuma = True
    Estructura_Ejecucion(6).Tipo = "Numero"
    Estructura_Ejecucion(6).Total = 0
       
       
    Estructura_Contratacion(1).Alineacion = 4
    Estructura_Contratacion(1).Encabezado = "Responsable"
    Estructura_Contratacion(1).Size_Fuente = 9
    Estructura_Contratacion(1).Ancho = CentimetrosAPuntos(4.09)
    Estructura_Contratacion(1).Fuente = FuenteOficial
    Estructura_Contratacion(1).SeSuma = False
    
    Estructura_Contratacion(2).Alineacion = 4
    Estructura_Contratacion(2).Encabezado = "Descripcion de la Actividad"
    Estructura_Contratacion(2).Size_Fuente = 9
    Estructura_Contratacion(2).Ancho = CentimetrosAPuntos(13.8)
    Estructura_Contratacion(2).Fuente = FuenteOficial
    Estructura_Contratacion(2).SeSuma = False
    
    Estructura_Contratacion(3).Alineacion = 2
    Estructura_Contratacion(3).Encabezado = "Plazo de Ejecución"
    Estructura_Contratacion(3).Size_Fuente = 9
    Estructura_Contratacion(3).Ancho = CentimetrosAPuntos(2.4)
    Estructura_Contratacion(3).Fuente = FuenteOficial
    Estructura_Contratacion(3).SeSuma = False
    
    Estructura_Contratacion(4).Alineacion = 2
    Estructura_Contratacion(4).Encabezado = "Folio SIT"
    Estructura_Contratacion(4).Size_Fuente = 9
    Estructura_Contratacion(4).Ancho = CentimetrosAPuntos(4)
    Estructura_Contratacion(4).Fuente = FuenteOficial
    Estructura_Contratacion(4).SeSuma = False
    
    Estructura_Contratacion(5).Alineacion = 2
    Estructura_Contratacion(5).Encabezado = "Folio RAF"
    Estructura_Contratacion(5).Size_Fuente = 9
    Estructura_Contratacion(5).Ancho = CentimetrosAPuntos(4)
    Estructura_Contratacion(5).Fuente = FuenteOficial
    Estructura_Contratacion(5).SeSuma = False
    
    Estructura_Contratacion(6).Alineacion = 2
    Estructura_Contratacion(6).Encabezado = "Fecha Solicitud FAO"
    Estructura_Contratacion(6).Size_Fuente = 9
    Estructura_Contratacion(6).Ancho = CentimetrosAPuntos(2.4)
    Estructura_Contratacion(6).Fuente = FuenteOficial
    Estructura_Contratacion(6).SeSuma = False
        
    
    Estructura_Contratacion(7).Alineacion = 3
    Estructura_Contratacion(7).Encabezado = "Monto"
    Estructura_Contratacion(7).Size_Fuente = 9
    Estructura_Contratacion(7).Ancho = CentimetrosAPuntos(2.4)
    Estructura_Contratacion(7).Fuente = FuenteOficial
    Estructura_Contratacion(7).SeSuma = True
    Estructura_Contratacion(7).Tipo = "Numero"
    Estructura_Contratacion(7).Total = 0
    
    Estructura_Proyeccion(1).Alineacion = 4
    Estructura_Proyeccion(1).Encabezado = "Responsable"
    Estructura_Proyeccion(1).Size_Fuente = 9
    Estructura_Proyeccion(1).Ancho = CentimetrosAPuntos(4.09)
    Estructura_Proyeccion(1).Fuente = FuenteOficial
    Estructura_Proyeccion(1).SeSuma = False
    
    Estructura_Proyeccion(2).Alineacion = 4
    Estructura_Proyeccion(2).Encabezado = "Descripcion de la Actividad"
    Estructura_Proyeccion(2).Size_Fuente = 9
    Estructura_Proyeccion(2).Ancho = CentimetrosAPuntos(19.4)
    Estructura_Proyeccion(2).Fuente = FuenteOficial
    Estructura_Proyeccion(2).SeSuma = False
    
    Estructura_Proyeccion(3).Alineacion = 4
    Estructura_Proyeccion(3).Encabezado = "Estatus"
    Estructura_Proyeccion(3).Size_Fuente = 9
    Estructura_Proyeccion(3).Ancho = CentimetrosAPuntos(2.4)
    Estructura_Proyeccion(3).Fuente = FuenteOficial
    Estructura_Proyeccion(3).SeSuma = False
    
    Estructura_Proyeccion(4).Alineacion = 3
    Estructura_Proyeccion(4).Encabezado = "Monto"
    Estructura_Proyeccion(4).Size_Fuente = 9
    Estructura_Proyeccion(4).Ancho = CentimetrosAPuntos(2.4)
    Estructura_Proyeccion(4).Fuente = FuenteOficial
    Estructura_Proyeccion(4).SeSuma = True
    Estructura_Proyeccion(4).Tipo = "Numero"
    Estructura_Proyeccion(4).Total = 0
    
    Dim Celda As PowerPoint.Cell
    
    For k = Num_DiapRetorno To Num_DiapRetorno + TotalDiap - 1
        Set Diapositiva = Pwr.Slides(k)
        Set pptTable = Diapositiva.Shapes("Tbl_Resumen").Table
        
        For Renglones = 3 To pptTable.Rows.Count - 1
            For Columnas = 2 To pptTable.Columns.Count - 2
                If (pptTable.Cell(Renglones, Columnas).Shape.TextFrame.TextRange.Text <> "") And (pptTable.Cell(Renglones, 1).Shape.TextFrame.TextRange.Text <> "SubTotal") Then
                    Acreedor = pptTable.Cell(Renglones, 1).Shape.TextFrame.TextRange.Text
                    NombreColumna = pptTable.Cell(2, Columnas).Shape.TextFrame.TextRange.Text
                    
                    If Columnas - 1 >= 1 And Columnas - 1 <= 4 Then
                        Set Qdf = Db.QueryDefs(ControlProceso(Columnas - 1).Proc_Almacenado)
                        Qdf.Parameters("Id_Adecuado") = Id_Adecuado
                        Qdf.Parameters("Acreedor") = Acreedor
                        Qdf.Parameters("Consulta") = ControlProceso(Columnas - 1).Consulta
                        Estructura_Consumo(5).Encabezado = ControlProceso(Columnas - 1).EncabezadoVariable
                        Set Rs = Qdf.OpenRecordset
                        
                        Num_DiapRetorno = Crea_LaTablaJorgito(Pwr, Estructura_Consumo, Rs, Variables.No_DiapResumenAcreedores, k, ControlProceso(Columnas - 1).EncabezdoDiap & Acreedor, False, True, False, 1, "Tbl_Resumen", CentimetrosAPuntos(0.55), CentimetrosAPuntos(3.21), CentimetrosAPuntos(0.52), Variables.Color_Encabezado, Variables.Color_Rng1, Variables.Color_Rng2, Variables.Color_Grupo, TotalDiap, Variables.Fecha_Consumo)
                        Call Pone_LigaEnCelda(pptTable.Cell(Renglones, Columnas), Num_DiapRetorno)
                        
                    End If
                    
                    If Columnas - 1 = 5 Then
                        Set Qdf = Db.QueryDefs(ControlProceso(Columnas - 1).Proc_Almacenado)
                        Qdf.Parameters("Id_Adecuado") = Id_Adecuado
                        Qdf.Parameters("Acreedor") = Acreedor
                        Qdf.Parameters("Consulta") = ControlProceso(Columnas - 1).Consulta
                        Set Rs = Qdf.OpenRecordset
                        
                        Num_DiapRetorno = Crea_LaTablaJorgito(Pwr, Estructura_Terminada, Rs, Variables.No_DiapResumenAcreedores, k, ControlProceso(Columnas - 1).EncabezdoDiap & Acreedor, False, True, False, 1, "Tbl_Resumen", CentimetrosAPuntos(0.55), CentimetrosAPuntos(3.21), CentimetrosAPuntos(0.52), Variables.Color_Encabezado, Variables.Color_Rng1, Variables.Color_Rng2, Variables.Color_Grupo, TotalDiap, Variables.Fecha_Consumo)
                        Call Pone_LigaEnCelda(pptTable.Cell(Renglones, Columnas), Num_DiapRetorno)
                    End If
                    
                    If Columnas - 1 >= 6 And Columnas - 1 <= 9 Then
                        Set Qdf = Db.QueryDefs(ControlProceso(Columnas - 1).Proc_Almacenado)
                        Qdf.Parameters("Id_Adecuado") = Id_Adecuado
                        Qdf.Parameters("Acreedor") = Acreedor
                        Qdf.Parameters("Adecuado") = ControlProceso(Columnas - 1).Consulta
                        Set Rs = Qdf.OpenRecordset
                        
                        Num_DiapRetorno = Crea_LaTablaJorgito(Pwr, Estructura_Ejecucion, Rs, Variables.No_DiapResumenAcreedores, k, ControlProceso(Columnas - 1).EncabezdoDiap & Acreedor, False, True, False, 1, "Tbl_Resumen", CentimetrosAPuntos(0.55), CentimetrosAPuntos(3.21), CentimetrosAPuntos(0.52), Variables.Color_Encabezado, Variables.Color_Rng1, Variables.Color_Rng2, Variables.Color_Grupo, TotalDiap, Variables.Fecha_Consumo)
                        Call Pone_LigaEnCelda(pptTable.Cell(Renglones, Columnas), Num_DiapRetorno)
                    End If
                    
                    If Columnas - 1 = 10 Then
                        Set Qdf = Db.QueryDefs(ControlProceso(Columnas - 1).Proc_Almacenado)
                        Qdf.Parameters("Id_Adecuado") = Id_Adecuado
                        Qdf.Parameters("Acreedor") = Acreedor
                        Qdf.Parameters("Consulta") = ControlProceso(Columnas - 1).Consulta
                        Estructura_Ejecucion(3).Encabezado = "No. SolP"
                        Set Rs = Qdf.OpenRecordset
                        
                        Num_DiapRetorno = Crea_LaTablaJorgito(Pwr, Estructura_Terminada, Rs, Variables.No_DiapResumenAcreedores, k, ControlProceso(Columnas - 1).EncabezdoDiap & Acreedor, False, True, False, 1, "Tbl_Resumen", CentimetrosAPuntos(0.55), CentimetrosAPuntos(3.21), CentimetrosAPuntos(0.52), Variables.Color_Encabezado, Variables.Color_Rng1, Variables.Color_Rng2, Variables.Color_Grupo, TotalDiap, Variables.Fecha_Consumo)
                        Call Pone_LigaEnCelda(pptTable.Cell(Renglones, Columnas), Num_DiapRetorno)
                    End If
                    
                    If Columnas - 1 >= 11 And Columnas - 1 <= 13 Then
                        Set Qdf = Db.QueryDefs(ControlProceso(Columnas - 1).Proc_Almacenado)
                        Qdf.Parameters("Id_Adecuado") = Id_Adecuado
                        Qdf.Parameters("Acreedor") = Acreedor
                        Qdf.Parameters("Consulta") = ControlProceso(Columnas - 1).Consulta
                        Set Rs = Qdf.OpenRecordset
                        
                        Num_DiapRetorno = Crea_LaTablaJorgito(Pwr, Estructura_Contratacion, Rs, Variables.No_DiapResumenAcreedores, k, ControlProceso(Columnas - 1).EncabezdoDiap & Acreedor, False, True, False, 1, "Tbl_Resumen", CentimetrosAPuntos(0.55), CentimetrosAPuntos(3.21), CentimetrosAPuntos(0.52), Variables.Color_Encabezado, Variables.Color_Rng1, Variables.Color_Rng2, Variables.Color_Grupo, TotalDiap, Variables.Fecha_Consumo)
                        Call Pone_LigaEnCelda(pptTable.Cell(Renglones, Columnas), Num_DiapRetorno)
                    End If
                    
                    If Columnas - 1 = 14 Then
                        Set Qdf = Db.QueryDefs(ControlProceso(Columnas - 1).Proc_Almacenado)
                        Qdf.Parameters("Id_Adecuado") = Id_Adecuado
                        Qdf.Parameters("Acreedor") = Acreedor
                        Set Rs = Qdf.OpenRecordset
                        
                        Num_DiapRetorno = Crea_LaTablaJorgito(Pwr, Estructura_Proyeccion, Rs, Variables.No_DiapResumenAcreedores, k, ControlProceso(Columnas - 1).EncabezdoDiap & Acreedor, False, True, False, 1, "Tbl_Resumen", CentimetrosAPuntos(0.55), CentimetrosAPuntos(3.21), CentimetrosAPuntos(0.52), Variables.Color_Encabezado, Variables.Color_Rng1, Variables.Color_Rng2, Variables.Color_Grupo, TotalDiap, Variables.Fecha_Consumo)
                        Call Pone_LigaEnCelda(pptTable.Cell(Renglones, Columnas), Num_DiapRetorno)
                    End If
                    
                    'Estructura_Contratacion
                    'pptTable.Cell(Renglones, Columnas).Shape.Fill.ForeColor.RGB = RGB(123, 12, 45)
                End If
            Next
        Next
    Next
    
    
    Set RsAux = Nothing
    Set Rs = Nothing
End Sub

Function Crea_LaTablaJorgito(Pp As PowerPoint.Presentation, ByRef Estructura() As Proceso.Tipo_Tabla, Datos As DAO.Recordset, NumDiapBase As Integer, QuienMeLlama As Integer, Encabezado As String, NumerarRenglones As Boolean, CerosEnBlanco As Boolean, Agrupar As Boolean, AgruparPor As Integer, NombreTabla As String, PosX As Double, PosY As Double, AltoFila As Double, ColorEncabezado As String, ColorAlterno1 As String, ColorAlterno2 As String, ColorSubTotal As String, ByRef CuantasDiap As Integer, FechaReporte As Date) As Integer
    Dim SbTotales() As Proceso.Tipo_Tabla
    Dim Totales() As Proceso.Tipo_Tabla
    Dim TotGrupo() As Proceso.Tipo_Tabla
    Dim RenglonesResumen() As Integer
    Dim ResaltaRenglon As Integer
    
    ReDim Preserve RenglonesResumen(0)
    ReDim Preserve SbTotales(UBound(Estructura))
    ReDim Preserve Totales(UBound(Estructura))
    ReDim Preserve TotGrupo(UBound(Estructura))
    
    Dim Ig As Integer
    For Ig = 1 To UBound(SbTotales)
        Totales(Ig).Total = 0
        SbTotales(Ig).Total = 0
        TotGrupo(Ig).Total = 0
    Next
    
    Dim Diapositivita As PowerPoint.Slide
    Dim Diap_Anterior As PowerPoint.Slide
    Pp.Slides(NumDiapBase).Duplicate
    Set Diapositivita = Pp.Slides(NumDiapBase)
    
    Diapositivita.SlideShowTransition.AdvanceOnClick = msoFalse
    Diapositivita.SlideShowTransition.Hidden = msoTrue
    Diapositivita.Shapes("Btn_Anterior").Delete
    Diapositivita.Shapes("Txt_Fechita").TextFrame.TextRange.Text = "Fecha de Consumos " & Format(FechaReporte, "dd mmm yyyy")

    Diapositivita.MoveTo (Pp.Slides.Count - 1)
    Crea_LaTablaJorgito = Diapositivita.SlideNumber
        
    Diapositivita.Shapes.Title.TextFrame.TextRange.Text = Encabezado
       
    If QuienMeLlama <> 0 Then
        Dim Link As PowerPoint.ActionSetting
        Set Link = Diapositivita.Shapes("Btn_Inicio").ActionSettings(ppMouseClick)
        Link.Action = ppActionHyperlink
        Link.Hyperlink.SubAddress = QuienMeLlama
        
    Else
        Diapositivita.Shapes("Btn_Inicio").Delete
    End If
        
    Dim Tabla As PowerPoint.Shape
    
    Dim Tot_T As Double
    Tot_T = 0
    
    Dim STot_T As Double
    STot_T = 0
    
    Dim CtasDiap As Integer
    CtasDiap = 1
    
    Set Tabla = Proceso.Crea_Tabla(Diapositivita, PosX, PosY, Estructura)
    Tabla.Name = NombreTabla
    
    Dim Campo As DAO.Field
    Dim NumCampo As Integer
    Dim i As Integer
    Dim TotRenglnes As Integer
    Dim ContaRn As Integer
    TotRenglnes = Datos.RecordCount
    ContaRn = 0
    
    Dim IniciaEn As Integer
    IniciaEn = 0
    If NumerarRenglones Then IniciaEn = 1
    
    Dim AgruparPorValor As String
    AgruparPorValor = ""
    Dim EsElPrimero As Boolean
    EsElPrimero = True
    
    Dim RespaldoGrupo As Integer
    
    Dim GeneraRenglonGrupo As Boolean
    GeneraRenglonGrupo = False
    
    Dim CuantosLlevas As Integer
    CuantosLlevas = 0
    
    Dim ij As Integer
    Dim ColorRng As String
    ColorRng = ""
    
    Dim Cuentagrupos As Integer
    Cuentagrupos = 0
    
   ' Tabla.Fill.ForeColor.RGB = RGB(Colorcito(ColorAlterno2, 1), Colorcito(ColorAlterno2, 2), Colorcito(ColorAlterno2, 3))
    While Not Datos.EOF
        Tabla.Table.Rows.Add
        Tabla.Table.Rows(Tabla.Table.Rows.Count).Height = AltoFila
        NumCampo = IniciaEn
        ContaRn = ContaRn + 1
        ColorRng = ColorAlterno2
        If ContaRn Mod 2 = 0 Then ColorRng = ColorAlterno1
        
        For Each Campo In Datos.Fields
            NumCampo = NumCampo + 1
            
            Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.Fill.ForeColor.RGB = RGB(Colorcito(ColorRng, 1), Colorcito(ColorRng, 2), Colorcito(ColorRng, 3))
            
            If NumerarRenglones Then Tabla.Table.Cell(Tabla.Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = ContaRn & " / " & TotRenglnes
            If Estructura(NumCampo).SeSuma Then
                'Estructura(NumCampo).Total = Estructura(NumCampo).Total + Campo
                Totales(NumCampo).Total = Totales(NumCampo).Total + Val(Format(Campo, "@"))
                SbTotales(NumCampo).Total = SbTotales(NumCampo).Total + Val(Format(Campo, "@"))
                TotGrupo(NumCampo).Total = TotGrupo(NumCampo).Total + Val(Format(Campo, "@"))
'                Tot_T = Tot_T + Val(Str(Campo))
'                STot_T = STot_T + Val(Str(Campo))
                If CerosEnBlanco And Format(Campo, "#,##0.00") <> "0.00" Then
                    Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Text = Format(Campo, "#,##0.00")
                End If
                
                If Not CerosEnBlanco Then
                    Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Text = Format(Campo, "#,##0.00")
                End If
                
                If CerosEnBlanco And Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Text = "0.00" Then
                    Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Text = ""
                End If
            Else
                 
                If NumCampo = AgruparPor And Agrupar Then
                    If EsElPrimero Then
                        AgruparPorValor = Format(Campo, "@")
                        EsElPrimero = False
                    End If
                    If AgruparPorValor = Format(Campo, "@") And Not EsElPrimero Then
                        CuantosLlevas = CuantosLlevas + 1
'                        MsgBox "Hay que agrupar"
                        ' se debe llevar un conteo de renglones
                    Else
 
                        If CuantosLlevas > 1 Then
                            Cuentagrupos = Cuentagrupos + 1
                            ReDim Preserve RenglonesResumen(Cuentagrupos)
                            RespaldoGrupo = AgruparPor
                            If AgruparPor = 1 Then RespaldoGrupo = RespaldoGrupo + 1
                            
                            RenglonesResumen(Cuentagrupos) = Tabla.Table.Rows.Count
                            
                            Tabla.Table.Rows.Add (Tabla.Table.Rows.Count)
                            
                            Tabla.Table.Rows(Tabla.Table.Rows.Count).Height = AltoFila
                            Tabla.Table.Cell(Tabla.Table.Rows.Count - 1, RespaldoGrupo - 1).Shape.TextFrame.TextRange.Text = "Subtotal"
                            Tabla.Table.Cell(Tabla.Table.Rows.Count - 1, RespaldoGrupo).Merge mergeto:=Tabla.Table.Cell(Tabla.Table.Rows.Count - 1, RespaldoGrupo + 1)
                            Tabla.Table.Cell(Tabla.Table.Rows.Count - 1, RespaldoGrupo).Shape.TextFrame.TextRange.Text = ReemplazarPor(AgruparPorValor)
                            
                            For ij = 1 To Tabla.Table.Columns.Count
                                Tabla.Table.Cell(Tabla.Table.Rows.Count - 1, ij).Shape.TextFrame.TextRange.Font.Bold = msoTrue
                                Tabla.Table.Cell(Tabla.Table.Rows.Count - 1, ij).Shape.Fill.ForeColor.RGB = RGB(189, 215, 238)
                                
                                Tabla.Table.Cell(Tabla.Table.Rows.Count - 1, ij).Borders(ppBorderBottom).Style = msoLineSingle
                                Tabla.Table.Cell(Tabla.Table.Rows.Count - 1, ij).Borders(ppBorderBottom).Weight = 0.5
                                Tabla.Table.Cell(Tabla.Table.Rows.Count - 1, ij).Borders(ppBorderTop).Style = msoLineSingle
                                Tabla.Table.Cell(Tabla.Table.Rows.Count - 1, ij).Borders(ppBorderTop).Weight = 1
                                
                                If Estructura(ij).SeSuma Then
                                    
                                    Tabla.Table.Cell(Tabla.Table.Rows.Count - 1, ij).Shape.TextFrame.TextRange.Text = Format(TotGrupo(ij).Total, "#,##0.00")
                                    TotGrupo(ij).Total = 0
                                End If
                            Next
                        End If
                        'GeneraRenglonGrupo = True
                        AgruparPorValor = Format(Campo, "@")
                        CuantosLlevas = 1
                        
                         For ij = 1 To Tabla.Table.Columns.Count
                                If Estructura(ij).SeSuma Then TotGrupo(ij).Total = 0
                            Next
                        
                        'aqui se debe contolar la cantidad de renglones
                        'mandar a restablecer el control de la suma del grupo y generar el nuevo renglon
                    End If
                End If
                
                 Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Text = Format(Campo, "@")
            End If
        Next
        
        If Tabla.Height >= 330 Then
            CtasDiap = CtasDiap + 1
            
            Tabla.Table.Rows.Add
            Tabla.Table.LastRow = True
            Tabla.Table.Rows(Tabla.Table.Rows.Count).Height = AltoFila
            For NumCampo = 1 To Tabla.Table.Columns.Count
                Tabla.Table.Cell(1, NumCampo).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
                Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.Fill.ForeColor.RGB = RGB(Colorcito(ColorEncabezado, 1), Colorcito(ColorEncabezado, 2), Colorcito(ColorEncabezado, 3))
                Tabla.Table.Cell(1, NumCampo).Shape.Fill.ForeColor.RGB = RGB(Colorcito(ColorEncabezado, 1), Colorcito(ColorEncabezado, 2), Colorcito(ColorEncabezado, 3))
                If Estructura(NumCampo).SeSuma Then
                    Tabla.Table.Cell(Tabla.Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = "SubTotal"
                    If CerosEnBlanco And SbTotales(NumCampo).Total <> 0 Then
                        Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Text = Format(SbTotales(NumCampo).Total, "#,##0.00")
                     End If
                     
                     If Not CerosEnBlanco Then
                        Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Text = Format(SbTotales(NumCampo).Total, "#,##0.00")
                     End If
                     
                     SbTotales(NumCampo).Total = 0
                     Cuentagrupos = 0
                End If
            Next
            
            If Agrupar And UBound(RenglonesResumen) > 0 Then
                For ResaltaRenglon = 1 To UBound(RenglonesResumen)
                    For ij = 1 To Tabla.Table.Columns.Count
                        Tabla.Table.Cell(RenglonesResumen(ResaltaRenglon), ij).Shape.TextFrame.TextRange.Font.Bold = msoTrue
                        Tabla.Table.Cell(RenglonesResumen(ResaltaRenglon), ij).Shape.Fill.ForeColor.RGB = RGB(Colorcito(ColorSubTotal, 1), Colorcito(ColorSubTotal, 2), Colorcito(ColorSubTotal, 3))
                        
                        Tabla.Table.Cell(RenglonesResumen(ResaltaRenglon), ij).Borders(ppBorderBottom).Style = msoLineSingle
                        Tabla.Table.Cell(RenglonesResumen(ResaltaRenglon), ij).Borders(ppBorderBottom).Weight = 0.5
                        Tabla.Table.Cell(RenglonesResumen(ResaltaRenglon), ij).Borders(ppBorderBottom).ForeColor.RGB = RGB(0, 0, 0)
                        Tabla.Table.Cell(RenglonesResumen(ResaltaRenglon), ij).Borders(ppBorderTop).Style = msoLineSingle
                        Tabla.Table.Cell(RenglonesResumen(ResaltaRenglon), ij).Borders(ppBorderTop).Weight = 1
                        Tabla.Table.Cell(RenglonesResumen(ResaltaRenglon), ij).Borders(ppBorderTop).ForeColor.RGB = RGB(0, 0, 0)
                    Next
                Next
                
                ReDim RenglonesResumen(0)
            End If
            
            'Tabla.Table.ApplyStyle ("{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}")
            
            'STot_T = 0
            Set Diap_Anterior = Diapositivita
            
            Pp.Slides(NumDiapBase).Duplicate
            
            Set Diapositivita = Pp.Slides(NumDiapBase)
            Diapositivita.MoveTo (Pp.Slides.Count - 1)
            
            Diapositivita.SlideShowTransition.AdvanceOnClick = msoFalse
            Diapositivita.SlideShowTransition.Hidden = msoTrue
            
            Diap_Anterior.Shapes("Btn_Siguiente").ActionSettings(ppMouseClick).Action = ppActionHyperlink
            Diap_Anterior.Shapes("Btn_Siguiente").ActionSettings(ppMouseClick).Hyperlink.SubAddress = Diapositivita.SlideNumber
            
            Diapositivita.Shapes("Btn_Anterior").ActionSettings(ppMouseClick).Action = ppActionHyperlink
            Diapositivita.Shapes("Btn_Anterior").ActionSettings(ppMouseClick).Hyperlink.SubAddress = Diap_Anterior.SlideNumber
                            
            Diapositivita.Shapes.Title.TextFrame.TextRange.Text = Encabezado
            Diapositivita.Shapes("Txt_Fechita").TextFrame.TextRange.Text = "Fecha de Consumos " & Format(FechaReporte, "dd mmm yyyy")
            
            If QuienMeLlama <> 0 Then
                Set Link = Diapositivita.Shapes("Btn_Inicio").ActionSettings(ppMouseClick)
                Link.Action = ppActionHyperlink
                Link.Hyperlink.SubAddress = QuienMeLlama
            Else
                Diapositivita.Shapes("Btn_Inicio").Delete
            End If
            
            Set Tabla = Proceso.Crea_Tabla(Diapositivita, PosX, PosY, Estructura)
            Tabla.Name = NombreTabla
        End If
        
        Datos.MoveNext
    Wend
    
    If CuantosLlevas > 1 Then
        RespaldoGrupo = AgruparPor
        If AgruparPor = 1 Then RespaldoGrupo = RespaldoGrupo + 1
        
        Cuentagrupos = Cuentagrupos + 1
        ReDim Preserve RenglonesResumen(Cuentagrupos)
        RenglonesResumen(Cuentagrupos) = Tabla.Table.Rows.Count
        
        Tabla.Table.Rows.Add '(Tabla.Table.Rows.Count)
        Tabla.Table.Rows(Tabla.Table.Rows.Count).Height = AltoFila
        Tabla.Table.Cell(Tabla.Table.Rows.Count, RespaldoGrupo - 1).Shape.TextFrame.TextRange.Text = "Subtotal"
        Tabla.Table.Cell(Tabla.Table.Rows.Count, RespaldoGrupo).Merge mergeto:=Tabla.Table.Cell(Tabla.Table.Rows.Count, RespaldoGrupo + 1)
        Tabla.Table.Cell(Tabla.Table.Rows.Count, RespaldoGrupo).Shape.TextFrame.TextRange.Text = ReemplazarPor(AgruparPorValor)
        For ij = 1 To Tabla.Table.Columns.Count
            Tabla.Table.Cell(Tabla.Table.Rows.Count, ij).Shape.TextFrame.TextRange.Font.Bold = msoTrue
            Tabla.Table.Cell(Tabla.Table.Rows.Count, ij).Shape.Fill.ForeColor.RGB = RGB(189, 215, 238)
            
           If Estructura(ij).SeSuma Then
                
                Tabla.Table.Cell(Tabla.Table.Rows.Count, ij).Shape.TextFrame.TextRange.Text = Format(TotGrupo(ij).Total, "#,##0.00")
                TotGrupo(ij).Total = 0
            End If
            
        If CerosEnBlanco And Tabla.Table.Cell(Tabla.Table.Rows.Count, ij).Shape.TextFrame.TextRange.Text = "0.00" Then
            Tabla.Table.Cell(Tabla.Table.Rows.Count, ij).Shape.TextFrame.TextRange.Text = ""
        End If
            
        Next
    End If
    
    ' Tabla.Table.ApplyStyle ("{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}")
    
    If CtasDiap > 1 Then
        Tabla.Table.Rows.Add
         Tabla.Table.Rows(Tabla.Table.Rows.Count).Height = AltoFila
        For NumCampo = 1 To Tabla.Table.Columns.Count
            Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.Fill.ForeColor.RGB = RGB(Colorcito(ColorSubTotal, 1), Colorcito(ColorSubTotal, 2), Colorcito(ColorSubTotal, 3))
            Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
            'Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Font.Bold = msoTrue
            If Estructura(NumCampo).SeSuma Then
                Tabla.Table.Cell(Tabla.Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = "SubTotal"
                
                Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Text = Format(SbTotales(NumCampo).Total, "#,##0.00")
            End If
            
            If CerosEnBlanco And Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Text = "0.00" Then
                Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Text = ""
            End If
        Next
    End If
    
    Tabla.Table.Rows.Add
    Tabla.Table.LastRow = True
     Tabla.Table.Rows(Tabla.Table.Rows.Count).Height = AltoFila
    
    For NumCampo = 1 To Tabla.Table.Columns.Count
        Tabla.Table.Cell(1, NumCampo).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
        Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.Fill.ForeColor.RGB = RGB(Colorcito(ColorEncabezado, 1), Colorcito(ColorEncabezado, 2), Colorcito(ColorEncabezado, 3))
        Tabla.Table.Cell(1, NumCampo).Shape.Fill.ForeColor.RGB = RGB(Colorcito(ColorEncabezado, 1), Colorcito(ColorEncabezado, 2), Colorcito(ColorEncabezado, 3))
        If Estructura(NumCampo).SeSuma Then
            Tabla.Table.Cell(Tabla.Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = "Total"
            Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Text = Format(Totales(NumCampo).Total, "#,##0.00")
        End If
        
        If CerosEnBlanco And Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Text = "0.00" Then
            Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCampo).Shape.TextFrame.TextRange.Text = ""
        End If
        
    Next
    
    'Tabla.Table.ApplyStyle ("{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}")
    CuantasDiap = CtasDiap
    Diapositivita.Shapes("Btn_Siguiente").Delete
    Set Diapositivita = Nothing
    Set Diap_Anterior = Nothing
End Function


Function Colorcito(Colores As String, NumeroColor As Integer) As Integer
    Colorcito = 0
    Dim Aux As String
    Aux = ""
    Select Case NumeroColor
        Case 1
            Aux = Mid(Colores, 1, InStr(1, Colores, ",", vbTextCompare) - 1)
            Colorcito = Val(Aux)
        Case 2
            Aux = Mid(Colores, InStr(1, Colores, ",", vbTextCompare) + 1, Len(Colores) - InStr(1, Colores, ",", vbTextCompare))
            Aux = Mid(Aux, 1, InStr(1, Aux, ",", vbTextCompare) - 1)
            Colorcito = Val(Aux)
        Case 3
            Aux = Mid(Colores, InStr(1, Colores, ",", vbTextCompare) + 1, Len(Colores) - InStr(1, Colores, ",", vbTextCompare))
            Aux = Mid(Aux, InStr(1, Aux, ",", vbTextCompare) + 1, Len(Aux) - InStr(1, Aux, ",", vbTextCompare))
            'Aux = Mid(Aux, 1, InStr(1, Aux, ",", vbTextCompare) - 1)
            Colorcito = Val(Aux)
    End Select
End Function


Sub Integra_VariacionPrograma_Real(Pwr As PowerPoint.Presentation)
    Dim Diapositiva As PowerPoint.Slide
    Dim Num_DiapRetorno As Integer
    
    Dim Qdf As DAO.QueryDef
    Dim Rs As DAO.Recordset
    Dim RsAux As DAO.Recordset
 
    Dim sql As String
    
    sql = "SELECT Max(Tbl_Programado_vs_Facturado.Id_Control) As MáxDeId_Control From Tbl_Programado_vs_Facturado;"
    
    Set RsAux = Db.OpenRecordset(sql)
    Dim Consecutivo As Integer
    'Format(Valor, "@") = ""
    If Format(RsAux(0).Value, "@") = "" Then
        Consecutivo = 1
    Else
        Consecutivo = RsAux(0).Value + 1
    End If
    
    Set Qdf = Db.QueryDefs("Cons_2026_InsertaMontoProgramado")
    Qdf.Parameters("Id_Consecutivo") = Consecutivo
    Qdf.Parameters("No_Adecuado") = Id_Adecuado
    Qdf.Parameters("Fecha_Actual") = Now()
    Qdf.Execute
    
    Set Qdf = Db.QueryDefs("Actualiza_2026_Mesesito")
    Qdf.Parameters("Id_Control") = Consecutivo
    Qdf.Execute
        
    Set Qdf = Db.QueryDefs("Cons_2026_ActualizarVariacion")
    Qdf.Parameters("Id_Adecuado") = Id_Adecuado
     
    Set Rs = Qdf.OpenRecordset
    
    While Not Rs.EOF
        
        sql = "UPDATE Tbl_Programado_vs_Facturado SET " & Proceso.RetornaMes3(Rs("Real")) & "=" & Rs("Monto") & ", Mes_Programado='" & Proceso.RetornaMes3(Rs("Programa")) & "' where Id_Control=" & Consecutivo & " and Id_Adecuado=" & Id_Adecuado & " and No_Mes=" & Rs("Programa") & ";"
        Db.Execute (sql)
        Rs.MoveNext
    Wend
    
     Dim Estructura(16) As Proceso.Tipo_Tabla
     
    Estructura(1).Alineacion = 4
    Estructura(1).Encabezado = "Mes Programado"
    Estructura(1).Size_Fuente = 9
    Estructura(1).Ancho = CentimetrosAPuntos(2.4)
    Estructura(1).Fuente = FuenteOficial
    Estructura(1).SeSuma = False
    
    Estructura(2).Alineacion = 2
    Estructura(2).Encabezado = "Monto Programado"
    Estructura(2).Size_Fuente = 9
    Estructura(2).Ancho = CentimetrosAPuntos(2.2)
    Estructura(2).Fuente = FuenteOficial
    Estructura(2).SeSuma = True
    Estructura(2).Tipo = "Numero"
    Estructura(2).Total = 0
    
    Estructura(3).Alineacion = 2
    Estructura(3).Encabezado = "Variación"
    Estructura(3).Size_Fuente = 9
    Estructura(3).Ancho = CentimetrosAPuntos(2.2)
    Estructura(3).Fuente = FuenteOficial
    Estructura(3).SeSuma = True
    Estructura(3).Tipo = "Numero"
    Estructura(3).Total = 0
    
    Estructura(4).Alineacion = 2
    Estructura(4).Encabezado = "Real Facturado"
    Estructura(4).Size_Fuente = 9
    Estructura(4).Ancho = CentimetrosAPuntos(2.2)
    Estructura(4).Fuente = FuenteOficial
    Estructura(4).SeSuma = True
    Estructura(4).Tipo = "Numero"
    Estructura(4).Total = 0
    
    Estructura(5).Alineacion = 2
    Estructura(5).Encabezado = "Ene"
    Estructura(5).Size_Fuente = 9
    Estructura(5).Ancho = CentimetrosAPuntos(1.6)
    Estructura(5).Fuente = FuenteOficial
    Estructura(5).SeSuma = True
    Estructura(5).Tipo = "Numero"
    Estructura(5).Total = 0

    Estructura(6).Alineacion = 2
    Estructura(6).Encabezado = "Feb"
    Estructura(6).Size_Fuente = 9
    Estructura(6).Ancho = CentimetrosAPuntos(1.6)
    Estructura(6).Fuente = FuenteOficial
    Estructura(6).SeSuma = True
    Estructura(6).Tipo = "Numero"
    Estructura(6).Total = 0
    
    Estructura(7).Alineacion = 2
    Estructura(7).Encabezado = "Mar"
    Estructura(7).Size_Fuente = 9
    Estructura(7).Ancho = CentimetrosAPuntos(1.6)
    Estructura(7).Fuente = FuenteOficial
    Estructura(7).SeSuma = True
    Estructura(7).Tipo = "Numero"
    Estructura(7).Total = 0
    
    Estructura(8).Alineacion = 2
    Estructura(8).Encabezado = "Abr"
    Estructura(8).Size_Fuente = 9
    Estructura(8).Ancho = CentimetrosAPuntos(1.6)
    Estructura(8).Fuente = FuenteOficial
    Estructura(8).SeSuma = True
    Estructura(8).Tipo = "Numero"
    Estructura(8).Total = 0
    
    Estructura(9).Alineacion = 2
    Estructura(9).Encabezado = "May"
    Estructura(9).Size_Fuente = 9
    Estructura(9).Ancho = CentimetrosAPuntos(1.6)
    Estructura(9).Fuente = FuenteOficial
    Estructura(9).SeSuma = True
    Estructura(9).Tipo = "Numero"
    Estructura(9).Total = 0
    
    Estructura(10).Alineacion = 2
    Estructura(10).Encabezado = "Jun"
    Estructura(10).Size_Fuente = 9
    Estructura(10).Ancho = CentimetrosAPuntos(1.6)
    Estructura(10).Fuente = FuenteOficial
    Estructura(10).SeSuma = True
    Estructura(10).Tipo = "Numero"
    Estructura(10).Total = 0
    
    Estructura(11).Alineacion = 2
    Estructura(11).Encabezado = "Jul"
    Estructura(11).Size_Fuente = 9
    Estructura(11).Ancho = CentimetrosAPuntos(1.6)
    Estructura(11).Fuente = FuenteOficial
    Estructura(11).SeSuma = True
    Estructura(11).Tipo = "Numero"
    Estructura(11).Total = 0
    
    Estructura(12).Alineacion = 2
    Estructura(12).Encabezado = "Ago"
    Estructura(12).Size_Fuente = 9
    Estructura(12).Ancho = CentimetrosAPuntos(1.6)
    Estructura(12).Fuente = FuenteOficial
    Estructura(12).SeSuma = True
    Estructura(12).Tipo = "Numero"
    Estructura(12).Total = 0
    
    Estructura(13).Alineacion = 2
    Estructura(13).Encabezado = "Sep"
    Estructura(13).Size_Fuente = 9
    Estructura(13).Ancho = CentimetrosAPuntos(1.6)
    Estructura(13).Fuente = FuenteOficial
    Estructura(13).SeSuma = True
    Estructura(13).Tipo = "Numero"
    Estructura(13).Total = 0
    
    Estructura(14).Alineacion = 2
    Estructura(14).Encabezado = "Oct"
    Estructura(14).Size_Fuente = 9
    Estructura(14).Ancho = CentimetrosAPuntos(1.6)
    Estructura(14).Fuente = FuenteOficial
    Estructura(14).SeSuma = True
    Estructura(14).Tipo = "Numero"
    Estructura(14).Total = 0
    
    Estructura(15).Alineacion = 2
    Estructura(15).Encabezado = "Nov"
    Estructura(15).Size_Fuente = 9
    Estructura(15).Ancho = CentimetrosAPuntos(1.6)
    Estructura(15).Fuente = FuenteOficial
    Estructura(15).SeSuma = True
    Estructura(15).Tipo = "Numero"
    Estructura(15).Total = 0
    
    Estructura(16).Alineacion = 2
    Estructura(16).Encabezado = "Dic"
    Estructura(16).Size_Fuente = 9
    Estructura(16).Ancho = CentimetrosAPuntos(1.6)
    Estructura(16).Fuente = FuenteOficial
    Estructura(16).SeSuma = True
    Estructura(16).Tipo = "Numero"
    Estructura(16).Total = 0


    
    '
    Set Qdf = Nothing
    Set Qdf = Db.QueryDefs("Cons_2026_Programado_Variacion")
    Qdf.Parameters("No_Control") = Consecutivo
    
        
     Set Rs = Qdf.OpenRecordset
     Dim TotalDiap As Integer
    Num_DiapRetorno = Crea_LaTablaJorgito(Pwr, Estructura, Rs, Variables.No_DiapResumenAcreedores, 5, "Calendario de la facturación total a la fecha de los consumos", False, True, False, 1, "Tbl_Resumen", CentimetrosAPuntos(2.89), CentimetrosAPuntos(3.59), CentimetrosAPuntos(0.52), Variables.Color_Encabezado, Variables.Color_Rng1, Variables.Color_Rng2, Variables.Color_Grupo, TotalDiap, Variables.Fecha_Consumo)
    
    
    
    
    Dim RespaldoNumDiap As Integer
    RespaldoNumDiap = Num_DiapRetorno
    
'    Set Qdf = Db.QueryDefs("Cons_2026_Soportes_Variacion")
'    Qdf.Parameters("Id_Adecuado") = Id_Adecuado
'
'     Dim Estruc_Sop(8) As Proceso.Tipo_Tabla
'
'
'    Estruc_Sop(1).Alineacion = 4
'    Estruc_Sop(1).Encabezado = "Mes"
'    Estruc_Sop(1).Size_Fuente = 9
'    Estruc_Sop(1).Ancho = CentimetrosAPuntos(2.1)
'    Estruc_Sop(1).Fuente = FuenteOficial
'    Estruc_Sop(1).SeSuma = False
'    Estruc_Sop(1).Tipo = "Numero"
'    Estruc_Sop(1).Total = 0
'
'    Estruc_Sop(2).Alineacion = 4
'    Estruc_Sop(2).Encabezado = "Acreedor"
'    Estruc_Sop(2).Size_Fuente = 9
'    Estruc_Sop(2).Ancho = CentimetrosAPuntos(3.4)
'    Estruc_Sop(2).Fuente = FuenteOficial
'    Estruc_Sop(2).SeSuma = False
'
'    Estruc_Sop(3).Alineacion = 4
'    Estruc_Sop(3).Encabezado = "Responsable"
'    Estruc_Sop(3).Size_Fuente = 9
'    Estruc_Sop(3).Ancho = CentimetrosAPuntos(3.2)
'    Estruc_Sop(3).Fuente = FuenteOficial
'    Estruc_Sop(3).SeSuma = False
'    Estruc_Sop(3).Tipo = "Numero"
'    Estruc_Sop(3).Total = 0
'
'    Estruc_Sop(4).Alineacion = 2
'    Estruc_Sop(4).Encabezado = "No. Pedido"
'    Estruc_Sop(4).Size_Fuente = 9
'    Estruc_Sop(4).Ancho = CentimetrosAPuntos(2.1)
'    Estruc_Sop(4).Fuente = FuenteOficial
'    Estruc_Sop(4).SeSuma = False
'    Estruc_Sop(4).Tipo = "Numero"
'    Estruc_Sop(4).Total = 0
'
'    Estruc_Sop(5).Alineacion = 4
'    Estruc_Sop(5).Encabezado = "Descripcion de la Actividad"
'    Estruc_Sop(5).Size_Fuente = 9
'    Estruc_Sop(5).Ancho = CentimetrosAPuntos(15.8)
'    Estruc_Sop(5).Fuente = FuenteOficial
'    Estruc_Sop(5).SeSuma = False
'    Estruc_Sop(5).Tipo = "Numero"
'    Estruc_Sop(5).Total = 0
'
'    Estruc_Sop(6).Alineacion = 4
'    Estruc_Sop(6).Encabezado = "Estatus"
'    Estruc_Sop(6).Size_Fuente = 9
'    Estruc_Sop(6).Ancho = CentimetrosAPuntos(2.1)
'    Estruc_Sop(6).Fuente = FuenteOficial
'    Estruc_Sop(6).SeSuma = False
'    Estruc_Sop(6).Tipo = "Numero"
'    Estruc_Sop(6).Total = 0
'
'    Estruc_Sop(7).Alineacion = 2
'    Estruc_Sop(7).Encabezado = "Fecha de Pago"
'    Estruc_Sop(7).Size_Fuente = 9
'    Estruc_Sop(7).Ancho = CentimetrosAPuntos(2.1)
'    Estruc_Sop(7).Fuente = FuenteOficial
'    Estruc_Sop(7).SeSuma = False
'    Estruc_Sop(7).Tipo = "Numero"
'    Estruc_Sop(7).Total = 0
'
'    Estruc_Sop(8).Alineacion = 3
'    Estruc_Sop(8).Encabezado = "Monto"
'    Estruc_Sop(8).Size_Fuente = 9
'    Estruc_Sop(8).Ancho = CentimetrosAPuntos(1.9)
'    Estruc_Sop(8).Fuente = FuenteOficial
'    Estruc_Sop(8).SeSuma = True
'    Estruc_Sop(8).Tipo = "Numero"
'    Estruc_Sop(8).Total = 0
'
'    Set Rs = Qdf.OpenRecordset
''
'    Num_DiapRetorno = Crea_LaTablaJorgito(Pwr, Estruc_Sop, Rs, Variables.No_DiapResumenAcreedores, Num_DiapRetorno, "Soporte de la Variación", False, True, True, 1, "Tbl_Resumen", CentimetrosAPuntos(0.74), CentimetrosAPuntos(2.29), CentimetrosAPuntos(0.52), Variables.Color_Encabezado, Variables.Color_Rng1, Variables.Color_Rng2, "245,245,245", TotalDiap, Variables.Fecha_Consumo)
'    Call Pone_LigaEnCelda(Pwr.Slides(RespaldoNumDiap).Shapes("Tbl_Resumen").Table.Cell(Pwr.Slides(RespaldoNumDiap).Shapes("Tbl_Resumen").Table.Rows.Count, 3), Num_DiapRetorno)
'
'    Pwr.Slides(Variables.No_DiapGraficaVariacion).Shapes("Graf_Cumplimiento").Copy
'    Pwr.Slides(Variables.No_DiapGraficaVariacion).Shapes("Graf_Cumplimiento").Copy
''    DoEvents
'    Pwr.Slides(RespaldoNumDiap).Shapes.Paste
'    Pwr.Slides(RespaldoNumDiap).Shapes("Graf_Cumplimiento").Top = CentimetrosAPuntos(12.93)
'    Pwr.Slides(RespaldoNumDiap).Shapes("Graf_Cumplimiento").Left = CentimetrosAPuntos(2.89)
'
'    Pwr.Slides(Variables.No_DiapGraficaVariacion).Shapes("Txt_Evaluacion").Copy
'    Pwr.Slides(Variables.No_DiapGraficaVariacion).Shapes("Txt_Evaluacion").Copy
''    DoEvents
'    Pwr.Slides(RespaldoNumDiap).Shapes.Paste
'    Pwr.Slides(RespaldoNumDiap).Shapes("Txt_Evaluacion").Top = CentimetrosAPuntos(12.16)
'    Pwr.Slides(RespaldoNumDiap).Shapes("Txt_Evaluacion").Left = CentimetrosAPuntos(2.35)
     
    
    
    
    
    Pwr.Slides(RespaldoNumDiap).Shapes("Tbl_Resumen").Table.Rows.Add (1)
    Pwr.Slides(RespaldoNumDiap).Shapes("Tbl_Resumen").Table.Cell(1, 1).Merge mergeto:=Pwr.Slides(RespaldoNumDiap).Shapes("Tbl_Resumen").Table.Cell(1, 4)
    Pwr.Slides(RespaldoNumDiap).Shapes("Tbl_Resumen").Table.Cell(1, 5).Merge mergeto:=Pwr.Slides(RespaldoNumDiap).Shapes("Tbl_Resumen").Table.Cell(1, 16)
    
    Pwr.Slides(RespaldoNumDiap).Shapes("Tbl_Resumen").Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Programado vs Real"
    Pwr.Slides(RespaldoNumDiap).Shapes("Tbl_Resumen").Table.Cell(1, 5).Shape.TextFrame.TextRange.Text = "Mes Real de Facturacion"
      
    Dim Num_Col As Integer
    For Num_Col = 1 To Pwr.Slides(RespaldoNumDiap).Shapes("Tbl_Resumen").Table.Columns.Count
        Pwr.Slides(RespaldoNumDiap).Shapes("Tbl_Resumen").Table.Cell(2, Num_Col).Shape.TextFrame.TextRange.Paragraphs.Font.Color.RGB = RGB(255, 255, 255)
    Next
    
    
    
'    Set Qdf = Db.QueryDefs("Cons_2026_ProgrTot_vs_RealTot")
'    Qdf.Parameters("Id_Adecuado") = Id_Adecuado
'    Qdf.Parameters("Periodo_Term") = Variables.Periodo_Fin
'
'
'    Set Rs = Qdf.OpenRecordset
'
'
'    Set Grafica = Pwr.Slides(RespaldoNumDiap).Shapes("Graf_Cumplimiento").Chart
'    Set gChartData = Grafica.ChartData
'    Set gWorkBook = gChartData.Workbook
'
'    Dim RenglXls As Integer
'    RenglXls = 3
'    While Not Rs.EOF
'        RenglXls = RenglXls + 1
'        On Error Resume Next
'        gWorkBook.Worksheets(1).Range("G" & RenglXls).Value = Rs("Monto")
'        Rs.MoveNext
'    Wend
'
'    Grafica.Refresh
'    gWorkBook.Close
    
    
    
'    Set Qdf = Db.QueryDefs("Cons_2026_ResponsablesVariacion")
'    Qdf.Parameters("Id_Adecuado") = Id_Adecuado
'    Qdf.Parameters("Periodo_Fin") = Variables.Periodo_Fin
'
'
'    Set Rs = Qdf.OpenRecordset
'
'
'    Dim EncabezadosdeTabla(4) As Proceso.Tipo_Tabla
'
'    EncabezadosdeTabla(1).Alineacion = 4
'    EncabezadosdeTabla(1).Encabezado = "Responsable"
'    EncabezadosdeTabla(1).Size_Fuente = 9
'    EncabezadosdeTabla(1).Ancho = CentimetrosAPuntos(5.9)
'    EncabezadosdeTabla(1).Fuente = FuenteOficial
'    EncabezadosdeTabla(1).SeSuma = False
'    EncabezadosdeTabla(1).Tipo = "Cadena"
'    EncabezadosdeTabla(1).Total = 0
'
'    EncabezadosdeTabla(2).Alineacion = 4
'    EncabezadosdeTabla(2).Encabezado = "Acreedor"
'    EncabezadosdeTabla(2).Size_Fuente = 9
'    EncabezadosdeTabla(2).Ancho = CentimetrosAPuntos(3.1)
'    EncabezadosdeTabla(2).Fuente = FuenteOficial
'    EncabezadosdeTabla(2).SeSuma = False
'    EncabezadosdeTabla(2).Tipo = "Cadena"
'    EncabezadosdeTabla(2).Total = 0
'
'    EncabezadosdeTabla(3).Alineacion = 4
'    EncabezadosdeTabla(3).Encabezado = "Estatus"
'    EncabezadosdeTabla(3).Size_Fuente = 9
'    EncabezadosdeTabla(3).Ancho = CentimetrosAPuntos(2.7)
'    EncabezadosdeTabla(3).Fuente = FuenteOficial
'    EncabezadosdeTabla(3).SeSuma = False
'    EncabezadosdeTabla(3).Tipo = "Cadena"
'    EncabezadosdeTabla(3).Total = 0
'
'    EncabezadosdeTabla(4).Alineacion = 3
'    EncabezadosdeTabla(4).Encabezado = "Monto"
'    EncabezadosdeTabla(4).Size_Fuente = 9
'    EncabezadosdeTabla(4).Ancho = CentimetrosAPuntos(1.4)
'    EncabezadosdeTabla(4).Fuente = FuenteOficial
'    EncabezadosdeTabla(4).SeSuma = True
'    EncabezadosdeTabla(4).Tipo = "Numero"
'    EncabezadosdeTabla(4).Total = 0
'
'
'    Dim Tabla As PowerPoint.Shape
'    Set Tabla = Proceso.Crea_Tabla(Pwr.Slides(RespaldoNumDiap), CentimetrosAPuntos(17.55), CentimetrosAPuntos(11.99), EncabezadosdeTabla)
'
'    Dim ColorRng As String
'
''
''
'    Dim NumCol As Integer
'
'    While Not Rs.EOF
'        Tabla.Table.Rows.Add
'        ColorRng = Variables.Color_Rng2
'        If Tabla.Table.Rows.Count Mod 2 = 0 Then ColorRng = Variables.Color_Rng1
'        For NumCol = 1 To Tabla.Table.Columns.Count
'            Tabla.Table.Cell(Tabla.Table.Rows.Count, NumCol).Shape.Fill.ForeColor.RGB = RGB(Colorcito(ColorRng, 1), Colorcito(ColorRng, 2), Colorcito(ColorRng, 3))
'        Next
'        Tabla.Table.Cell(Tabla.Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = Rs("Responsable")
'        Tabla.Table.Cell(Tabla.Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Rs("CIA_ABREV")
'        Tabla.Table.Cell(Tabla.Table.Rows.Count, 3).Shape.TextFrame.TextRange.Text = Rs("Concepto_Corto")
'        Tabla.Table.Cell(Tabla.Table.Rows.Count, 4).Shape.TextFrame.TextRange.Text = Format(Rs("Monto"), "#,##0.00")
'        EncabezadosdeTabla(4).Total = EncabezadosdeTabla(4).Total + Rs("Monto")
'        Rs.MoveNext
'    Wend
'
'    Tabla.Table.Rows.Add
'    Tabla.Table.LastRow = True
'    Tabla.Table.Cell(Tabla.Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = "Total"
'    Tabla.Table.Cell(Tabla.Table.Rows.Count, 4).Shape.TextFrame.TextRange.Text = Format(EncabezadosdeTabla(4).Total, "#,##0.00")
'
'    Dim ij As Integer
'
'    For ij = 1 To Tabla.Table.Columns.Count
'        Tabla.Table.Cell(Tabla.Table.Rows.Count, ij).Shape.Fill.ForeColor.RGB = RGB(Colorcito(Variables.Color_Encabezado, 1), Colorcito(Variables.Color_Encabezado, 2), Colorcito(Variables.Color_Encabezado, 3))
'        Tabla.Table.Cell(1, ij).Shape.Fill.ForeColor.RGB = RGB(Colorcito(Variables.Color_Encabezado, 1), Colorcito(Variables.Color_Encabezado, 2), Colorcito(Variables.Color_Encabezado, 3))
'    Next
'
'    Tabla.Height = 0
    
'    Stp_2026_Averigua_MesVariacion

    If Variables.Soportes Then
        Dim StructuraX(7) As Proceso.Tipo_Tabla
        StructuraX(1).Alineacion = 4
        StructuraX(1).Encabezado = "Acreedor"
        StructuraX(1).Size_Fuente = 9
        StructuraX(1).Ancho = CentimetrosAPuntos(5.5)
        StructuraX(1).Fuente = FuenteOficial
        StructuraX(1).SeSuma = False
        StructuraX(1).Tipo = "Cadena"
        StructuraX(1).Total = 0
        
        StructuraX(2).Alineacion = 4
        StructuraX(2).Encabezado = "Supervisor Tecnico"
        StructuraX(2).Size_Fuente = 9
        StructuraX(2).Ancho = CentimetrosAPuntos(5.1)
        StructuraX(2).Fuente = FuenteOficial
        StructuraX(2).SeSuma = False
        StructuraX(2).Tipo = "Cadena"
        StructuraX(2).Total = 0
        
        StructuraX(3).Alineacion = 4
        StructuraX(3).Encabezado = "Actividad"
        StructuraX(3).Size_Fuente = 9
        StructuraX(3).Ancho = CentimetrosAPuntos(8.6)
        StructuraX(3).Fuente = FuenteOficial
        StructuraX(3).SeSuma = False
        StructuraX(3).Tipo = "Cadena"
        StructuraX(3).Total = 0
        
        StructuraX(4).Alineacion = 2
        StructuraX(4).Encabezado = "Fecha Termino"
        StructuraX(4).Size_Fuente = 9
        StructuraX(4).Ancho = CentimetrosAPuntos(3.1)
        StructuraX(4).Fuente = FuenteOficial
        StructuraX(4).SeSuma = False
        StructuraX(4).Tipo = "Cadena"
        StructuraX(4).Total = 0
        
        StructuraX(5).Alineacion = 2
        StructuraX(5).Encabezado = "Pedido"
        StructuraX(5).Size_Fuente = 9
        StructuraX(5).Ancho = CentimetrosAPuntos(2.9)
        StructuraX(5).Fuente = FuenteOficial
        StructuraX(5).SeSuma = False
        StructuraX(5).Tipo = "Cadena"
        StructuraX(5).Total = 0
        
        StructuraX(6).Alineacion = 4
        StructuraX(6).Encabezado = "Estatus"
        StructuraX(6).Size_Fuente = 9
        StructuraX(6).Ancho = CentimetrosAPuntos(3.1)
        StructuraX(6).Fuente = FuenteOficial
        StructuraX(6).SeSuma = False
        StructuraX(6).Tipo = "Cadena"
        StructuraX(6).Total = 0
        
        StructuraX(7).Alineacion = 3
        StructuraX(7).Encabezado = "Monto"
        StructuraX(7).Size_Fuente = 9
        StructuraX(7).Ancho = CentimetrosAPuntos(2.12)
        StructuraX(7).Fuente = FuenteOficial
        StructuraX(7).SeSuma = True
        StructuraX(7).Tipo = "Numerico"
        StructuraX(7).Total = 0
        
        Set Qdf = Db.QueryDefs("Stp_2026_Soporte_MesVariacion")
        Qdf.Parameters("Id_Adecuado") = Id_AdecuadoOficial
    
    
        Dim i As Integer
        
        For i = 3 To Variables.MesFinal + 2
             If Pwr.Slides(RespaldoNumDiap).Shapes("Tbl_Resumen").Table.Cell(i, 3).Shape.TextFrame.TextRange.Text <> "" Then
                    Qdf.Parameters("No_Mes") = i - 2
                    Set Rs = Qdf.OpenRecordset
                    Num_DiapRetorno = Crea_LaTablaJorgito(Pwr, StructuraX, Rs, Variables.No_DiapResumenAcreedores, RespaldoNumDiap, "Soporte de la Variación del mes de " & Proceso.RetornaMes3(i - 2), False, True, False, 1, "Tbl_Resumen", CentimetrosAPuntos(0.74), CentimetrosAPuntos(2.89), CentimetrosAPuntos(0.52), Variables.Color_Encabezado, Variables.Color_Rng1, Variables.Color_Rng2, "245,245,245", TotalDiap, Variables.Fecha_Consumo)
                    Call Pone_LigaEnCelda(Pwr.Slides(RespaldoNumDiap).Shapes("Tbl_Resumen").Table.Cell(i, 3), Num_DiapRetorno)
                  
                    
             End If
        Next
    End If
    Set Rs = Nothing
    Set RsAux = Nothing
    Set Qdf = Nothing

End Sub

Sub Pone_LigaEnCelda(Celdita As PowerPoint.Cell, Hacia As Integer)
    Dim Link As PowerPoint.ActionSetting
    
'    Set Liguita = Tbl.Table.Cell(ContadorRenglon, Rs.Fields(ColNum).Name + 1).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick)
'                    Liguita.Action = ppActionHyperlink
'                    Liguita.Hyperlink.SubAddress = RetornoNumDiap
    
    Set Link = Celdita.Shape.TextFrame.TextRange.ActionSettings(ppMouseClick)
    
    Link.Action = ppActionHyperlink
    Link.Hyperlink.SubAddress = Hacia
End Sub


Sub Integra_Iniciativa(Pwr As PowerPoint.Presentation)
    Dim Diapositiva As PowerPoint.Slide
    Dim Num_DiapRetorno As Integer
    Dim TotalDiap As Integer
    
    Dim Qdf As DAO.QueryDef
    Dim Rs As DAO.Recordset
    
    
    Set Qdf = Db.QueryDefs("Cons_2026_Iniciativas")
    Qdf.Parameters("Id_Adecuado") = Id_Adecuado
    Set Rs = Qdf.OpenRecordset
    
    Dim Estruc(4) As Proceso.Tipo_Tabla
     
    
    Estruc(1).Alineacion = 4
    Estruc(1).Encabezado = "Estatus"
    Estruc(1).Size_Fuente = 9
    Estruc(1).Ancho = CentimetrosAPuntos(2.1)
    Estruc(1).Fuente = FuenteOficial
    Estruc(1).SeSuma = False
    Estruc(1).Tipo = "Cadena"
    Estruc(1).Total = 0
    
    Estruc(2).Alineacion = 4
    Estruc(2).Encabezado = "Numero"
    Estruc(2).Size_Fuente = 9
    Estruc(2).Ancho = CentimetrosAPuntos(2.1)
    Estruc(2).Fuente = FuenteOficial
    Estruc(2).SeSuma = False
    Estruc(2).Tipo = "Cadena"
    Estruc(2).Total = 0
    
    Estruc(3).Alineacion = 4
    Estruc(3).Encabezado = "Descripción"
    Estruc(3).Size_Fuente = 9
    Estruc(3).Ancho = CentimetrosAPuntos(8)
    Estruc(3).Fuente = FuenteOficial
    Estruc(3).SeSuma = False
    Estruc(3).Tipo = "Cadena"
    Estruc(3).Total = 0
    
    Estruc(4).Alineacion = 2
    Estruc(4).Encabezado = "Monto"
    Estruc(4).Size_Fuente = 9
    Estruc(4).Ancho = CentimetrosAPuntos(2.1)
    Estruc(4).Fuente = FuenteOficial
    Estruc(4).SeSuma = True
    Estruc(4).Tipo = "Numero"
    Estruc(4).Total = 0
    
    Num_DiapRetorno = Crea_LaTablaJorgito(Pwr, Estruc, Rs, Variables.No_DiapResumenAcreedores, 5, "Iniciativas", False, True, False, 1, "Tbl_Iniciativas", CentimetrosAPuntos(3.98), CentimetrosAPuntos(3.29), CentimetrosAPuntos(0.52), Variables.Color_Encabezado, Variables.Color_Rng1, Variables.Color_Rng2, "245,245,245", TotalDiap, Variables.Fecha_Consumo)
    
    Set Qdf = Nothing
    Set Rs = Nothing
    '
End Sub

Sub Pedidos_Calendarizados(Pwr As PowerPoint.Presentation)
    Call PrepararInformacion
        
    Dim Grafica As PowerPoint.Chart
    Dim gChartData As PowerPoint.ChartData
    Dim gWorkBook As excel.Workbook
    Dim gWorkSheet As New excel.Worksheet
    Dim Diapositiva As PowerPoint.Slide
    Dim NumDiapRetorno As Integer
    
    Dim Qdf As DAO.QueryDef
    Dim Rs As DAO.Recordset
    
    Dim sql As String
    
    Dim TxtComent As PowerPoint.Shape
    Dim Flechita As PowerPoint.Shape
    Dim Ancho_Barrita As Double
    Dim Ancho_Coment As Double
    Dim Ki As Integer
   Dim i As Integer
   Dim I2 As Integer
    Dim RsCia As DAO.Recordset
    
    Set RsCia = Db.OpenRecordset("Stp_AcreedoresPedidosCalendar")
    
    Dim Diap As PowerPoint.Slide
    Set Diap = Pwr.Slides(Variables.No_DiapGraficaAcreedores)
    
    Diap.SlideShowTransition.AdvanceOnClick = msoFalse
    Diap.SlideShowTransition.Hidden = msoTrue
    
    Diap.Shapes("Txt_Fechita").TextFrame.TextRange.Text = "Consumos al " & Variables.Fecha_Consumo
    Dim Tabla As PowerPoint.Shape
    Set Tabla = Diap.Shapes.AddTable(1, 2, CentimetrosAPuntos(1.36), CentimetrosAPuntos(2.85), CentimetrosAPuntos(6.7), CentimetrosAPuntos(0.2))
    Tabla.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Acreedores"
    Tabla.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Montos"
    
    Set Qdf = Db.QueryDefs("Stp_2026_ActividadesCalen_CIAS")
    
    Dim Renglon As Integer
    Dim Totalito As Double
    Dim rngEtiquetas As Range
    Dim Link As PowerPoint.ActionSetting
    Dim Botoncito As PowerPoint.Shape
    Dim Tot_ As Double
    Tot_ = 0
    
    Dim Qdf_TblRes As DAO.QueryDef
    Dim Rs_TblRed As DAO.Recordset
    
    Set Qdf_TblRes = Db.QueryDefs("Stp_Consulta_acreedorParaResumen")
    Qdf_TblRes.Parameters("Consecutivo") = maxConcec
    
    Dim Inicia_Tabla As Integer
    Inicia_Tabla = 3
    
    Dim RequiereTblResumen As Boolean
    Dim Total_TblRes As Double
    
    While Not RsCia.EOF
        RequiereTblResumen = False
        Total_TblRes = 0
        
        Qdf.Parameters("Acreedores") = RsCia("CIA_ABREV")
        Set Rs = Qdf.OpenRecordset
        
        Qdf_TblRes.Parameters("Acreedor") = RsCia("CIA_ABREV")
        Set Rs_TblRed = Qdf_TblRes.OpenRecordset
        
        
        Tabla.Table.Rows.Add
        Tabla.Table.Cell(Tabla.Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = RsCia("CIA_ABREV")
        Tabla.Table.Cell(Tabla.Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(RsCia("Total"), "#,##0.00")
        Tot_ = Tot_ + RsCia("Total")
        Pwr.Slides(Variables.No_DiapGraficaCalendarizado).Duplicate
        
        Set Diapositiva = Pwr.Slides(Variables.No_DiapGraficaCalendarizado)
        Diapositiva.MoveTo (Pwr.Slides.Count - 1)
        
        Diapositiva.SlideShowTransition.AdvanceOnClick = msoFalse
        Diapositiva.SlideShowTransition.Hidden = msoTrue
        
        Diapositiva.Shapes("Txt_Fechita").TextFrame.TextRange.Text = "Consumos al " & Variables.Fecha_Consumo
        Set Link = Diapositiva.Shapes("Btn_Inicio").ActionSettings(ppMouseClick)
        
        Link.Action = ppActionHyperlink
        Link.Hyperlink.SubAddress = Diap.SlideNumber
        
'        Diap.Shapes("Btn_Acreedor").Duplicate
'        Set Botoncito = Diap.Shapes("Btn_Acreedor")
'        Botoncito.Name = "Btn_" & RsCia("CIA_ABREV")
'        Botoncito.ActionSettings(ppMouseClick).Action = ppActionHyperlink
'        Botoncito.ActionSettings(ppMouseClick).Hyperlink.SubAddress = Diapositiva.SlideNumber
'        Botoncito.Left = Tabla.Left
'        Botoncito.Top = Tabla.Top
        
        Tabla.Table.Cell(Tabla.Table.Rows.Count, 1).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick).Action = ppActionHyperlink
        Tabla.Table.Cell(Tabla.Table.Rows.Count, 1).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick).Hyperlink.SubAddress = Diapositiva.SlideNumber
        
        While Not Rs_TblRed.EOF
            Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Add
            Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = Rs_TblRed("Estatus_Actividad")
            Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(Rs_TblRed("Total"), "#,##0.00")
            Total_TblRes = Total_TblRes + Rs_TblRed("Total")
            RequiereTblResumen = True
            Rs_TblRed.MoveNext
        Wend
        
        If Not RequiereTblResumen Then
            Diapositiva.Shapes("Tbl_Resumen").Delete
            
        Else
            Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Add
            Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = "Total"
            Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(Total_TblRes, "#,##0.00")
            Diapositiva.Shapes("Tbl_Resumen").Table.LastRow = True
        End If
        
        Set Grafica = Diapositiva.Shapes("Grafica").Chart
        Set gChartData = Grafica.ChartData
        Set gWorkBook = gChartData.Workbook
        Renglon = Inicia_Tabla
        Totalito = 0
        While Not Rs.EOF
            Renglon = Renglon + 1
       
            gWorkBook.Worksheets(1).Range("A" & Renglon).Value = Rs("Actividad")
            gWorkBook.Worksheets(1).Range("B" & Renglon).Value = Rs("FechaIni")
            gWorkBook.Worksheets(1).Range("D" & Renglon).Value = Rs("Plazo")
            gWorkBook.Worksheets(1).Range("G" & Renglon).Value = Rs("Monto")
            gWorkBook.Worksheets(1).Range("K" & Renglon).Value = Rs("Avance")
            
            gWorkBook.Worksheets(1).Range("B" & Renglon).NumberFormat = "dd/mm/aaaa"
            gWorkBook.Worksheets(1).Range("G" & Renglon).NumberFormat = "_-$* #,##0_-;-$* #,##0_-;_-$* "" - ""??_-;_-@_-"
            gWorkBook.Worksheets(1).Range("H" & Renglon).NumberFormat = "dd/mm/aaaa"
            
            gWorkBook.Worksheets(1).Range("C" & Renglon).Formula = "=((1+SQRT(5))/2)"
            gWorkBook.Worksheets(1).Range("E" & Renglon).Formula = "=((1+SQRT(5))/2)"
            gWorkBook.Worksheets(1).Range("F" & Renglon).Formula = "=(((1+SQRT(5))/2)^10)/SQRT(4)"
            gWorkBook.Worksheets(1).Range("H" & Renglon).Formula = "=CONCATENATE(TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm""),""   "",TEXT(G" & Renglon & ",""#,##0.00""),"" MM$"")" '"=TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm"") & "" "" TEXT(G" & Renglon & ",""#,##0.00"")"
            '"=CONCATENATE(TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm""),""   "",TEXT(G" & Renglon & ",""#,##0.00""),"" MM$"")"
            gWorkBook.Worksheets(1).Range("I" & Renglon).Formula = "=D" & Renglon & "&"" días"""
            'gWorkBook.Worksheets(1).Range("J" & Renglon).Formula = "=TEXT(G" & Renglon & ",""$ #,##0.00"")"
            gWorkBook.Worksheets(1).Range("J" & Renglon).Value = ""
            gWorkBook.Worksheets(1).Range("L" & Renglon).Formula = "=CONCATENATE(TEXT(B" & Renglon & ",""d-mmm""),""   "",I" & Renglon & ")"
            gWorkBook.Worksheets(1).Range("M" & Renglon).Formula = "=if(k" & Renglon & "<>"""",d" & Renglon & "*k" & Renglon & "/100,0)" ' "=d" & Renglon & "*k" & Renglon & "/100"
            gWorkBook.Worksheets(1).Range("N" & Renglon).Formula = "=d" & Renglon & "-m" & Renglon
            gWorkBook.Worksheets(1).Range("O" & Renglon).Formula = "=if(k" & Renglon & "<>"""",concatenate(k" & Renglon & ","" %""),"""")"
            gWorkBook.Worksheets(1).Range("P" & Renglon).Value = RsCia("CIA_ABREV")
            gWorkBook.Worksheets(1).Range("T" & Renglon).Value = Rs("Nombre_Corto")
            gWorkBook.Worksheets(1).Range("U" & Renglon).Value = Rs("Comentarios")
            
            gWorkBook.Worksheets(1).Range("D2").Formula = "=sum(G" & Inicia_Tabla + 1 & ":G" & Renglon & ")"
            
            Totalito = Totalito + Rs("Monto")
            'Grafica.SetSourceData ("=" & gWorkBook.Worksheets(1).Name & "!$b$1:$C$" & i)
            
            Rs.MoveNext
        Wend
        
      
        For Ki = 1 To 45 - (Inicia_Tabla + Renglon)
            gWorkBook.Worksheets(1).Rows(Renglon + 1).EntireRow.Delete
        Next
        
                
        Grafica.SeriesCollection(1).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$B$" & Inicia_Tabla + 1 & ":$B$" & Renglon
        Grafica.SeriesCollection(1).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        
        Grafica.SeriesCollection(2).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$M$" & Inicia_Tabla + 1 & ":$M$" & Renglon
        Grafica.SeriesCollection(2).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        
        Grafica.SeriesCollection(3).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$N$" & Inicia_Tabla + 1 & ":$N$" & Renglon
        Grafica.SeriesCollection(3).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        
'        Grafica.SeriesCollection(3).ApplyDataLabels
'        Grafica.SeriesCollection(3).DataLabels.ShowRange = True
'         Grafica.SeriesCollection(3).HasDataLabels = True
         
         'xlChartFieldUser
         
'         Grafica.SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$L$3:$L$" & Renglon, 0
'        Grafica.SeriesCollection(1).DataLabels.ShowValue = False
'        Grafica.SeriesCollection(1).DataLabels.ShowCategoryName = False
'        Grafica.SeriesCollection(1).DataLabels.ShowPercentage = False
         
        Grafica.SeriesCollection(2).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$O$" & Inicia_Tabla + 1 & ":$O$" & Renglon, 0
        Grafica.SeriesCollection(2).DataLabels.ShowValue = False
        Grafica.SeriesCollection(2).DataLabels.ShowCategoryName = False
        Grafica.SeriesCollection(2).DataLabels.ShowPercentage = False
        
'        Grafica.SeriesCollection(3).DataLabels.ShowValue = False
'        Grafica.SeriesCollection(3).DataLabels.ShowCategoryName = False
'        Grafica.SeriesCollection(3).DataLabels.ShowPercentage = False
        
        Grafica.SeriesCollection(4).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$E$" & Inicia_Tabla + 1 & ":$E$" & Renglon
        Grafica.SeriesCollection(4).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        Grafica.SeriesCollection(4).ApplyDataLabels
        Grafica.SeriesCollection(4).DataLabels.ShowRange = True
         Grafica.SeriesCollection(4).HasDataLabels = True
         
         'xlChartFieldUser
        Grafica.SeriesCollection(4).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$H$" & Inicia_Tabla + 1 & ":$H$" & Renglon, 0
        Grafica.SeriesCollection(4).DataLabels.ShowValue = False
        Grafica.SeriesCollection(4).DataLabels.ShowCategoryName = False
        Grafica.SeriesCollection(4).DataLabels.ShowPercentage = False
'        Grafica.SeriesCollection(4).DataLabels.
        
        Grafica.SeriesCollection(5).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$F$" & Inicia_Tabla + 1 & ":$F$" & Renglon
        Grafica.SeriesCollection(5).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
          Grafica.SeriesCollection(5).ApplyDataLabels
        Grafica.SeriesCollection(5).DataLabels.ShowRange = True
         Grafica.SeriesCollection(5).HasDataLabels = True
         
'         xlChartFieldUser
        Grafica.SeriesCollection(5).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$J$" & Inicia_Tabla + 1 & ":$J$" & Renglon, 0
        Grafica.SeriesCollection(5).DataLabels.ShowValue = False
        Grafica.SeriesCollection(5).DataLabels.ShowCategoryName = False
        Grafica.SeriesCollection(5).DataLabels.ShowPercentage = False

        Grafica.Refresh
        
        If Renglon - Inicia_Tabla = 1 Then
            Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
            Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(3.26)
        End If
        
        If Renglon - Inicia_Tabla = 2 Then
            Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
            Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(4.41)
        End If
        
        
          If Renglon - Inicia_Tabla = 3 Then
            Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
            Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(6.05)
        End If
        
        
        
        Diapositiva.Shapes("Txt_Total").TextFrame.TextRange.Text = "Total " & Format(Totalito, "#,##0.00") & " Millones de Pesos"
        Diapositiva.Shapes("Título 1").TextFrame.TextRange.Text = "Actividades Calendarizadas de " & RsCia("CIA_ABREV")
      
        
    
        For Ki = 1 To Renglon - Inicia_Tabla
            'MsgBox Grafica.SeriesCollection(4).Points(Ki).Top & " " & Diapositiva.Shapes(Grafica.Name).Top
            If gWorkBook.Worksheets(1).Range("U" & Ki + Inicia_Tabla).Value <> "" Then
                Diapositiva.Shapes("Txt_Comentario").Duplicate
                Diapositiva.Shapes("Flechita").Duplicate
                Set TxtComent = Diapositiva.Shapes("Txt_Comentario")
                Set Flechita = Diapositiva.Shapes("Flechita")
                TxtComent.Name = "Txt_Comentario_" & Ki
                Flechita.Name = "Flechita_" & Ki
                Ancho_Barrita = Grafica.SeriesCollection(4).Points(Ki).Height
                Ancho_Coment = TxtComent.Height
                
                TxtComent.TextFrame.TextRange.Text = gWorkBook.Worksheets(1).Range("U" & Ki + Inicia_Tabla).Value
                TxtComent.Top = Grafica.SeriesCollection(4).Points(Ki).Top + Diapositiva.Shapes(Grafica.Name).Top + ((Ancho_Barrita - Ancho_Coment) / 2)
                TxtComent.Left = Diapositiva.Shapes(Grafica.Name).Left + Diapositiva.Shapes(Grafica.Name).Width + 30
                Flechita.Top = Grafica.SeriesCollection(4).Points(Ki).Top + Diapositiva.Shapes(Grafica.Name).Top + (Ancho_Barrita / 2)
                Flechita.Width = 28
                Flechita.Left = Grafica.SeriesCollection(4).DataLabels(Ki).Width + Grafica.SeriesCollection(4).DataLabels(Ki).Left + 15
                Flechita.Width = TxtComent.Left - (Grafica.SeriesCollection(4).DataLabels(Ki).Width + Grafica.SeriesCollection(4).DataLabels(Ki).Left + 15)
            End If
        Next
        
          gWorkBook.Close
        
        Diapositiva.Shapes("Txt_Comentario").Delete
        Diapositiva.Shapes("Flechita").Delete
        
        For i = 1 To Grafica.SeriesCollection(1).DataLabels.Count
            If Grafica.SeriesCollection(1).DataLabels(i).Left + Grafica.SeriesCollection(1).DataLabels(i).Width > Grafica.SeriesCollection(2).DataLabels(i).Left Then
'                MsgBox "Se trasponen las cajitas"
                Grafica.SeriesCollection(1).DataLabels(i).Left = Grafica.SeriesCollection(1).DataLabels(i).Left - ((Grafica.SeriesCollection(1).DataLabels(i).Left + Grafica.SeriesCollection(1).DataLabels(i).Width) - Grafica.SeriesCollection(2).DataLabels(i).Left)
            End If
        Next
        
        
        Set Grafica = Nothing
        Set gChartData = Nothing
        Set gWorkBook = Nothing

        RsCia.MoveNext
    Wend
    
    Tabla.Table.Rows.Add
    Tabla.Table.Cell(Tabla.Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = "Total"
    Tabla.Table.Cell(Tabla.Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(Tot_, "#,##0.00")
    Tabla.Table.LastRow = True
    
    Dim Rng_Tbl As Integer
    Dim Col_Tbl As Integer
    
    For Rng_Tbl = 1 To Tabla.Table.Rows.Count
        For Col_Tbl = 1 To Tabla.Table.Columns.Count
            Tabla.Table.Cell(Rng_Tbl, Col_Tbl).Shape.TextFrame.TextRange.Font.Size = 9
            Tabla.Table.Cell(Rng_Tbl, Col_Tbl).Shape.TextFrame.MarginBottom = 0
            Tabla.Table.Cell(Rng_Tbl, Col_Tbl).Shape.TextFrame.MarginTop = 0
'        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginLeft = 2
'        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginRight = 2
'        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginBottom = 2
'        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginTop = 2
        Next
    Next
    Tabla.Height = 0
    Tabla.Table.Columns(1).Width = CentimetrosAPuntos(5.8)
    Tabla.Table.Columns(2).Width = CentimetrosAPuntos(1.9)
    
    
    Set Grafica = Diap.Shapes("Grafico").Chart
    Set gChartData = Grafica.ChartData
    Set gWorkBook = gChartData.Workbook
    Renglon = Inicia_Tabla
    Set Rs = Db.OpenRecordset("Stp_2026_ActivPorMes")
    While Not Rs.EOF
        gWorkBook.Worksheets(1).Range("B" & Rs("Pos") + 1).Value = Rs("Total")
        gWorkBook.Worksheets(1).Range("B" & Rs("Pos") + 1).NumberFormat = "#,##0.00"
        Rs.MoveNext
    Wend
    
    
    Grafica.SetSourceData ("=" & gWorkBook.Worksheets(1).Name & "!$A$1:$B$13")
    Grafica.Refresh
    
    gWorkBook.Close
    Set Grafica = Nothing
    Set gChartData = Nothing
    Set gWorkBook = Nothing
    
    Set RsCia = Db.OpenRecordset("Stp_2026_ActivPorSupervisores")
    
    Dim Tbl2 As PowerPoint.Shape
    Set Tbl2 = Diap.Shapes.AddTable(1, 2, CentimetrosAPuntos(26.08), CentimetrosAPuntos(2.85), CentimetrosAPuntos(6.7), CentimetrosAPuntos(0.2))
    
    Tbl2.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Supervisores Técnicos"
    Tbl2.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Montos"
    
     Set Qdf = Db.QueryDefs("Stp_2026_ActividadesCalen_RESPONSABLES")
    While Not RsCia.EOF
        Tbl2.Table.Rows.Add
        Tbl2.Table.Cell(Tbl2.Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = RsCia("Nombre_Corto")
        Tbl2.Table.Cell(Tbl2.Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(RsCia("Total"), "#,##0.00")
        
        Qdf.Parameters("Responsable") = RsCia("Nombre_Corto")
        Set Rs = Qdf.OpenRecordset
        
        Pwr.Slides(Variables.No_DiapGraficaCalendarizado).Duplicate
        
        Set Diapositiva = Pwr.Slides(Variables.No_DiapGraficaCalendarizado)
        Diapositiva.MoveTo (Pwr.Slides.Count - 1)
        Diapositiva.Shapes("Txt_Fechita").TextFrame.TextRange.Text = "Consumos al " & Variables.Fecha_Consumo
        
        Diapositiva.SlideShowTransition.AdvanceOnClick = msoFalse
        Diapositiva.SlideShowTransition.Hidden = msoTrue
        
        Set Link = Diapositiva.Shapes("Btn_Inicio").ActionSettings(ppMouseClick)
        
        Link.Action = ppActionHyperlink
        Link.Hyperlink.SubAddress = Diap.SlideNumber
        
        Tbl2.Table.Cell(Tbl2.Table.Rows.Count, 1).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick).Action = ppActionHyperlink
        Tbl2.Table.Cell(Tbl2.Table.Rows.Count, 1).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick).Hyperlink.SubAddress = Diapositiva.SlideNumber
        
'        Diap.Shapes("Btn_Acreedor").Duplicate
'        Set Botoncito = Diap.Shapes("Btn_Acreedor")
'        Botoncito.Name = "Btn_" & RsCia("Nombre_Corto")
'        Botoncito.ActionSettings(ppMouseClick).Action = ppActionHyperlink
'        Botoncito.ActionSettings(ppMouseClick).Hyperlink.SubAddress = Diapositiva.SlideNumber
'        Botoncito.Left = Tbl2.Left
'        Botoncito.Top = Tbl2.Top
        
        Set Grafica = Diapositiva.Shapes("Grafica").Chart
        Set gChartData = Grafica.ChartData
        Set gWorkBook = gChartData.Workbook
        Renglon = Inicia_Tabla
        Totalito = 0
        While Not Rs.EOF
            Renglon = Renglon + 1
                        
            gWorkBook.Worksheets(1).Range("A" & Renglon).Value = Rs("Actividad")
            gWorkBook.Worksheets(1).Range("B" & Renglon).Value = Rs("FechaIni")
            gWorkBook.Worksheets(1).Range("D" & Renglon).Value = Rs("Plazo")
            gWorkBook.Worksheets(1).Range("G" & Renglon).Value = Rs("Monto")
            gWorkBook.Worksheets(1).Range("K" & Renglon).Value = Rs("Avance")
            
            gWorkBook.Worksheets(1).Range("B" & Renglon).NumberFormat = "dd/mm/aaaa"
            gWorkBook.Worksheets(1).Range("G" & Renglon).NumberFormat = "_-$* #,##0_-;-$* #,##0_-;_-$* "" - ""??_-;_-@_-"
            gWorkBook.Worksheets(1).Range("H" & Renglon).NumberFormat = "dd/mm/aaaa"
            
            gWorkBook.Worksheets(1).Range("C" & Renglon).Formula = "=((1+SQRT(5))/2)"
            gWorkBook.Worksheets(1).Range("E" & Renglon).Formula = "=((1+SQRT(5))/2)"
            gWorkBook.Worksheets(1).Range("F" & Renglon).Formula = "=(((1+SQRT(5))/2)^10)/SQRT(4)"
            gWorkBook.Worksheets(1).Range("H" & Renglon).Formula = "=CONCATENATE(TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm""),""   "",TEXT(G" & Renglon & ",""#,##0.00""),"" MM$"")" '"=TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm"") & "" "" TEXT(G" & Renglon & ",""#,##0.00"")"'"=TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm"")"
            gWorkBook.Worksheets(1).Range("I" & Renglon).Formula = "=D" & Renglon & "&"" días"""
'            gWorkBook.Worksheets(1).Range("J" & Renglon).Formula = "=TEXT(G" & Renglon & ",""$ #,##0.00"")"
            gWorkBook.Worksheets(1).Range("J" & Renglon).Value = ""
             gWorkBook.Worksheets(1).Range("L" & Renglon).Formula = "=CONCATENATE(TEXT(B" & Renglon & ",""d-mmm""),""   "",I" & Renglon & ")"
            gWorkBook.Worksheets(1).Range("M" & Renglon).Formula = "=if(k" & Renglon & "<>"""",d" & Renglon & "*k" & Renglon & "/100,0)" '  "=d" & Renglon & "*k" & Renglon & "/100"
            gWorkBook.Worksheets(1).Range("N" & Renglon).Formula = "=d" & Renglon & "-m" & Renglon
            gWorkBook.Worksheets(1).Range("O" & Renglon).Formula = "=if(k" & Renglon & "<>"""",concatenate(k" & Renglon & ","" %""),"""")"
            
            gWorkBook.Worksheets(1).Range("P" & Renglon).Value = Rs("CIA_ABREV")
            gWorkBook.Worksheets(1).Range("T" & Renglon).Value = RsCia("Nombre_Corto")
            gWorkBook.Worksheets(1).Range("U" & Renglon).Value = Rs("Comentarios")
            
            gWorkBook.Worksheets(1).Range("D2").Formula = "=sum(G" & Inicia_Tabla + 1 & ":G" & Renglon & ")"
            
            Totalito = Totalito + Rs("Monto")
            'Grafica.SetSourceData ("=" & gWorkBook.Worksheets(1).Name & "!$b$1:$C$" & i)
            
            Rs.MoveNext
        Wend
        
        For Ki = 1 To 45 - (Inicia_Tabla + Renglon)
            gWorkBook.Worksheets(1).Rows(Renglon + 1).EntireRow.Delete
        Next
        
        
        Grafica.SeriesCollection(1).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$B$" & Inicia_Tabla + 1 & ":$B$" & Renglon
        Grafica.SeriesCollection(1).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        
        Grafica.SeriesCollection(2).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$M$" & Inicia_Tabla + 1 & ":$M$" & Renglon
        Grafica.SeriesCollection(2).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        
        Grafica.SeriesCollection(3).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$N$" & Inicia_Tabla + 1 & ":$N$" & Renglon
        Grafica.SeriesCollection(3).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        
'        Grafica.SeriesCollection(3).ApplyDataLabels
'        Grafica.SeriesCollection(3).DataLabels.ShowRange = True
'         Grafica.SeriesCollection(3).HasDataLabels = True
         
         'xlChartFieldUser
'
'        Grafica.SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$L$3:$L$" & Renglon, 0
'        Grafica.SeriesCollection(1).DataLabels.ShowValue = False
'        Grafica.SeriesCollection(1).DataLabels.ShowCategoryName = False
'        Grafica.SeriesCollection(1).DataLabels.ShowPercentage = False
        
        Grafica.SeriesCollection(2).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$O$" & Inicia_Tabla + 1 & ":$O$" & Renglon, 0
        Grafica.SeriesCollection(2).DataLabels.ShowValue = False
        Grafica.SeriesCollection(2).DataLabels.ShowCategoryName = False
        Grafica.SeriesCollection(2).DataLabels.ShowPercentage = False
        
'        Grafica.SeriesCollection(3).DataLabels.ShowValue = False
'        Grafica.SeriesCollection(3).DataLabels.ShowCategoryName = False
'        Grafica.SeriesCollection(3).DataLabels.ShowPercentage = False

                
        Grafica.SeriesCollection(4).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$E$" & Inicia_Tabla + 1 & ":$E$" & Renglon
        Grafica.SeriesCollection(4).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        Grafica.SeriesCollection(4).ApplyDataLabels
        Grafica.SeriesCollection(4).DataLabels.ShowRange = True
         Grafica.SeriesCollection(4).HasDataLabels = True
         
         'xlChartFieldUser
        Grafica.SeriesCollection(4).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$H$" & Inicia_Tabla + 1 & ":$H$" & Renglon, 0
        Grafica.SeriesCollection(4).DataLabels.ShowValue = False
        Grafica.SeriesCollection(4).DataLabels.ShowCategoryName = False
        Grafica.SeriesCollection(4).DataLabels.ShowPercentage = False
'        Grafica.SeriesCollection(4).DataLabels.
        
        Grafica.SeriesCollection(5).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$F$" & Inicia_Tabla + 1 & ":$F$" & Renglon
        Grafica.SeriesCollection(5).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
          Grafica.SeriesCollection(5).ApplyDataLabels
        Grafica.SeriesCollection(5).DataLabels.ShowRange = True
         Grafica.SeriesCollection(5).HasDataLabels = True
         
         'xlChartFieldUser
        Grafica.SeriesCollection(5).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$J$" & Inicia_Tabla + 1 & ":$J$" & Renglon, 0
        Grafica.SeriesCollection(5).DataLabels.ShowValue = False
        Grafica.SeriesCollection(5).DataLabels.ShowCategoryName = False
        Grafica.SeriesCollection(5).DataLabels.ShowPercentage = False

        Grafica.Refresh
        
        If Renglon - Inicia_Tabla = 1 Then
          Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
          Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(3.26)
        End If
        
        If Renglon - Inicia_Tabla = 2 Then
          Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
          Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(4.41)
        End If
        
        If Renglon - Inicia_Tabla = 3 Then
          Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
          Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(6.05)
        End If
        
        Diapositiva.Shapes("Txt_Total").TextFrame.TextRange.Text = "Total " & Format(Totalito, "#,##0.00") & " Millones de Pesos"
        Diapositiva.Shapes("Título 1").TextFrame.TextRange.Text = "Actividades Calendarizadas de " & RsCia("Nombre_Corto")
        
        For Ki = 1 To Renglon - Inicia_Tabla
            'MsgBox Grafica.SeriesCollection(4).Points(Ki).Top & " " & Diapositiva.Shapes(Grafica.Name).Top
            If gWorkBook.Worksheets(1).Range("U" & Ki + Inicia_Tabla).Value <> "" Then
                Diapositiva.Shapes("Txt_Comentario").Duplicate
                Diapositiva.Shapes("Flechita").Duplicate
                Set TxtComent = Diapositiva.Shapes("Txt_Comentario")
                Set Flechita = Diapositiva.Shapes("Flechita")
                TxtComent.Name = "Txt_Comentario_" & Ki
                Flechita.Name = "Flechita_" & Ki
                Ancho_Barrita = Grafica.SeriesCollection(4).Points(Ki).Height
                Ancho_Coment = TxtComent.Height
                
                TxtComent.TextFrame.TextRange.Text = gWorkBook.Worksheets(1).Range("U" & Ki + Inicia_Tabla).Value
                TxtComent.Top = Grafica.SeriesCollection(4).Points(Ki).Top + Diapositiva.Shapes(Grafica.Name).Top + ((Ancho_Barrita - Ancho_Coment) / 2)
                TxtComent.Left = Diapositiva.Shapes(Grafica.Name).Left + Diapositiva.Shapes(Grafica.Name).Width + 30
                Flechita.Top = Grafica.SeriesCollection(4).Points(Ki).Top + Diapositiva.Shapes(Grafica.Name).Top + (Ancho_Barrita / 2)
                Flechita.Width = 28
                Flechita.Left = Grafica.SeriesCollection(4).DataLabels(Ki).Width + Grafica.SeriesCollection(4).DataLabels(Ki).Left + 15
                Flechita.Width = TxtComent.Left - (Grafica.SeriesCollection(4).DataLabels(Ki).Width + Grafica.SeriesCollection(4).DataLabels(Ki).Left + 15)
            End If
        Next
        
        
        gWorkBook.Close
        
        Diapositiva.Shapes("Txt_Comentario").Delete
        Diapositiva.Shapes("Flechita").Delete
        
        
        For i = 1 To Grafica.SeriesCollection(1).DataLabels.Count
            If Grafica.SeriesCollection(1).DataLabels(i).Left + Grafica.SeriesCollection(1).DataLabels(i).Width > Grafica.SeriesCollection(2).DataLabels(i).Left Then
'                MsgBox "Se trasponen las cajitas"
                Grafica.SeriesCollection(1).DataLabels(i).Left = Grafica.SeriesCollection(1).DataLabels(i).Left - ((Grafica.SeriesCollection(1).DataLabels(i).Left + Grafica.SeriesCollection(1).DataLabels(i).Width) - Grafica.SeriesCollection(2).DataLabels(i).Left)
            End If
        Next
        
        
        Set Grafica = Nothing
        Set gChartData = Nothing
        Set gWorkBook = Nothing
        
        RsCia.MoveNext
    Wend
        
    Tbl2.Table.Rows.Add
    Tbl2.Table.Cell(Tbl2.Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = "Total"
    Tbl2.Table.Cell(Tbl2.Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(Tot_, "#,##0.00")
    Tbl2.Table.LastRow = True
    
    
    For Rng_Tbl = 1 To Tbl2.Table.Rows.Count
        For Col_Tbl = 1 To Tbl2.Table.Columns.Count
            Tbl2.Table.Cell(Rng_Tbl, Col_Tbl).Shape.TextFrame.TextRange.Font.Size = 9
            Tbl2.Table.Cell(Rng_Tbl, Col_Tbl).Shape.TextFrame.MarginBottom = 0
            Tbl2.Table.Cell(Rng_Tbl, Col_Tbl).Shape.TextFrame.MarginTop = 0
        '        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginLeft = 2
        '        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginRight = 2
        '        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginBottom = 2
        '        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginTop = 2
        Next
    Next
    Tbl2.Height = 0
    
    Tbl2.Table.Columns(1).Width = CentimetrosAPuntos(5.1)
    Tbl2.Table.Columns(2).Width = CentimetrosAPuntos(1.9)
    
    Set Grafica = Diap.Shapes("Grafico").Chart
     Set Qdf = Db.QueryDefs("Stp_2026_ActivCalendarizadasCIA_Responsable")
    ' Dim i As Integer
     Dim Graf As PowerPoint.Chart
     
    Set Qdf_TblRes = Db.QueryDefs("Stp_TablaResumenMensual")
    Qdf_TblRes.Parameters("Consecutivo") = maxConcec
     
    For i = 1 To Grafica.SeriesCollection(1).DataLabels.Count
        
        If Grafica.SeriesCollection(1).DataLabels(i).Text <> "" Then
            RequiereTblResumen = False
            Diap.Shapes("Btn_Acreedor").Duplicate
            Set Botoncito = Diap.Shapes("Btn_Acreedor")
            
            Botoncito.Name = "Btn_" & i
            
            Botoncito.Top = Grafica.SeriesCollection(1).DataLabels(i).Top + Diap.Shapes("Grafico").Top + Botoncito.Width
            Botoncito.Left = Grafica.SeriesCollection(1).DataLabels(i).Left + Diap.Shapes("Grafico").Left

            Qdf.Parameters("No_mes") = i
            Set Rs = Qdf.OpenRecordset
            
            Qdf_TblRes.Parameters("No_Mes") = i
             Set Rs_TblRed = Qdf_TblRes.OpenRecordset
            
            
            
            
            '***************************************
            
            Pwr.Slides(Variables.No_DiapGraficaCalendarizado).Duplicate
            
           
            
            Set Diapositiva = Pwr.Slides(Variables.No_DiapGraficaCalendarizado)
            Diapositiva.MoveTo (Pwr.Slides.Count - 1)
            Diapositiva.Shapes("Txt_Fechita").TextFrame.TextRange.Text = "Consumos al " & Variables.Fecha_Consumo
            
            Diapositiva.SlideShowTransition.AdvanceOnClick = msoFalse
            Diapositiva.SlideShowTransition.Hidden = msoTrue
            
            Set Link = Diapositiva.Shapes("Btn_Inicio").ActionSettings(ppMouseClick)
            
            Link.Action = ppActionHyperlink
            Link.Hyperlink.SubAddress = Diap.SlideNumber
            
            Botoncito.ActionSettings(ppMouseClick).Action = ppActionHyperlink
            Botoncito.ActionSettings(ppMouseClick).Hyperlink.SubAddress = Diapositiva.SlideNumber
            
            
         
            
            
            Set Graf = Diapositiva.Shapes("Grafica").Chart
            Set gChartData = Graf.ChartData
            Set gWorkBook = gChartData.Workbook
            Renglon = Inicia_Tabla
            Totalito = 0
            While Not Rs.EOF
              Renglon = Renglon + 1
                
                gWorkBook.Worksheets(1).Range("A" & Renglon).Value = Rs("ActividadCorta") & "/" & Rs("CIA_ABREV") & vbNewLine & Rs("Nombre_Corto")
                
                gWorkBook.Worksheets(1).Range("B" & Renglon).Value = Rs("FechaIni")
                gWorkBook.Worksheets(1).Range("D" & Renglon).Value = Rs("Plazo")
                gWorkBook.Worksheets(1).Range("G" & Renglon).Value = Rs("Monto")
                gWorkBook.Worksheets(1).Range("K" & Renglon).Value = Rs("Avance")
                
                gWorkBook.Worksheets(1).Range("B" & Renglon).NumberFormat = "dd/mm/aaaa"
                gWorkBook.Worksheets(1).Range("G" & Renglon).NumberFormat = "_-$* #,##0_-;-$* #,##0_-;_-$* "" - ""??_-;_-@_-"
                gWorkBook.Worksheets(1).Range("H" & Renglon).NumberFormat = "dd/mm/aaaa"
                
                gWorkBook.Worksheets(1).Range("C" & Renglon).Formula = "=((1+SQRT(5))/2)"
                gWorkBook.Worksheets(1).Range("E" & Renglon).Formula = "=((1+SQRT(5))/2)"
                gWorkBook.Worksheets(1).Range("F" & Renglon).Formula = "=(((1+SQRT(5))/2)^10)/SQRT(4)"
                gWorkBook.Worksheets(1).Range("H" & Renglon).Formula = "=CONCATENATE(TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm""),""   "",TEXT(G" & Renglon & ",""#,##0.00""),"" MM$"")" '"=TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm"") & "" "" TEXT(G" & Renglon & ",""#,##0.00"")"'"=TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm"")"
                gWorkBook.Worksheets(1).Range("I" & Renglon).Formula = "=D" & Renglon & "&"" días"""
                '              gWorkBook.Worksheets(1).Range("J" & Renglon).Formula = "=TEXT(G" & Renglon & ",""$ #,##0.00"")"
                gWorkBook.Worksheets(1).Range("J" & Renglon).Value = ""
                gWorkBook.Worksheets(1).Range("L" & Renglon).Formula = "=CONCATENATE(TEXT(B" & Renglon & ",""d-mmm""),""   "",I" & Renglon & ")"
                gWorkBook.Worksheets(1).Range("M" & Renglon).Formula = "=if(k" & Renglon & "<>"""",d" & Renglon & "*k" & Renglon & "/100,0)" ' "=d" & Renglon & "*k" & Renglon & "/100"
                gWorkBook.Worksheets(1).Range("N" & Renglon).Formula = "=d" & Renglon & "-m" & Renglon
                gWorkBook.Worksheets(1).Range("O" & Renglon).Formula = "=if(k" & Renglon & "<>"""",concatenate(k" & Renglon & ","" %""),"""")"
                
                gWorkBook.Worksheets(1).Range("P" & Renglon).Value = Rs("CIA_ABREV")
                gWorkBook.Worksheets(1).Range("S" & Renglon).Value = Rs("ActividadCorta")
                gWorkBook.Worksheets(1).Range("T" & Renglon).Value = Rs("Nombre_Corto")
                gWorkBook.Worksheets(1).Range("U" & Renglon).Value = Rs("Comentarios")
                
                gWorkBook.Worksheets(1).Range("D2").Formula = "=sum(G" & Inicia_Tabla + 1 & ":G" & Renglon & ")"
                
              Totalito = Totalito + Rs("Monto")
              'Grafica.SetSourceData ("=" & gWorkBook.Worksheets(1).Name & "!$b$1:$C$" & i)
              
              Rs.MoveNext
            Wend
            
            For Ki = 1 To 45 - (Inicia_Tabla + Renglon)
                gWorkBook.Worksheets(1).Rows(Renglon + 1).EntireRow.Delete
            Next
            
            Graf.SeriesCollection(1).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$B$" & Inicia_Tabla + 1 & ":$B$" & Renglon
            Graf.SeriesCollection(1).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
            
            Graf.SeriesCollection(2).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$M$" & Inicia_Tabla + 1 & ":$M$" & Renglon
            Graf.SeriesCollection(2).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
            
            Graf.SeriesCollection(3).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$N$" & Inicia_Tabla + 1 & ":$N$" & Renglon
            Graf.SeriesCollection(3).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
            
'            Graf.SeriesCollection(3).ApplyDataLabels
'            Graf.SeriesCollection(3).DataLabels.ShowRange = True
'            Graf.SeriesCollection(3).HasDataLabels = True
            
            'xlChartFieldUser
'            Grafica.SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$L$3:$L$" & Renglon, 0
'            Grafica.SeriesCollection(1).DataLabels.ShowValue = False
'            Grafica.SeriesCollection(1).DataLabels.ShowCategoryName = False
'            Grafica.SeriesCollection(1).DataLabels.ShowPercentage = False
            
            On Error Resume Next
            Graf.SeriesCollection(2).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$O$" & Inicia_Tabla + 1 & ":$O$" & Renglon, 0
            Graf.SeriesCollection(2).DataLabels.ShowValue = False
            Graf.SeriesCollection(2).DataLabels.ShowCategoryName = False
            Graf.SeriesCollection(2).DataLabels.ShowPercentage = False
            
'            Grafica.SeriesCollection(3).DataLabels.ShowValue = False
'            Grafica.SeriesCollection(3).DataLabels.ShowCategoryName = False
'            Grafica.SeriesCollection(3).DataLabels.ShowPercentage = False
            
            
            Graf.SeriesCollection(4).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$E$" & Inicia_Tabla + 1 & ":$E$" & Renglon
            Graf.SeriesCollection(4).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
            Graf.SeriesCollection(4).ApplyDataLabels
            Graf.SeriesCollection(4).DataLabels.ShowRange = True
            Graf.SeriesCollection(4).HasDataLabels = True
            
            'xlChartFieldUser
            Graf.SeriesCollection(4).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$H$" & Inicia_Tabla + 1 & ":$H$" & Renglon, 0
            Graf.SeriesCollection(4).DataLabels.ShowValue = False
            Graf.SeriesCollection(4).DataLabels.ShowCategoryName = False
            Graf.SeriesCollection(4).DataLabels.ShowPercentage = False
            '        Grafica.SeriesCollection(4).DataLabels.
            
            Graf.SeriesCollection(5).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$F$" & Inicia_Tabla + 1 & ":$F$" & Renglon
            Graf.SeriesCollection(5).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
            Graf.SeriesCollection(5).ApplyDataLabels
            Graf.SeriesCollection(5).DataLabels.ShowRange = True
            Graf.SeriesCollection(5).HasDataLabels = True
            
            'xlChartFieldUser
            Graf.SeriesCollection(5).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$J$" & Inicia_Tabla + 1 & ":$J$" & Renglon, 0
            Graf.SeriesCollection(5).DataLabels.ShowValue = False
            Graf.SeriesCollection(5).DataLabels.ShowCategoryName = False
            Graf.SeriesCollection(5).DataLabels.ShowPercentage = False
            
            Graf.Refresh
            
            If Renglon - Inicia_Tabla = 1 Then
                Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
                Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(3.26)
            End If
            
            If Renglon - Inicia_Tabla = 2 Then
                Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
                Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(4.41)
            End If
            
            
              If Renglon - Inicia_Tabla = 3 Then
                Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
                Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(6.05)
            End If
        
            Diapositiva.Shapes("Txt_Total").TextFrame.TextRange.Text = "Total " & Format(Totalito, "#,##0.00") & " Millones de Pesos"
            Diapositiva.Shapes("Título 1").TextFrame.TextRange.Text = "Actividades programadas para terminar en " & Nombre_Mes(i)
            
            
            Total_TblRes = 0
            While Not Rs_TblRed.EOF
                Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Add
                Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = Rs_TblRed("Estatus_Actividad")
                Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(Rs_TblRed("Total"), "#,##0.00")
                Total_TblRes = Total_TblRes + Rs_TblRed("Total")
                RequiereTblResumen = True
                Rs_TblRed.MoveNext
            Wend
            
            If Not RequiereTblResumen Then
                Diapositiva.Shapes("Tbl_Resumen").Delete
                
            Else
                Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Add
                Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = "Total"
                Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(Total_TblRes, "#,##0.00")
                Diapositiva.Shapes("Tbl_Resumen").Table.LastRow = True
            End If
            
            
            
            
            For Ki = 1 To Renglon - Inicia_Tabla
            'MsgBox Grafica.SeriesCollection(4).Points(Ki).Top & " " & Diapositiva.Shapes(Grafica.Name).Top
                If gWorkBook.Worksheets(1).Range("U" & Ki + Inicia_Tabla).Value <> "" Then
                    Diapositiva.Shapes("Txt_Comentario").Duplicate
                    Diapositiva.Shapes("Flechita").Duplicate
                    Set TxtComent = Diapositiva.Shapes("Txt_Comentario")
                    Set Flechita = Diapositiva.Shapes("Flechita")
                    TxtComent.Name = "Txt_Comentario_" & Ki
                    Flechita.Name = "Flechita_" & Ki
                    Ancho_Barrita = Graf.SeriesCollection(4).Points(Ki).Height
                    Ancho_Coment = TxtComent.Height
                    
                    TxtComent.TextFrame.TextRange.Text = gWorkBook.Worksheets(1).Range("U" & Ki + Inicia_Tabla).Value
                    TxtComent.Top = Graf.SeriesCollection(4).Points(Ki).Top + Diapositiva.Shapes(Graf.Name).Top + ((Ancho_Barrita - Ancho_Coment) / 2)
                    TxtComent.Left = Diapositiva.Shapes(Graf.Name).Left + Diapositiva.Shapes(Graf.Name).Width + 30
                    Flechita.Top = Graf.SeriesCollection(4).Points(Ki).Top + Diapositiva.Shapes(Graf.Name).Top + (Ancho_Barrita / 2)
                    
                    Flechita.Left = Graf.SeriesCollection(4).DataLabels(Ki).Width + Graf.SeriesCollection(4).DataLabels(Ki).Left + 15
                    Flechita.Width = TxtComent.Left - (Graf.SeriesCollection(4).DataLabels(Ki).Width + Graf.SeriesCollection(4).DataLabels(Ki).Left + 15)
                End If
            Next
        
            
              gWorkBook.Close
              
            Diapositiva.Shapes("Txt_Comentario").Delete
            Diapositiva.Shapes("Flechita").Delete
            
        For I2 = 1 To Graf.SeriesCollection(1).DataLabels.Count
            If Graf.SeriesCollection(1).DataLabels(I2).Left + Graf.SeriesCollection(1).DataLabels(I2).Width > Graf.SeriesCollection(2).DataLabels(I2).Left Then
'                MsgBox "Se trasponen las cajitas"
                Graf.SeriesCollection(1).DataLabels(I2).Left = Graf.SeriesCollection(1).DataLabels(I2).Left - ((Graf.SeriesCollection(1).DataLabels(I2).Left + Graf.SeriesCollection(1).DataLabels(I2).Width) - Graf.SeriesCollection(2).DataLabels(I2).Left)
            End If
        Next
            
            Set Graf = Nothing
            Set gChartData = Nothing
            Set gWorkBook = Nothing
            
            '***************************************
        End If
    'Grafica.SeriesCollection(2).DataLabels(i).Font.Size = 11
    Next
'        gWorkBook.Close
    Set Grafica = Nothing
    Set gChartData = Nothing
    Set gWorkBook = Nothing
    
    Set Qdf = Nothing
    Set Rs = Nothing
    Set RsCia = Nothing
    Set Diap = Nothing
    Set Diapositiva = Nothing
End Sub

Public Function Nombre_Mes(NumMes As Integer) As String
    Select Case NumMes
        Case 1
            Nombre_Mes = "Enero"
        Case 2
            Nombre_Mes = "Febrero"
        Case 3
            Nombre_Mes = "Marzo"
        Case 4
            Nombre_Mes = "Abril"
        Case 5
            Nombre_Mes = "Mayo"
        Case 6
            Nombre_Mes = "Junio"
        Case 7
            Nombre_Mes = "Julio"
        Case 8
            Nombre_Mes = "Agosto"
        Case 9
            Nombre_Mes = "Septiembre"
        Case 10
            Nombre_Mes = "Octubre"
        Case 11
            Nombre_Mes = "Noviembre"
        Case 12
            Nombre_Mes = "Diciembre"
    End Select
    
End Function


Sub Soporte_Variacion(Pwr As PowerPoint.Presentation)
    '
    '
     Dim Diapositiva As PowerPoint.Slide
    Dim Num_DiapRetorno As Integer
    
    Dim Qdf As DAO.QueryDef
    Dim Rs As DAO.Recordset
    
     Set Qdf = Db.QueryDefs("Stp_2026_VariacionenelPeriodo")
    Qdf.Parameters("Id_Adecuado") = Id_AdecuadoOficial
    Qdf.Parameters("Fecha") = Variables.Periodo_Fin
    
    Set Rs = Qdf.OpenRecordset
       
   
     Dim Estruc(4) As Proceso.Tipo_Tabla
     
    
    Estruc(1).Alineacion = 4
    Estruc(1).Encabezado = "Acreedores"
    Estruc(1).Size_Fuente = 9
    Estruc(1).Ancho = CentimetrosAPuntos(4.1)
    Estruc(1).Fuente = FuenteOficial
    Estruc(1).SeSuma = False
    Estruc(1).Tipo = "Cadena"
    Estruc(1).Total = 0
    
    Estruc(2).Alineacion = 2
    Estruc(2).Encabezado = "Programado"
    Estruc(2).Size_Fuente = 9
    Estruc(2).Ancho = CentimetrosAPuntos(2.1)
    Estruc(2).Fuente = FuenteOficial
    Estruc(2).SeSuma = True
    Estruc(2).Tipo = "Numero"
    Estruc(2).Total = 0
    
    Estruc(3).Alineacion = 2
    Estruc(3).Encabezado = "Real"
    Estruc(3).Size_Fuente = 9
    Estruc(3).Ancho = CentimetrosAPuntos(2.1)
    Estruc(3).Fuente = FuenteOficial
    Estruc(3).SeSuma = True
    Estruc(3).Tipo = "Numero"
    Estruc(3).Total = 0
    
    Estruc(4).Alineacion = 2
    Estruc(4).Encabezado = "Variacion"
    Estruc(4).Size_Fuente = 9
    Estruc(4).Ancho = CentimetrosAPuntos(2.1)
    Estruc(4).Fuente = FuenteOficial
    Estruc(4).SeSuma = True
    Estruc(4).Tipo = "Numero"
    Estruc(4).Total = 0
    
    Dim TotalDiap As Integer
    Dim TotalDiapS As Integer
    Num_DiapRetorno = Crea_LaTablaJorgito(Pwr, Estruc, Rs, Variables.No_DiapDetall_Variacion, Variables.No_DiapGrafAdec, "Detalles de la variacion en el Periodo " & Variables.PeriodoaEvaluar, False, True, False, 1, "Tbl_Resumen", CentimetrosAPuntos(0.74), CentimetrosAPuntos(2.89), CentimetrosAPuntos(0.52), Variables.Color_Encabezado, Variables.Color_Rng1, Variables.Color_Rng2, "245,245,245", TotalDiap, Variables.Fecha_Consumo)
    '
    Pwr.Slides(Variables.No_DiapGrafAdec).Shapes("Btn_Det_Var").ActionSettings(ppMouseClick).Action = ppActionHyperlink
    Pwr.Slides(Variables.No_DiapGrafAdec).Shapes("Btn_Det_Var").ActionSettings(ppMouseClick).Hyperlink.SubAddress = Num_DiapRetorno
    
    Set Diapositiva = Pwr.Slides(Num_DiapRetorno)
    
    Dim Grafica As PowerPoint.Chart
    Dim gChartData As PowerPoint.ChartData
    Dim gWorkBook As excel.Workbook
    Dim gWorkSheet As New excel.Worksheet
    
    Set Grafica = Diapositiva.Shapes("Graf_Cumplimiento").Chart
    Set gChartData = Grafica.ChartData
    Set gWorkBook = gChartData.Workbook
    
    gWorkBook.Worksheets(1).Range("G4").Value = Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text
    gWorkBook.Worksheets(1).Range("G5").Value = Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 3).Shape.TextFrame.TextRange.Text
    Grafica.Refresh
    
    Set Grafica = Nothing
    Set gChartData = Nothing
    Set gWorkBook = Nothing
    Set gWorkSheet = Nothing
    
    Set Grafica = Diapositiva.Shapes("Graf_Variacion").Chart
    Set gChartData = Grafica.ChartData
    Set gWorkBook = gChartData.Workbook
    
    gWorkBook.Worksheets(1).Range("C2").Value = Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text
    gWorkBook.Worksheets(1).Range("C4").Value = Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 3).Shape.TextFrame.TextRange.Text
    Grafica.Refresh
    
    Set Grafica = Nothing
    Set gChartData = Nothing
    Set gWorkBook = Nothing
    Set gWorkSheet = Nothing
    
    If Variables.Soportes Then
        Dim Tbl As PowerPoint.Shape
        Set Tbl = Diapositiva.Shapes("Tbl_Resumen")
        Set Qdf = Db.QueryDefs("Stp_2026_Soporte_VariacionenelPeriodo")
        Qdf.Parameters("Id_Adecuado") = Id_AdecuadoOficial
        Qdf.Parameters("Fecha") = Variables.Periodo_Fin
        
        
        Dim Estruc2(6) As Proceso.Tipo_Tabla
        Estruc2(1).Alineacion = 2
        Estruc2(1).Encabezado = "Pedido"
        Estruc2(1).Size_Fuente = 9
        Estruc2(1).Ancho = CentimetrosAPuntos(2.5)
        Estruc2(1).Fuente = FuenteOficial
        Estruc2(1).SeSuma = False
        Estruc2(1).Tipo = "Numero"
        Estruc2(1).Total = 0
            
        Estruc2(2).Alineacion = 4
        Estruc2(2).Encabezado = "Descripcion de la Actividad"
        Estruc2(2).Size_Fuente = 9
        Estruc2(2).Ancho = CentimetrosAPuntos(18.3)
        Estruc2(2).Fuente = FuenteOficial
        Estruc2(2).SeSuma = False
        Estruc2(2).Tipo = "Numero"
        Estruc2(2).Total = 0
        
        Estruc2(3).Alineacion = 2
        Estruc2(3).Encabezado = "Supervisor Tecnico"
        Estruc2(3).Size_Fuente = 9
        Estruc2(3).Ancho = CentimetrosAPuntos(3.9)
        Estruc2(3).Fuente = FuenteOficial
        Estruc2(3).SeSuma = False
        Estruc2(3).Tipo = "Numero"
        Estruc2(3).Total = 0
    
        Estruc2(4).Alineacion = 2
        Estruc2(4).Encabezado = "Programado"
        Estruc2(4).Size_Fuente = 9
        Estruc2(4).Ancho = CentimetrosAPuntos(2.7)
        Estruc2(4).Fuente = FuenteOficial
        Estruc2(4).SeSuma = True
        Estruc2(4).Tipo = "Numero"
        Estruc2(4).Total = 0
    
        Estruc2(5).Alineacion = 2
        Estruc2(5).Encabezado = "Facturado"
        Estruc2(5).Size_Fuente = 9
        Estruc2(5).Ancho = CentimetrosAPuntos(2.7)
        Estruc2(5).Fuente = FuenteOficial
        Estruc2(5).SeSuma = True
        Estruc2(5).Tipo = "Numero"
        Estruc2(5).Total = 0
        
        Estruc2(6).Alineacion = 2
        Estruc2(6).Encabezado = "Variacion"
        Estruc2(6).Size_Fuente = 9
        Estruc2(6).Ancho = CentimetrosAPuntos(2.7)
        Estruc2(6).Fuente = FuenteOficial
        Estruc2(6).SeSuma = True
        Estruc2(6).Tipo = "Numero"
        Estruc2(6).Total = 0
        
        Dim NumDiap2 As Integer
        Dim i As Integer
        For i = 2 To Tbl.Table.Rows.Count - 1
            Qdf.Parameters("Acreedor") = Tbl.Table.Cell(i, 1).Shape.TextFrame.TextRange.Text
            Set Rs = Qdf.OpenRecordset
            NumDiap2 = Crea_LaTablaJorgito(Pwr, Estruc2, Rs, Variables.No_DiapResumenAcreedores, Num_DiapRetorno, "Variacion en el periodo " & Variables.PeriodoaEvaluar & "para " & Tbl.Table.Cell(i, 1).Shape.TextFrame.TextRange.Text, False, True, False, 1, "Tbl_Resumen", CentimetrosAPuntos(0.74), CentimetrosAPuntos(2.85), CentimetrosAPuntos(0.52), Variables.Color_Encabezado, Variables.Color_Rng1, Variables.Color_Rng2, "245,245,245", TotalDiapS, Variables.Fecha_Consumo)
            
            Tbl.Table.Cell(i, 1).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick).Action = ppActionHyperlink
            Tbl.Table.Cell(i, 1).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick).Hyperlink.SubAddress = NumDiap2
        Next
    End If
End Sub


Sub Pedidos_Calendarizadosv2(Pwr As PowerPoint.Presentation)
On Error GoTo Error_Pedidos_Calendarizados
    Call PrepararInformacion
    'Pwr.Application.ScreenUpdating = False
    
        
    Dim Grafica As PowerPoint.Chart
    Dim gChartData As PowerPoint.ChartData
    Dim gWorkBook As excel.Workbook
    Dim gWorkSheet As New excel.Worksheet
    Dim Diapositiva As PowerPoint.Slide
    Dim NumDiapRetorno As Integer
    
    Dim Qdf As DAO.QueryDef
    Dim Rs As DAO.Recordset
    
    Dim sql As String
    
    Dim TxtComent As PowerPoint.Shape
    Dim Flechita As PowerPoint.Shape
    Dim Ancho_Barrita As Double
    Dim Ancho_Coment As Double
    Dim Ki As Integer
   Dim i As Integer
   Dim I2 As Integer
    Dim RsCia As DAO.Recordset
    
    Set RsCia = Db.OpenRecordset("Stp_AcreedoresPedidosCalendar")
    
    Dim Diap As PowerPoint.Slide
    Set Diap = Pwr.Slides(Variables.No_DiapGraficaAcreedores)
    
    Diap.SlideShowTransition.AdvanceOnClick = msoFalse
    Diap.SlideShowTransition.Hidden = msoTrue
    
    Diap.Shapes("Txt_Fechita").TextFrame.TextRange.Text = "Consumos al " & Variables.Fecha_Consumo
    Dim Tabla As PowerPoint.Shape
    Set Tabla = Diap.Shapes.AddTable(1, 2, CentimetrosAPuntos(1.36), CentimetrosAPuntos(2.85), CentimetrosAPuntos(6.7), CentimetrosAPuntos(0.2))
    Tabla.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Acreedores"
    Tabla.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Montos"
    
    Set Qdf = Db.QueryDefs("Stp_2026_ActividadesCalen_CIAS")
    
    Dim Renglon As Integer
    Dim Totalito As Double
    Dim rngEtiquetas As Range
    Dim Link As PowerPoint.ActionSetting
    Dim Botoncito As PowerPoint.Shape
    Dim Tot_ As Double
    Tot_ = 0
    
    Dim Qdf_TblRes As DAO.QueryDef
    Dim Rs_TblRed As DAO.Recordset
    
    Set Qdf_TblRes = Db.QueryDefs("Stp_Consulta_acreedorParaResumen")
    Qdf_TblRes.Parameters("Consecutivo") = maxConcec
    
    Dim Inicia_Tabla As Integer
    Inicia_Tabla = 3
    
    Dim RequiereTblResumen As Boolean
    Dim Total_TblRes As Double
    
    While Not RsCia.EOF
        RequiereTblResumen = False
        Total_TblRes = 0
        
        Qdf.Parameters("Acreedores") = RsCia("CIA_ABREV")
        Set Rs = Qdf.OpenRecordset
        
        Qdf_TblRes.Parameters("Acreedor") = RsCia("CIA_ABREV")
        Set Rs_TblRed = Qdf_TblRes.OpenRecordset
        
        
        Tabla.Table.Rows.Add
        Tabla.Table.Cell(Tabla.Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = RsCia("CIA_ABREV")
        Tabla.Table.Cell(Tabla.Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(RsCia("Total"), "#,##0.00")
        Tot_ = Tot_ + RsCia("Total")
        Pwr.Slides(Variables.No_DiapGraficaCalendarizado).Duplicate
        
        Set Diapositiva = Pwr.Slides(Variables.No_DiapGraficaCalendarizado)
        Diapositiva.MoveTo (Pwr.Slides.Count - 1)
        
        Diapositiva.SlideShowTransition.AdvanceOnClick = msoFalse
        Diapositiva.SlideShowTransition.Hidden = msoTrue
        
        Diapositiva.Shapes("Txt_Fechita").TextFrame.TextRange.Text = "Consumos al " & Variables.Fecha_Consumo
        Set Link = Diapositiva.Shapes("Btn_Inicio").ActionSettings(ppMouseClick)
        
        Link.Action = ppActionHyperlink
        Link.Hyperlink.SubAddress = Diap.SlideNumber
        
'        Diap.Shapes("Btn_Acreedor").Duplicate
'        Set Botoncito = Diap.Shapes("Btn_Acreedor")
'        Botoncito.Name = "Btn_" & RsCia("CIA_ABREV")
'        Botoncito.ActionSettings(ppMouseClick).Action = ppActionHyperlink
'        Botoncito.ActionSettings(ppMouseClick).Hyperlink.SubAddress = Diapositiva.SlideNumber
'        Botoncito.Left = Tabla.Left
'        Botoncito.Top = Tabla.Top
        
        Tabla.Table.Cell(Tabla.Table.Rows.Count, 1).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick).Action = ppActionHyperlink
        Tabla.Table.Cell(Tabla.Table.Rows.Count, 1).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick).Hyperlink.SubAddress = Diapositiva.SlideNumber
        
        While Not Rs_TblRed.EOF
            Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Add
            Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = Rs_TblRed("Estatus_Actividad")
            Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(Rs_TblRed("Total"), "#,##0.00")
            Total_TblRes = Total_TblRes + Rs_TblRed("Total")
            RequiereTblResumen = True
            Rs_TblRed.MoveNext
        Wend
        
        If Not RequiereTblResumen Then
            Diapositiva.Shapes("Tbl_Resumen").Delete
            
        Else
            Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Add
            Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = "Total"
            Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(Total_TblRes, "#,##0.00")
            Diapositiva.Shapes("Tbl_Resumen").Table.LastRow = True
        End If
        
        Set Grafica = Diapositiva.Shapes("Grafica").Chart
        Set gChartData = Grafica.ChartData
        Set gWorkBook = gChartData.Workbook
        With gWorkBook.Application
            .EnableEvents = False
            .ScreenUpdating = False
'            .Calculation = -4135
        End With
        Renglon = Inicia_Tabla
        Totalito = 0
        While Not Rs.EOF
            Renglon = Renglon + 1
       
            gWorkBook.Worksheets(1).Range("A" & Renglon).Value = Rs("Actividad")
            gWorkBook.Worksheets(1).Range("B" & Renglon).Value = Rs("FechaIni")
            gWorkBook.Worksheets(1).Range("D" & Renglon).Value = Rs("Plazo")
            gWorkBook.Worksheets(1).Range("G" & Renglon).Value = Rs("Monto")
            gWorkBook.Worksheets(1).Range("K" & Renglon).Value = Rs("Avance")
            
            gWorkBook.Worksheets(1).Range("B" & Renglon).NumberFormat = "dd/mm/aaaa"
            gWorkBook.Worksheets(1).Range("G" & Renglon).NumberFormat = "_-$* #,##0_-;-$* #,##0_-;_-$* "" - ""??_-;_-@_-"
            gWorkBook.Worksheets(1).Range("H" & Renglon).NumberFormat = "dd/mm/aaaa"
            
            gWorkBook.Worksheets(1).Range("C" & Renglon).Formula = "=((1+SQRT(5))/2)"
            gWorkBook.Worksheets(1).Range("E" & Renglon).Formula = "=((1+SQRT(5))/2)"
            gWorkBook.Worksheets(1).Range("F" & Renglon).Formula = "=(((1+SQRT(5))/2)^10)/SQRT(4)"
            gWorkBook.Worksheets(1).Range("H" & Renglon).Formula = "=CONCATENATE(TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm""),""   "",TEXT(G" & Renglon & ",""#,##0.00""),"" MM$"")" '"=TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm"") & "" "" TEXT(G" & Renglon & ",""#,##0.00"")"
            '"=CONCATENATE(TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm""),""   "",TEXT(G" & Renglon & ",""#,##0.00""),"" MM$"")"
            gWorkBook.Worksheets(1).Range("I" & Renglon).Formula = "=D" & Renglon & "&"" d?as"""
            'gWorkBook.Worksheets(1).Range("J" & Renglon).Formula = "=TEXT(G" & Renglon & ",""$ #,##0.00"")"
            gWorkBook.Worksheets(1).Range("J" & Renglon).Value = ""
            gWorkBook.Worksheets(1).Range("L" & Renglon).Formula = "=CONCATENATE(TEXT(B" & Renglon & ",""d-mmm""),""   "",I" & Renglon & ")"
            gWorkBook.Worksheets(1).Range("M" & Renglon).Formula = "=if(k" & Renglon & "<>"""",d" & Renglon & "*k" & Renglon & "/100,0)" ' "=d" & Renglon & "*k" & Renglon & "/100"
            gWorkBook.Worksheets(1).Range("N" & Renglon).Formula = "=d" & Renglon & "-m" & Renglon
            gWorkBook.Worksheets(1).Range("O" & Renglon).Formula = "=if(k" & Renglon & "<>"""",concatenate(k" & Renglon & ","" %""),"""")"
            gWorkBook.Worksheets(1).Range("P" & Renglon).Value = RsCia("CIA_ABREV")
            gWorkBook.Worksheets(1).Range("T" & Renglon).Value = Rs("Nombre_Corto")
            gWorkBook.Worksheets(1).Range("U" & Renglon).Value = Rs("Comentarios")
            
            gWorkBook.Worksheets(1).Range("D2").Formula = "=sum(G" & Inicia_Tabla + 1 & ":G" & Renglon & ")"
            
            Totalito = Totalito + Rs("Monto")
            'Grafica.SetSourceData ("=" & gWorkBook.Worksheets(1).Name & "!$b$1:$C$" & i)
            
            Rs.MoveNext
        Wend
        
      
        If Renglon < 45 Then
            gWorkBook.Worksheets(1).Range("A" & Renglon + 1 & ":U45").ClearContents
        End If
        
                
        Grafica.SeriesCollection(1).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$B$" & Inicia_Tabla + 1 & ":$B$" & Renglon
        Grafica.SeriesCollection(1).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        
        Grafica.SeriesCollection(2).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$M$" & Inicia_Tabla + 1 & ":$M$" & Renglon
        Grafica.SeriesCollection(2).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        
        Grafica.SeriesCollection(3).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$N$" & Inicia_Tabla + 1 & ":$N$" & Renglon
        Grafica.SeriesCollection(3).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        
'        Grafica.SeriesCollection(3).ApplyDataLabels
'        Grafica.SeriesCollection(3).DataLabels.ShowRange = True
'         Grafica.SeriesCollection(3).HasDataLabels = True
         
         'xlChartFieldUser
         
'         Grafica.SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$L$3:$L$" & Renglon, 0
'        Grafica.SeriesCollection(1).DataLabels.ShowValue = False
'        Grafica.SeriesCollection(1).DataLabels.ShowCategoryName = False
'        Grafica.SeriesCollection(1).DataLabels.ShowPercentage = False
         
        Grafica.SeriesCollection(2).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$O$" & Inicia_Tabla + 1 & ":$O$" & Renglon, 0
        Grafica.SeriesCollection(2).DataLabels.ShowValue = False
        Grafica.SeriesCollection(2).DataLabels.ShowCategoryName = False
        Grafica.SeriesCollection(2).DataLabels.ShowPercentage = False
        
'        Grafica.SeriesCollection(3).DataLabels.ShowValue = False
'        Grafica.SeriesCollection(3).DataLabels.ShowCategoryName = False
'        Grafica.SeriesCollection(3).DataLabels.ShowPercentage = False
        
        Grafica.SeriesCollection(4).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$E$" & Inicia_Tabla + 1 & ":$E$" & Renglon
        Grafica.SeriesCollection(4).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        Grafica.SeriesCollection(4).ApplyDataLabels
        Grafica.SeriesCollection(4).DataLabels.ShowRange = True
         Grafica.SeriesCollection(4).HasDataLabels = True
         
         'xlChartFieldUser
        Grafica.SeriesCollection(4).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$H$" & Inicia_Tabla + 1 & ":$H$" & Renglon, 0
        Grafica.SeriesCollection(4).DataLabels.ShowValue = False
        Grafica.SeriesCollection(4).DataLabels.ShowCategoryName = False
        Grafica.SeriesCollection(4).DataLabels.ShowPercentage = False
'        Grafica.SeriesCollection(4).DataLabels.
        
        Grafica.SeriesCollection(5).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$F$" & Inicia_Tabla + 1 & ":$F$" & Renglon
        Grafica.SeriesCollection(5).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
          Grafica.SeriesCollection(5).ApplyDataLabels
        Grafica.SeriesCollection(5).DataLabels.ShowRange = True
         Grafica.SeriesCollection(5).HasDataLabels = True
         
'         xlChartFieldUser
        Grafica.SeriesCollection(5).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$J$" & Inicia_Tabla + 1 & ":$J$" & Renglon, 0
        Grafica.SeriesCollection(5).DataLabels.ShowValue = False
        Grafica.SeriesCollection(5).DataLabels.ShowCategoryName = False
        Grafica.SeriesCollection(5).DataLabels.ShowPercentage = False

        Grafica.Refresh
        
        If Renglon - Inicia_Tabla = 1 Then
            Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
            Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(3.26)
        End If
        
        If Renglon - Inicia_Tabla = 2 Then
            Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
            Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(4.41)
        End If
        
        
          If Renglon - Inicia_Tabla = 3 Then
            Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
            Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(6.05)
        End If
        
        
        
        Diapositiva.Shapes("Txt_Total").TextFrame.TextRange.Text = "Total " & Format(Totalito, "#,##0.00") & " Millones de Pesos"
        Diapositiva.Shapes("Título 1").TextFrame.TextRange.Text = "Actividades Calendarizadas de " & RsCia("CIA_ABREV")
      
        
    
        For Ki = 1 To Renglon - Inicia_Tabla
            'MsgBox Grafica.SeriesCollection(4).Points(Ki).Top & " " & Diapositiva.Shapes(Grafica.Name).Top
            If gWorkBook.Worksheets(1).Range("U" & Ki + Inicia_Tabla).Value <> "" Then
                Diapositiva.Shapes("Txt_Comentario").Duplicate
                Diapositiva.Shapes("Flechita").Duplicate
                Set TxtComent = Diapositiva.Shapes("Txt_Comentario")
                Set Flechita = Diapositiva.Shapes("Flechita")
                TxtComent.Name = "Txt_Comentario_" & Ki
                Flechita.Name = "Flechita_" & Ki
                Ancho_Barrita = Grafica.SeriesCollection(4).Points(Ki).Height
                Ancho_Coment = TxtComent.Height
                
                TxtComent.TextFrame.TextRange.Text = gWorkBook.Worksheets(1).Range("U" & Ki + Inicia_Tabla).Value
                TxtComent.Top = Grafica.SeriesCollection(4).Points(Ki).Top + Diapositiva.Shapes(Grafica.Name).Top + ((Ancho_Barrita - Ancho_Coment) / 2)
                TxtComent.Left = Diapositiva.Shapes(Grafica.Name).Left + Diapositiva.Shapes(Grafica.Name).Width + 30
                Flechita.Top = Grafica.SeriesCollection(4).Points(Ki).Top + Diapositiva.Shapes(Grafica.Name).Top + (Ancho_Barrita / 2)
                Flechita.Width = 28
                Flechita.Left = Grafica.SeriesCollection(4).DataLabels(Ki).Width + Grafica.SeriesCollection(4).DataLabels(Ki).Left + 15
                Flechita.Width = TxtComent.Left - (Grafica.SeriesCollection(4).DataLabels(Ki).Width + Grafica.SeriesCollection(4).DataLabels(Ki).Left + 15)
            End If
        Next
        
        With gWorkBook.Application
'            .Calculation = -4105
            .EnableEvents = True
            .ScreenUpdating = True
        End With
        gWorkBook.Close
'        gWorkBook.Application.Quit
        
        Diapositiva.Shapes("Txt_Comentario").Delete
        Diapositiva.Shapes("Flechita").Delete
        
        For i = 1 To Grafica.SeriesCollection(1).DataLabels.Count
            If Grafica.SeriesCollection(1).DataLabels(i).Left + Grafica.SeriesCollection(1).DataLabels(i).Width > Grafica.SeriesCollection(2).DataLabels(i).Left Then
'                MsgBox "Se trasponen las cajitas"
                Grafica.SeriesCollection(1).DataLabels(i).Left = Grafica.SeriesCollection(1).DataLabels(i).Left - ((Grafica.SeriesCollection(1).DataLabels(i).Left + Grafica.SeriesCollection(1).DataLabels(i).Width) - Grafica.SeriesCollection(2).DataLabels(i).Left)
            End If
        Next
        
        
        Set Grafica = Nothing
        Set gChartData = Nothing
        Set gWorkBook = Nothing

        RsCia.MoveNext
    Wend
    
    Tabla.Table.Rows.Add
    Tabla.Table.Cell(Tabla.Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = "Total"
    Tabla.Table.Cell(Tabla.Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(Tot_, "#,##0.00")
    Tabla.Table.LastRow = True
    
    Dim Rng_Tbl As Integer
    Dim Col_Tbl As Integer
    
    For Rng_Tbl = 1 To Tabla.Table.Rows.Count
        For Col_Tbl = 1 To Tabla.Table.Columns.Count
            Tabla.Table.Cell(Rng_Tbl, Col_Tbl).Shape.TextFrame.TextRange.Font.Size = 9
            Tabla.Table.Cell(Rng_Tbl, Col_Tbl).Shape.TextFrame.MarginBottom = 0
            Tabla.Table.Cell(Rng_Tbl, Col_Tbl).Shape.TextFrame.MarginTop = 0
'        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginLeft = 2
'        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginRight = 2
'        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginBottom = 2
'        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginTop = 2
        Next
    Next
    Tabla.Height = 0
    Tabla.Table.Columns(1).Width = CentimetrosAPuntos(5.8)
    Tabla.Table.Columns(2).Width = CentimetrosAPuntos(1.9)
    
    
    Set Grafica = Diap.Shapes("Grafico").Chart
    Set gChartData = Grafica.ChartData
    Set gWorkBook = gChartData.Workbook
    With gWorkBook.Application
        .EnableEvents = False
        .ScreenUpdating = False
'        .Calculation = -4135
    End With
    Renglon = Inicia_Tabla
    Set Rs = Db.OpenRecordset("Stp_2026_ActivPorMes")
    While Not Rs.EOF
        gWorkBook.Worksheets(1).Range("B" & Rs("Pos") + 1).Value = Rs("Total")
        gWorkBook.Worksheets(1).Range("B" & Rs("Pos") + 1).NumberFormat = "#,##0.00"
        Rs.MoveNext
    Wend
    
    
    Grafica.SetSourceData ("=" & gWorkBook.Worksheets(1).Name & "!$A$1:$B$13")
    Grafica.Refresh
    
    With gWorkBook.Application
'        .Calculation = -4105
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    gWorkBook.Close
'    gWorkBook.Application.Quit
    Set Grafica = Nothing
    Set gChartData = Nothing
    Set gWorkBook = Nothing
    
    Set RsCia = Db.OpenRecordset("Stp_2026_ActivPorSupervisores")
    
    Dim Tbl2 As PowerPoint.Shape
    Set Tbl2 = Diap.Shapes.AddTable(1, 2, CentimetrosAPuntos(26.08), CentimetrosAPuntos(2.85), CentimetrosAPuntos(6.7), CentimetrosAPuntos(0.2))
    
    Tbl2.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Supervisores T?cnicos"
    Tbl2.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Montos"
    
     Set Qdf = Db.QueryDefs("Stp_2026_ActividadesCalen_RESPONSABLES")
    While Not RsCia.EOF
        Tbl2.Table.Rows.Add
        Tbl2.Table.Cell(Tbl2.Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = RsCia("Nombre_Corto")
        Tbl2.Table.Cell(Tbl2.Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(RsCia("Total"), "#,##0.00")
        
        Qdf.Parameters("Responsable") = RsCia("Nombre_Corto")
        Set Rs = Qdf.OpenRecordset
        
        Pwr.Slides(Variables.No_DiapGraficaCalendarizado).Duplicate
        
        Set Diapositiva = Pwr.Slides(Variables.No_DiapGraficaCalendarizado)
        Diapositiva.MoveTo (Pwr.Slides.Count - 1)
        Diapositiva.Shapes("Txt_Fechita").TextFrame.TextRange.Text = "Consumos al " & Variables.Fecha_Consumo
        
        Diapositiva.SlideShowTransition.AdvanceOnClick = msoFalse
        Diapositiva.SlideShowTransition.Hidden = msoTrue
        
        Set Link = Diapositiva.Shapes("Btn_Inicio").ActionSettings(ppMouseClick)
        
        Link.Action = ppActionHyperlink
        Link.Hyperlink.SubAddress = Diap.SlideNumber
        
        Tbl2.Table.Cell(Tbl2.Table.Rows.Count, 1).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick).Action = ppActionHyperlink
        Tbl2.Table.Cell(Tbl2.Table.Rows.Count, 1).Shape.TextFrame.TextRange.ActionSettings(ppMouseClick).Hyperlink.SubAddress = Diapositiva.SlideNumber
        
'        Diap.Shapes("Btn_Acreedor").Duplicate
'        Set Botoncito = Diap.Shapes("Btn_Acreedor")
'        Botoncito.Name = "Btn_" & RsCia("Nombre_Corto")
'        Botoncito.ActionSettings(ppMouseClick).Action = ppActionHyperlink
'        Botoncito.ActionSettings(ppMouseClick).Hyperlink.SubAddress = Diapositiva.SlideNumber
'        Botoncito.Left = Tbl2.Left
'        Botoncito.Top = Tbl2.Top
        
        Set Grafica = Diapositiva.Shapes("Grafica").Chart
        Set gChartData = Grafica.ChartData
        Set gWorkBook = gChartData.Workbook
        With gWorkBook.Application
            .EnableEvents = False
            .ScreenUpdating = False
'            .Calculation = -4135
        End With
        Renglon = Inicia_Tabla
        Totalito = 0
        While Not Rs.EOF
            Renglon = Renglon + 1
                        
            gWorkBook.Worksheets(1).Range("A" & Renglon).Value = Rs("Actividad")
            gWorkBook.Worksheets(1).Range("B" & Renglon).Value = Rs("FechaIni")
            gWorkBook.Worksheets(1).Range("D" & Renglon).Value = Rs("Plazo")
            gWorkBook.Worksheets(1).Range("G" & Renglon).Value = Rs("Monto")
            gWorkBook.Worksheets(1).Range("K" & Renglon).Value = Rs("Avance")
            
            gWorkBook.Worksheets(1).Range("B" & Renglon).NumberFormat = "dd/mm/aaaa"
            gWorkBook.Worksheets(1).Range("G" & Renglon).NumberFormat = "_-$* #,##0_-;-$* #,##0_-;_-$* "" - ""??_-;_-@_-"
            gWorkBook.Worksheets(1).Range("H" & Renglon).NumberFormat = "dd/mm/aaaa"
            
            gWorkBook.Worksheets(1).Range("C" & Renglon).Formula = "=((1+SQRT(5))/2)"
            gWorkBook.Worksheets(1).Range("E" & Renglon).Formula = "=((1+SQRT(5))/2)"
            gWorkBook.Worksheets(1).Range("F" & Renglon).Formula = "=(((1+SQRT(5))/2)^10)/SQRT(4)"
            gWorkBook.Worksheets(1).Range("H" & Renglon).Formula = "=CONCATENATE(TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm""),""   "",TEXT(G" & Renglon & ",""#,##0.00""),"" MM$"")" '"=TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm"") & "" "" TEXT(G" & Renglon & ",""#,##0.00"")"'"=TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm"")"
            gWorkBook.Worksheets(1).Range("I" & Renglon).Formula = "=D" & Renglon & "&"" d?as"""
'            gWorkBook.Worksheets(1).Range("J" & Renglon).Formula = "=TEXT(G" & Renglon & ",""$ #,##0.00"")"
            gWorkBook.Worksheets(1).Range("J" & Renglon).Value = ""
             gWorkBook.Worksheets(1).Range("L" & Renglon).Formula = "=CONCATENATE(TEXT(B" & Renglon & ",""d-mmm""),""   "",I" & Renglon & ")"
            gWorkBook.Worksheets(1).Range("M" & Renglon).Formula = "=if(k" & Renglon & "<>"""",d" & Renglon & "*k" & Renglon & "/100,0)" '  "=d" & Renglon & "*k" & Renglon & "/100"
            gWorkBook.Worksheets(1).Range("N" & Renglon).Formula = "=d" & Renglon & "-m" & Renglon
            gWorkBook.Worksheets(1).Range("O" & Renglon).Formula = "=if(k" & Renglon & "<>"""",concatenate(k" & Renglon & ","" %""),"""")"
            
            gWorkBook.Worksheets(1).Range("P" & Renglon).Value = Rs("CIA_ABREV")
            gWorkBook.Worksheets(1).Range("T" & Renglon).Value = RsCia("Nombre_Corto")
            gWorkBook.Worksheets(1).Range("U" & Renglon).Value = Rs("Comentarios")
            
            gWorkBook.Worksheets(1).Range("D2").Formula = "=sum(G" & Inicia_Tabla + 1 & ":G" & Renglon & ")"
            
            Totalito = Totalito + Rs("Monto")
            'Grafica.SetSourceData ("=" & gWorkBook.Worksheets(1).Name & "!$b$1:$C$" & i)
            
            Rs.MoveNext
        Wend
        
        If Renglon < 45 Then
            gWorkBook.Worksheets(1).Range("A" & Renglon + 1 & ":U45").ClearContents
        End If
        
        
        Grafica.SeriesCollection(1).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$B$" & Inicia_Tabla + 1 & ":$B$" & Renglon
        Grafica.SeriesCollection(1).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        
        Grafica.SeriesCollection(2).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$M$" & Inicia_Tabla + 1 & ":$M$" & Renglon
        Grafica.SeriesCollection(2).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        
        Grafica.SeriesCollection(3).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$N$" & Inicia_Tabla + 1 & ":$N$" & Renglon
        Grafica.SeriesCollection(3).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        
'        Grafica.SeriesCollection(3).ApplyDataLabels
'        Grafica.SeriesCollection(3).DataLabels.ShowRange = True
'         Grafica.SeriesCollection(3).HasDataLabels = True
         
         'xlChartFieldUser
'
'        Grafica.SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$L$3:$L$" & Renglon, 0
'        Grafica.SeriesCollection(1).DataLabels.ShowValue = False
'        Grafica.SeriesCollection(1).DataLabels.ShowCategoryName = False
'        Grafica.SeriesCollection(1).DataLabels.ShowPercentage = False
        
        Grafica.SeriesCollection(2).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$O$" & Inicia_Tabla + 1 & ":$O$" & Renglon, 0
        Grafica.SeriesCollection(2).DataLabels.ShowValue = False
        Grafica.SeriesCollection(2).DataLabels.ShowCategoryName = False
        Grafica.SeriesCollection(2).DataLabels.ShowPercentage = False
        
'        Grafica.SeriesCollection(3).DataLabels.ShowValue = False
'        Grafica.SeriesCollection(3).DataLabels.ShowCategoryName = False
'        Grafica.SeriesCollection(3).DataLabels.ShowPercentage = False

                
        Grafica.SeriesCollection(4).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$E$" & Inicia_Tabla + 1 & ":$E$" & Renglon
        Grafica.SeriesCollection(4).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
        Grafica.SeriesCollection(4).ApplyDataLabels
        Grafica.SeriesCollection(4).DataLabels.ShowRange = True
         Grafica.SeriesCollection(4).HasDataLabels = True
         
         'xlChartFieldUser
        Grafica.SeriesCollection(4).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$H$" & Inicia_Tabla + 1 & ":$H$" & Renglon, 0
        Grafica.SeriesCollection(4).DataLabels.ShowValue = False
        Grafica.SeriesCollection(4).DataLabels.ShowCategoryName = False
        Grafica.SeriesCollection(4).DataLabels.ShowPercentage = False
'        Grafica.SeriesCollection(4).DataLabels.
        
        Grafica.SeriesCollection(5).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$F$" & Inicia_Tabla + 1 & ":$F$" & Renglon
        Grafica.SeriesCollection(5).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
          Grafica.SeriesCollection(5).ApplyDataLabels
        Grafica.SeriesCollection(5).DataLabels.ShowRange = True
         Grafica.SeriesCollection(5).HasDataLabels = True
         
         'xlChartFieldUser
        Grafica.SeriesCollection(5).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$J$" & Inicia_Tabla + 1 & ":$J$" & Renglon, 0
        Grafica.SeriesCollection(5).DataLabels.ShowValue = False
        Grafica.SeriesCollection(5).DataLabels.ShowCategoryName = False
        Grafica.SeriesCollection(5).DataLabels.ShowPercentage = False

        Grafica.Refresh
        
        If Renglon - Inicia_Tabla = 1 Then
          Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
          Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(3.26)
        End If
        
        If Renglon - Inicia_Tabla = 2 Then
          Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
          Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(4.41)
        End If
        
        If Renglon - Inicia_Tabla = 3 Then
          Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
          Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(6.05)
        End If
        
        Diapositiva.Shapes("Txt_Total").TextFrame.TextRange.Text = "Total " & Format(Totalito, "#,##0.00") & " Millones de Pesos"
        Diapositiva.Shapes("Título 1").TextFrame.TextRange.Text = "Actividades Calendarizadas de " & RsCia("Nombre_Corto")
        
        For Ki = 1 To Renglon - Inicia_Tabla
            'MsgBox Grafica.SeriesCollection(4).Points(Ki).Top & " " & Diapositiva.Shapes(Grafica.Name).Top
            If gWorkBook.Worksheets(1).Range("U" & Ki + Inicia_Tabla).Value <> "" Then
                Diapositiva.Shapes("Txt_Comentario").Duplicate
                Diapositiva.Shapes("Flechita").Duplicate
                Set TxtComent = Diapositiva.Shapes("Txt_Comentario")
                Set Flechita = Diapositiva.Shapes("Flechita")
                TxtComent.Name = "Txt_Comentario_" & Ki
                Flechita.Name = "Flechita_" & Ki
                Ancho_Barrita = Grafica.SeriesCollection(4).Points(Ki).Height
                Ancho_Coment = TxtComent.Height
                
                TxtComent.TextFrame.TextRange.Text = gWorkBook.Worksheets(1).Range("U" & Ki + Inicia_Tabla).Value
                TxtComent.Top = Grafica.SeriesCollection(4).Points(Ki).Top + Diapositiva.Shapes(Grafica.Name).Top + ((Ancho_Barrita - Ancho_Coment) / 2)
                TxtComent.Left = Diapositiva.Shapes(Grafica.Name).Left + Diapositiva.Shapes(Grafica.Name).Width + 30
                Flechita.Top = Grafica.SeriesCollection(4).Points(Ki).Top + Diapositiva.Shapes(Grafica.Name).Top + (Ancho_Barrita / 2)
                Flechita.Width = 28
                Flechita.Left = Grafica.SeriesCollection(4).DataLabels(Ki).Width + Grafica.SeriesCollection(4).DataLabels(Ki).Left + 15
                Flechita.Width = TxtComent.Left - (Grafica.SeriesCollection(4).DataLabels(Ki).Width + Grafica.SeriesCollection(4).DataLabels(Ki).Left + 15)
            End If
        Next
        
        
        With gWorkBook.Application
'            .Calculation = -4105
            .EnableEvents = True
            .ScreenUpdating = True
        End With
        gWorkBook.Close
'        gWorkBook.Application.Quit
        
        Diapositiva.Shapes("Txt_Comentario").Delete
        Diapositiva.Shapes("Flechita").Delete
        
        
        For i = 1 To Grafica.SeriesCollection(1).DataLabels.Count
            If Grafica.SeriesCollection(1).DataLabels(i).Left + Grafica.SeriesCollection(1).DataLabels(i).Width > Grafica.SeriesCollection(2).DataLabels(i).Left Then
'                MsgBox "Se trasponen las cajitas"
                Grafica.SeriesCollection(1).DataLabels(i).Left = Grafica.SeriesCollection(1).DataLabels(i).Left - ((Grafica.SeriesCollection(1).DataLabels(i).Left + Grafica.SeriesCollection(1).DataLabels(i).Width) - Grafica.SeriesCollection(2).DataLabels(i).Left)
            End If
        Next
        
        
        Set Grafica = Nothing
        Set gChartData = Nothing
        Set gWorkBook = Nothing
        
        RsCia.MoveNext
    Wend
        
    Tbl2.Table.Rows.Add
    Tbl2.Table.Cell(Tbl2.Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = "Total"
    Tbl2.Table.Cell(Tbl2.Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(Tot_, "#,##0.00")
    Tbl2.Table.LastRow = True
    
    
    For Rng_Tbl = 1 To Tbl2.Table.Rows.Count
        For Col_Tbl = 1 To Tbl2.Table.Columns.Count
            Tbl2.Table.Cell(Rng_Tbl, Col_Tbl).Shape.TextFrame.TextRange.Font.Size = 9
            Tbl2.Table.Cell(Rng_Tbl, Col_Tbl).Shape.TextFrame.MarginBottom = 0
            Tbl2.Table.Cell(Rng_Tbl, Col_Tbl).Shape.TextFrame.MarginTop = 0
        '        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginLeft = 2
        '        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginRight = 2
        '        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginBottom = 2
        '        Tabla.Table.Cell(1, j).Shape.TextFrame.MarginTop = 2
        Next
    Next
    Tbl2.Height = 0
    
    Tbl2.Table.Columns(1).Width = CentimetrosAPuntos(5.1)
    Tbl2.Table.Columns(2).Width = CentimetrosAPuntos(1.9)
    
    Set Grafica = Diap.Shapes("Grafico").Chart
     Set Qdf = Db.QueryDefs("Stp_2026_ActivCalendarizadasCIA_Responsable")
    ' Dim i As Integer
     Dim Graf As PowerPoint.Chart
     
    Set Qdf_TblRes = Db.QueryDefs("Stp_TablaResumenMensual")
    Qdf_TblRes.Parameters("Consecutivo") = maxConcec
     
    For i = 1 To Grafica.SeriesCollection(1).DataLabels.Count
        
        If Grafica.SeriesCollection(1).DataLabels(i).Text <> "" Then
            RequiereTblResumen = False
            Diap.Shapes("Btn_Acreedor").Duplicate
            Set Botoncito = Diap.Shapes("Btn_Acreedor")
            
            Botoncito.Name = "Btn_" & i
            
            Botoncito.Top = Grafica.SeriesCollection(1).DataLabels(i).Top + Diap.Shapes("Grafico").Top + Botoncito.Width
            Botoncito.Left = Grafica.SeriesCollection(1).DataLabels(i).Left + Diap.Shapes("Grafico").Left

            Qdf.Parameters("No_mes") = i
            Set Rs = Qdf.OpenRecordset
            
            Qdf_TblRes.Parameters("No_Mes") = i
             Set Rs_TblRed = Qdf_TblRes.OpenRecordset
            
            
            
            
            '***************************************
            
            Pwr.Slides(Variables.No_DiapGraficaCalendarizado).Duplicate
            
           
            
            Set Diapositiva = Pwr.Slides(Variables.No_DiapGraficaCalendarizado)
            Diapositiva.MoveTo (Pwr.Slides.Count - 1)
            Diapositiva.Shapes("Txt_Fechita").TextFrame.TextRange.Text = "Consumos al " & Variables.Fecha_Consumo
            
            Diapositiva.SlideShowTransition.AdvanceOnClick = msoFalse
            Diapositiva.SlideShowTransition.Hidden = msoTrue
            
            Set Link = Diapositiva.Shapes("Btn_Inicio").ActionSettings(ppMouseClick)
            
            Link.Action = ppActionHyperlink
            Link.Hyperlink.SubAddress = Diap.SlideNumber
            
            Botoncito.ActionSettings(ppMouseClick).Action = ppActionHyperlink
            Botoncito.ActionSettings(ppMouseClick).Hyperlink.SubAddress = Diapositiva.SlideNumber
            
            
         
            
            
            Set Graf = Diapositiva.Shapes("Grafica").Chart
            Set gChartData = Graf.ChartData
            Set gWorkBook = gChartData.Workbook
            With gWorkBook.Application
                .EnableEvents = False
                .ScreenUpdating = False
'                .Calculation = -4135
            End With
            Renglon = Inicia_Tabla
            Totalito = 0
            While Not Rs.EOF
              Renglon = Renglon + 1
                
                gWorkBook.Worksheets(1).Range("A" & Renglon).Value = Rs("ActividadCorta") & "/" & Rs("CIA_ABREV") & vbNewLine & Rs("Nombre_Corto")
                
                gWorkBook.Worksheets(1).Range("B" & Renglon).Value = Rs("FechaIni")
                gWorkBook.Worksheets(1).Range("D" & Renglon).Value = Rs("Plazo")
                gWorkBook.Worksheets(1).Range("G" & Renglon).Value = Rs("Monto")
                gWorkBook.Worksheets(1).Range("K" & Renglon).Value = Rs("Avance")
                
                gWorkBook.Worksheets(1).Range("B" & Renglon).NumberFormat = "dd/mm/aaaa"
                gWorkBook.Worksheets(1).Range("G" & Renglon).NumberFormat = "_-$* #,##0_-;-$* #,##0_-;_-$* "" - ""??_-;_-@_-"
                gWorkBook.Worksheets(1).Range("H" & Renglon).NumberFormat = "dd/mm/aaaa"
                
                gWorkBook.Worksheets(1).Range("C" & Renglon).Formula = "=((1+SQRT(5))/2)"
                gWorkBook.Worksheets(1).Range("E" & Renglon).Formula = "=((1+SQRT(5))/2)"
                gWorkBook.Worksheets(1).Range("F" & Renglon).Formula = "=(((1+SQRT(5))/2)^10)/SQRT(4)"
                gWorkBook.Worksheets(1).Range("H" & Renglon).Formula = "=CONCATENATE(TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm""),""   "",TEXT(G" & Renglon & ",""#,##0.00""),"" MM$"")" '"=TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm"") & "" "" TEXT(G" & Renglon & ",""#,##0.00"")"'"=TEXT(B" & Renglon & "+D" & Renglon & ",""d-mmm"")"
                gWorkBook.Worksheets(1).Range("I" & Renglon).Formula = "=D" & Renglon & "&"" d?as"""
                '              gWorkBook.Worksheets(1).Range("J" & Renglon).Formula = "=TEXT(G" & Renglon & ",""$ #,##0.00"")"
                gWorkBook.Worksheets(1).Range("J" & Renglon).Value = ""
                gWorkBook.Worksheets(1).Range("L" & Renglon).Formula = "=CONCATENATE(TEXT(B" & Renglon & ",""d-mmm""),""   "",I" & Renglon & ")"
                gWorkBook.Worksheets(1).Range("M" & Renglon).Formula = "=if(k" & Renglon & "<>"""",d" & Renglon & "*k" & Renglon & "/100,0)" ' "=d" & Renglon & "*k" & Renglon & "/100"
                gWorkBook.Worksheets(1).Range("N" & Renglon).Formula = "=d" & Renglon & "-m" & Renglon
                gWorkBook.Worksheets(1).Range("O" & Renglon).Formula = "=if(k" & Renglon & "<>"""",concatenate(k" & Renglon & ","" %""),"""")"
                
                gWorkBook.Worksheets(1).Range("P" & Renglon).Value = Rs("CIA_ABREV")
                gWorkBook.Worksheets(1).Range("S" & Renglon).Value = Rs("ActividadCorta")
                gWorkBook.Worksheets(1).Range("T" & Renglon).Value = Rs("Nombre_Corto")
                gWorkBook.Worksheets(1).Range("U" & Renglon).Value = Rs("Comentarios")
                
                gWorkBook.Worksheets(1).Range("D2").Formula = "=sum(G" & Inicia_Tabla + 1 & ":G" & Renglon & ")"
                
              Totalito = Totalito + Rs("Monto")
              'Grafica.SetSourceData ("=" & gWorkBook.Worksheets(1).Name & "!$b$1:$C$" & i)
              
              Rs.MoveNext
            Wend
            
            If Renglon < 45 Then
                gWorkBook.Worksheets(1).Range("A" & Renglon + 1 & ":U45").ClearContents
            End If
            
            Graf.SeriesCollection(1).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$B$" & Inicia_Tabla + 1 & ":$B$" & Renglon
            Graf.SeriesCollection(1).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
            
            Graf.SeriesCollection(2).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$M$" & Inicia_Tabla + 1 & ":$M$" & Renglon
            Graf.SeriesCollection(2).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
            
            Graf.SeriesCollection(3).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$N$" & Inicia_Tabla + 1 & ":$N$" & Renglon
            Graf.SeriesCollection(3).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
            
'            Graf.SeriesCollection(3).ApplyDataLabels
'            Graf.SeriesCollection(3).DataLabels.ShowRange = True
'            Graf.SeriesCollection(3).HasDataLabels = True
            
            'xlChartFieldUser
'            Grafica.SeriesCollection(1).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$L$3:$L$" & Renglon, 0
'            Grafica.SeriesCollection(1).DataLabels.ShowValue = False
'            Grafica.SeriesCollection(1).DataLabels.ShowCategoryName = False
'            Grafica.SeriesCollection(1).DataLabels.ShowPercentage = False
            
            On Error Resume Next
            Graf.SeriesCollection(2).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$O$" & Inicia_Tabla + 1 & ":$O$" & Renglon, 0
            Graf.SeriesCollection(2).DataLabels.ShowValue = False
            Graf.SeriesCollection(2).DataLabels.ShowCategoryName = False
            Graf.SeriesCollection(2).DataLabels.ShowPercentage = False
            
'            Grafica.SeriesCollection(3).DataLabels.ShowValue = False
'            Grafica.SeriesCollection(3).DataLabels.ShowCategoryName = False
'            Grafica.SeriesCollection(3).DataLabels.ShowPercentage = False
            
            
            Graf.SeriesCollection(4).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$E$" & Inicia_Tabla + 1 & ":$E$" & Renglon
            Graf.SeriesCollection(4).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
            Graf.SeriesCollection(4).ApplyDataLabels
            Graf.SeriesCollection(4).DataLabels.ShowRange = True
            Graf.SeriesCollection(4).HasDataLabels = True
            
            'xlChartFieldUser
            Graf.SeriesCollection(4).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$H$" & Inicia_Tabla + 1 & ":$H$" & Renglon, 0
            Graf.SeriesCollection(4).DataLabels.ShowValue = False
            Graf.SeriesCollection(4).DataLabels.ShowCategoryName = False
            Graf.SeriesCollection(4).DataLabels.ShowPercentage = False
            '        Grafica.SeriesCollection(4).DataLabels.
            
            Graf.SeriesCollection(5).Values = "='" & gWorkBook.Worksheets(1).Name & "'!$F$" & Inicia_Tabla + 1 & ":$F$" & Renglon
            Graf.SeriesCollection(5).XValues = "='" & gWorkBook.Worksheets(1).Name & "'!$A$" & Inicia_Tabla + 1 & ":$A$" & Renglon
            Graf.SeriesCollection(5).ApplyDataLabels
            Graf.SeriesCollection(5).DataLabels.ShowRange = True
            Graf.SeriesCollection(5).HasDataLabels = True
            
            'xlChartFieldUser
            Graf.SeriesCollection(5).DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & gWorkBook.Worksheets(1).Name & "'!$J$" & Inicia_Tabla + 1 & ":$J$" & Renglon, 0
            Graf.SeriesCollection(5).DataLabels.ShowValue = False
            Graf.SeriesCollection(5).DataLabels.ShowCategoryName = False
            Graf.SeriesCollection(5).DataLabels.ShowPercentage = False
            
            Graf.Refresh
            
            If Renglon - Inicia_Tabla = 1 Then
                Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
                Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(3.26)
            End If
            
            If Renglon - Inicia_Tabla = 2 Then
                Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
                Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(4.41)
            End If
            
            
              If Renglon - Inicia_Tabla = 3 Then
                Diapositiva.Shapes("Grafica").Top = CentimetrosAPuntos(3.18)
                Diapositiva.Shapes("Grafica").Height = CentimetrosAPuntos(6.05)
            End If
        
            Diapositiva.Shapes("Txt_Total").TextFrame.TextRange.Text = "Total " & Format(Totalito, "#,##0.00") & " Millones de Pesos"
            Diapositiva.Shapes("Título 1").TextFrame.TextRange.Text = "Actividades programadas para terminar en " & Nombre_Mes(i)
            
            
            Total_TblRes = 0
            While Not Rs_TblRed.EOF
                Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Add
                Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = Rs_TblRed("Estatus_Actividad")
                Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(Rs_TblRed("Total"), "#,##0.00")
                Total_TblRes = Total_TblRes + Rs_TblRed("Total")
                RequiereTblResumen = True
                Rs_TblRed.MoveNext
            Wend
            
            If Not RequiereTblResumen Then
                Diapositiva.Shapes("Tbl_Resumen").Delete
                
            Else
                Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Add
                Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 1).Shape.TextFrame.TextRange.Text = "Total"
                Diapositiva.Shapes("Tbl_Resumen").Table.Cell(Diapositiva.Shapes("Tbl_Resumen").Table.Rows.Count, 2).Shape.TextFrame.TextRange.Text = Format(Total_TblRes, "#,##0.00")
                Diapositiva.Shapes("Tbl_Resumen").Table.LastRow = True
            End If
            
            
            
            
            For Ki = 1 To Renglon - Inicia_Tabla
            'MsgBox Grafica.SeriesCollection(4).Points(Ki).Top & " " & Diapositiva.Shapes(Grafica.Name).Top
                If gWorkBook.Worksheets(1).Range("U" & Ki + Inicia_Tabla).Value <> "" Then
                    Diapositiva.Shapes("Txt_Comentario").Duplicate
                    Diapositiva.Shapes("Flechita").Duplicate
                    Set TxtComent = Diapositiva.Shapes("Txt_Comentario")
                    Set Flechita = Diapositiva.Shapes("Flechita")
                    TxtComent.Name = "Txt_Comentario_" & Ki
                    Flechita.Name = "Flechita_" & Ki
                    Ancho_Barrita = Graf.SeriesCollection(4).Points(Ki).Height
                    Ancho_Coment = TxtComent.Height
                    
                    TxtComent.TextFrame.TextRange.Text = gWorkBook.Worksheets(1).Range("U" & Ki + Inicia_Tabla).Value
                    TxtComent.Top = Graf.SeriesCollection(4).Points(Ki).Top + Diapositiva.Shapes(Graf.Name).Top + ((Ancho_Barrita - Ancho_Coment) / 2)
                    TxtComent.Left = Diapositiva.Shapes(Graf.Name).Left + Diapositiva.Shapes(Graf.Name).Width + 30
                    Flechita.Top = Graf.SeriesCollection(4).Points(Ki).Top + Diapositiva.Shapes(Graf.Name).Top + (Ancho_Barrita / 2)
                    
                    Flechita.Left = Graf.SeriesCollection(4).DataLabels(Ki).Width + Graf.SeriesCollection(4).DataLabels(Ki).Left + 15
                    Flechita.Width = TxtComent.Left - (Graf.SeriesCollection(4).DataLabels(Ki).Width + Graf.SeriesCollection(4).DataLabels(Ki).Left + 15)
                End If
            Next
        
            
            With gWorkBook.Application
'                .Calculation = -4105
                .EnableEvents = True
                .ScreenUpdating = True
            End With
            gWorkBook.Close
'            gWorkBook.Application.Quit
              
            Diapositiva.Shapes("Txt_Comentario").Delete
            Diapositiva.Shapes("Flechita").Delete
            
        For I2 = 1 To Graf.SeriesCollection(1).DataLabels.Count
            If Graf.SeriesCollection(1).DataLabels(I2).Left + Graf.SeriesCollection(1).DataLabels(I2).Width > Graf.SeriesCollection(2).DataLabels(I2).Left Then
'                MsgBox "Se trasponen las cajitas"
                Graf.SeriesCollection(1).DataLabels(I2).Left = Graf.SeriesCollection(1).DataLabels(I2).Left - ((Graf.SeriesCollection(1).DataLabels(I2).Left + Graf.SeriesCollection(1).DataLabels(I2).Width) - Graf.SeriesCollection(2).DataLabels(I2).Left)
            End If
        Next
            
            Set Graf = Nothing
            Set gChartData = Nothing
            Set gWorkBook = Nothing
            
            '***************************************
        End If
    'Grafica.SeriesCollection(2).DataLabels(i).Font.Size = 11
    Next

    GoTo Salida_Pedidos_Calendarizados

Error_Pedidos_Calendarizados:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Procedimiento: Pedidos_Calendarizados", vbExclamation, "Error en pedidos calendarizados"

Salida_Pedidos_Calendarizados:
    On Error Resume Next

    If Not Rs_TblRed Is Nothing Then Rs_TblRed.Close
    If Not Rs Is Nothing Then Rs.Close
    If Not RsCia Is Nothing Then RsCia.Close
    
    If Not gWorkBook Is Nothing Then
        On Error Resume Next
        With gWorkBook.Application
            '.Calculation = -4105
            .EnableEvents = True
            .ScreenUpdating = True
        End With
        gWorkBook.Close SaveChanges:=False
        gWorkBook.Application.Quit
    End If

    Set gWorkSheet = Nothing
    Set gWorkBook = Nothing
    Set gChartData = Nothing
    Set Grafica = Nothing
    Set Graf = Nothing

    Set Qdf_TblRes = Nothing
    Set Qdf = Nothing
    Set Rs_TblRed = Nothing
    Set Rs = Nothing
    Set RsCia = Nothing

    Set Diapositiva = Nothing
    Set Diap = Nothing
    Set Tabla = Nothing
    Set Tbl2 = Nothing
    Set TxtComent = Nothing
    Set Flechita = Nothing
    Set Link = Nothing
    Set Botoncito = Nothing

    If Not Pwr Is Nothing Then
        'Pwr.Application.ScreenUpdating = True
    End If

    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

