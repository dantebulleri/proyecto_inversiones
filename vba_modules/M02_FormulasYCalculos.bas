Attribute VB_Name = "M02_FormulasYCalculos"
'===============================================================================
' MODULO: M02_FormulasYCalculos
' DESCRIPCION: Inserta formulas automaticas en las hojas de calculo
'              y actualiza los calculos del Dashboard
' VERSION: 1.0
'===============================================================================
Option Explicit

Public Sub InsertarFormulasRegistroMoto(fila As Long)
    '===========================================================================
    ' Inserta formulas automaticas al registrar una nueva moto
    ' Se llama desde el formulario de alta de motos
    '===========================================================================
    Dim wsEstado As Worksheet
    Dim wsGanancia As Worksheet
    Dim wsMotos As Worksheet

    Set wsMotos = ThisWorkbook.Sheets(HOJA_MOTOS)
    Set wsEstado = ThisWorkbook.Sheets(HOJA_ESTADO)
    Set wsGanancia = ThisWorkbook.Sheets(HOJA_GANANCIA)

    Dim idMoto As String
    idMoto = wsMotos.Cells(fila, 1).Value

    ' --- Hoja Estado_Proyecto ---
    Dim filaEstado As Long
    filaEstado = ObtenerSiguienteFilaVacia(wsEstado, "A", 4)

    wsEstado.Cells(filaEstado, 1).Value = idMoto
    ' Moto (concatenar Marca + Modelo)
    wsEstado.Cells(filaEstado, 2).Formula = _
        "=IFERROR(INDEX(" & HOJA_MOTOS & "!B4:B503,MATCH(A" & filaEstado & "," & HOJA_MOTOS & "!A4:A503,0))&"" ""&INDEX(" & HOJA_MOTOS & "!C4:C503,MATCH(A" & filaEstado & "," & HOJA_MOTOS & "!A4:A503,0)),"""")"
    wsEstado.Cells(filaEstado, 3).Value = "Desarmado"
    wsEstado.Cells(filaEstado, 4).Value = Date
    wsEstado.Cells(filaEstado, 4).NumberFormat = "dd/mm/yyyy"
    ' Dias en estado actual
    wsEstado.Cells(filaEstado, 5).Formula = "=IF(D" & filaEstado & "<>"""",TODAY()-D" & filaEstado & ","""")"
    ' Dias totales de stock
    wsEstado.Cells(filaEstado, 6).Formula = _
        "=IFERROR(IF(INDEX(" & HOJA_MOTOS & "!I4:I503,MATCH(A" & filaEstado & "," & HOJA_MOTOS & "!A4:A503,0))<>"""",INDEX(" & HOJA_MOTOS & "!I4:I503,MATCH(A" & filaEstado & "," & HOJA_MOTOS & "!A4:A503,0))-INDEX(" & HOJA_MOTOS & "!G4:G503,MATCH(A" & filaEstado & "," & HOJA_MOTOS & "!A4:A503,0)),TODAY()-INDEX(" & HOJA_MOTOS & "!G4:G503,MATCH(A" & filaEstado & "," & HOJA_MOTOS & "!A4:A503,0))),0)"

    ' --- Hoja Calculo_Ganancia ---
    Dim filaGanancia As Long
    filaGanancia = ObtenerSiguienteFilaVacia(wsGanancia, "A", 4)

    wsGanancia.Cells(filaGanancia, 1).Value = idMoto
    ' Moto
    wsGanancia.Cells(filaGanancia, 2).Formula = _
        "=IFERROR(INDEX(" & HOJA_MOTOS & "!B4:B503,MATCH(A" & filaGanancia & "," & HOJA_MOTOS & "!A4:A503,0))&"" ""&INDEX(" & HOJA_MOTOS & "!C4:C503,MATCH(A" & filaGanancia & "," & HOJA_MOTOS & "!A4:A503,0)),"""")"
    ' Precio Compra
    wsGanancia.Cells(filaGanancia, 3).Formula = _
        "=IFERROR(INDEX(" & HOJA_MOTOS & "!F4:F503,MATCH(A" & filaGanancia & "," & HOJA_MOTOS & "!A4:A503,0)),0)"
    ' Gastos Socio A
    wsGanancia.Cells(filaGanancia, 4).Formula = _
        "=SUMPRODUCT((" & HOJA_GASTOS & "!B4:B503=A" & filaGanancia & ")*(" & HOJA_GASTOS & "!G4:G503=NombreSocioA)*(" & HOJA_GASTOS & "!F4:F503))"
    ' Gastos Socio B
    wsGanancia.Cells(filaGanancia, 5).Formula = _
        "=SUMPRODUCT((" & HOJA_GASTOS & "!B4:B503=A" & filaGanancia & ")*(" & HOJA_GASTOS & "!G4:G503=NombreSocioB)*(" & HOJA_GASTOS & "!F4:F503))"
    ' Total Gastos
    wsGanancia.Cells(filaGanancia, 6).Formula = "=D" & filaGanancia & "+E" & filaGanancia
    ' Inversion Total (Compra + Gastos)
    wsGanancia.Cells(filaGanancia, 7).Formula = "=C" & filaGanancia & "+F" & filaGanancia
    ' Precio Venta
    wsGanancia.Cells(filaGanancia, 8).Formula = _
        "=IFERROR(INDEX(" & HOJA_MOTOS & "!H4:H503,MATCH(A" & filaGanancia & "," & HOJA_MOTOS & "!A4:A503,0)),0)"
    ' Utilidad Neta
    wsGanancia.Cells(filaGanancia, 9).Formula = "=IF(H" & filaGanancia & ">0,H" & filaGanancia & "-G" & filaGanancia & ",0)"
    ' Devolucion Socio A = su inversion + % de ganancia
    wsGanancia.Cells(filaGanancia, 10).Formula = _
        "=IF(H" & filaGanancia & ">0,(C" & filaGanancia & "/2)+D" & filaGanancia & "+IF(I" & filaGanancia & ">0,I" & filaGanancia & "*PctSocioA,I" & filaGanancia & "*0.5),0)"
    ' Devolucion Socio B = su inversion + % de ganancia
    wsGanancia.Cells(filaGanancia, 11).Formula = _
        "=IF(H" & filaGanancia & ">0,(C" & filaGanancia & "/2)+E" & filaGanancia & "+IF(I" & filaGanancia & ">0,I" & filaGanancia & "*PctSocioB,I" & filaGanancia & "*0.5),0)"
    ' ROI %
    wsGanancia.Cells(filaGanancia, 12).Formula = _
        "=IFERROR(I" & filaGanancia & "/G" & filaGanancia & ",0)"
End Sub

Public Sub ActualizarGraficos()
    '===========================================================================
    ' Crea o actualiza los graficos del Dashboard
    '===========================================================================
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_DASHBOARD)

    ' Eliminar graficos existentes
    Dim cht As ChartObject
    For Each cht In ws.ChartObjects
        cht.Delete
    Next cht

    ' --- Grafico 1: Inversion por Socio (Barras) ---
    CrearGraficoInversionSocios ws

    ' --- Grafico 2: ROI por Moto (Barras) ---
    CrearGraficoROI ws

    ' --- Grafico 3: Estado de Motos (Torta) ---
    CrearGraficoEstados ws

    ' --- Grafico 4: Flujo de Caja (Barras) ---
    CrearGraficoFlujoCaja ws
End Sub

Private Sub CrearGraficoInversionSocios(ws As Worksheet)
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=10, Top:=240, Width:=350, Height:=250)

    With cht.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=ws.Range("A7:B8")
        .HasTitle = True
        .ChartTitle.Text = "Inversion Acumulada por Socio"
        .ChartTitle.Font.Size = 11
        .ChartTitle.Font.Bold = True

        ' Formato de barras
        If .SeriesCollection.Count > 0 Then
            .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 102, 204)
        End If

        .HasLegend = False
        .Axes(xlValue).NumberFormat = "$#,##0"
        .Axes(xlValue).HasTitle = False

        ' Estilo limpio
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(250, 250, 250)
        .ChartArea.Format.Line.ForeColor.RGB = RGB(200, 200, 200)
    End With
End Sub

Private Sub CrearGraficoROI(ws As Worksheet)
    Dim wsG As Worksheet
    Set wsG = ThisWorkbook.Sheets(HOJA_GANANCIA)

    ' Contar filas con datos
    Dim ultimaFila As Long
    ultimaFila = ObtenerUltimaFilaConDatos(wsG, "A", 4)

    If ultimaFila < 4 Then Exit Sub

    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=370, Top:=240, Width:=350, Height:=250)

    With cht.Chart
        .ChartType = xlColumnClustered

        Dim ser As Series
        Set ser = .SeriesCollection.NewSeries
        ser.Values = wsG.Range("L4:L" & ultimaFila)
        ser.XValues = wsG.Range("B4:B" & ultimaFila)
        ser.Name = "ROI"

        .HasTitle = True
        .ChartTitle.Text = "ROI por Moto"
        .ChartTitle.Font.Size = 11

        ser.Format.Fill.ForeColor.RGB = RGB(0, 176, 80)

        .HasLegend = False
        .Axes(xlValue).NumberFormat = "0%"
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(250, 250, 250)
        .ChartArea.Format.Line.ForeColor.RGB = RGB(200, 200, 200)
    End With
End Sub

Private Sub CrearGraficoEstados(ws As Worksheet)
    ' Solo crear si hay datos
    If Application.WorksheetFunction.Sum(ws.Range("I7:I11")) = 0 Then Exit Sub

    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=730, Top:=240, Width:=350, Height:=250)

    With cht.Chart
        .ChartType = xlPie
        .SetSourceData Source:=ws.Range("H7:I11")
        .HasTitle = True
        .ChartTitle.Text = "Motos por Estado"
        .ChartTitle.Font.Size = 11

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8

        ' Colores por estado
        If .SeriesCollection.Count > 0 Then
            Dim s As Series
            Set s = .SeriesCollection(1)
            s.Points(1).Format.Fill.ForeColor.RGB = RGB(255, 199, 206)  ' Desarmado
            s.Points(2).Format.Fill.ForeColor.RGB = RGB(255, 235, 156)  ' Esperando
            s.Points(3).Format.Fill.ForeColor.RGB = RGB(189, 215, 238)  ' Taller
            s.Points(4).Format.Fill.ForeColor.RGB = RGB(198, 239, 206)  ' Lista
            s.Points(5).Format.Fill.ForeColor.RGB = RGB(217, 217, 217)  ' Vendida
        End If

        .PlotArea.Format.Fill.Visible = msoFalse
        .ChartArea.Format.Line.ForeColor.RGB = RGB(200, 200, 200)
    End With
End Sub

Private Sub CrearGraficoFlujoCaja(ws As Worksheet)
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=10, Top:=500, Width:=350, Height:=250)

    With cht.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=ws.Range("D7:E10")
        .HasTitle = True
        .ChartTitle.Text = "Flujo de Caja"
        .ChartTitle.Font.Size = 11

        If .SeriesCollection.Count > 0 Then
            .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 102, 204)
        End If

        .HasLegend = False
        .Axes(xlValue).NumberFormat = "$#,##0"
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(250, 250, 250)
        .ChartArea.Format.Line.ForeColor.RGB = RGB(200, 200, 200)
    End With
End Sub

' ============================================================
' FUNCIONES AUXILIARES
' ============================================================
Public Function ObtenerSiguienteFilaVacia(ws As Worksheet, columna As String, filaInicio As Long) As Long
    Dim fila As Long
    fila = filaInicio
    Do While ws.Range(columna & fila).Value <> ""
        fila = fila + 1
    Loop
    ObtenerSiguienteFilaVacia = fila
End Function

Public Function ObtenerUltimaFilaConDatos(ws As Worksheet, columna As String, filaInicio As Long) As Long
    Dim ultimaCelda As Range
    Set ultimaCelda = ws.Range(columna & ws.Rows.Count).End(xlUp)
    If ultimaCelda.Row < filaInicio Then
        ObtenerUltimaFilaConDatos = filaInicio - 1
    Else
        ObtenerUltimaFilaConDatos = ultimaCelda.Row
    End If
End Function

Public Function GenerarIDMoto() As String
    '===========================================================================
    ' Genera un ID secuencial para motos: MOTO-001, MOTO-002, etc.
    '===========================================================================
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_MOTOS)

    Dim ultimaFila As Long
    ultimaFila = ObtenerUltimaFilaConDatos(ws, "A", 4)

    If ultimaFila < 4 Then
        GenerarIDMoto = "MOTO-001"
    Else
        Dim ultimoID As String
        ultimoID = ws.Cells(ultimaFila, 1).Value
        Dim numActual As Long
        numActual = CLng(Mid(ultimoID, 6))  ' Despues de "MOTO-"
        GenerarIDMoto = "MOTO-" & Format(numActual + 1, "000")
    End If
End Function

Public Function GenerarIDGasto() As String
    '===========================================================================
    ' Genera un ID secuencial para gastos: GASTO-001, GASTO-002, etc.
    '===========================================================================
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_GASTOS)

    Dim ultimaFila As Long
    ultimaFila = ObtenerUltimaFilaConDatos(ws, "A", 4)

    If ultimaFila < 4 Then
        GenerarIDGasto = "GASTO-001"
    Else
        Dim ultimoID As String
        ultimoID = ws.Cells(ultimaFila, 1).Value
        Dim numActual As Long
        numActual = CLng(Mid(ultimoID, 7))  ' Despues de "GASTO-"
        GenerarIDGasto = "GASTO-" & Format(numActual + 1, "000")
    End If
End Function

Public Function ObtenerListaMotos() As String()
    '===========================================================================
    ' Devuelve un array con los IDs de motos + descripcion
    '===========================================================================
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_MOTOS)

    Dim ultimaFila As Long
    ultimaFila = ObtenerUltimaFilaConDatos(ws, "A", 4)

    Dim resultado() As String

    If ultimaFila < 4 Then
        ReDim resultado(0)
        resultado(0) = "(Sin motos registradas)"
        ObtenerListaMotos = resultado
        Exit Function
    End If

    Dim cantidad As Long
    cantidad = ultimaFila - 3
    ReDim resultado(1 To cantidad)

    Dim i As Long
    For i = 4 To ultimaFila
        resultado(i - 3) = ws.Cells(i, 1).Value & " - " & _
                           ws.Cells(i, 2).Value & " " & ws.Cells(i, 3).Value
    Next i

    ObtenerListaMotos = resultado
End Function
