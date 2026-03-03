Attribute VB_Name = "M01_ConfiguracionInicial"
'===============================================================================
' MODULO: M01_ConfiguracionInicial
' DESCRIPCION: Crea toda la estructura del libro (hojas, encabezados, formatos)
' AUTOR: Herramienta de Gestion - Flipping de Motos
' VERSION: 1.0
'===============================================================================
Option Explicit

' --- CONSTANTES GLOBALES ---
Public Const HOJA_MOTOS As String = "Registro_Motos"
Public Const HOJA_GASTOS As String = "Gastos"
Public Const HOJA_ESTADO As String = "Estado_Proyecto"
Public Const HOJA_GANANCIA As String = "Calculo_Ganancia"
Public Const HOJA_DASHBOARD As String = "Dashboard"
Public Const HOJA_CONFIG As String = "Configuracion"

Public Const COLOR_HEADER As Long = 2631720     ' Azul oscuro profesional (#002828)
Public Const COLOR_HEADER_TEXT As Long = 16777215 ' Blanco
Public Const COLOR_FILA_PAR As Long = 15921906   ' Gris muy claro
Public Const COLOR_ACCENT As Long = 5288000      ' Verde acento
Public Const COLOR_SOCIO_A As Long = 11893734    ' Azul claro
Public Const COLOR_SOCIO_B As Long = 10079487    ' Naranja claro

Public Sub CrearEstructuraCompleta()
    '===========================================================================
    ' Procedimiento principal: crea todas las hojas y su estructura
    '===========================================================================
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Eliminar hojas existentes (excepto la primera si esta vacia)
    Dim ws As Worksheet
    Do While ThisWorkbook.Sheets.Count > 1
        ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Delete
    Loop

    ' Crear hojas en orden
    CrearHojaConfiguracion
    CrearHojaRegistroMotos
    CrearHojaGastos
    CrearHojaEstadoProyecto
    CrearHojaCalculoGanancia
    CrearHojaDashboard

    ' Eliminar Hoja1 por defecto si existe
    On Error Resume Next
    ThisWorkbook.Sheets("Hoja1").Delete
    ThisWorkbook.Sheets("Sheet1").Delete
    On Error GoTo 0

    ' Activar Dashboard
    ThisWorkbook.Sheets(HOJA_DASHBOARD).Activate

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Estructura creada exitosamente." & vbCrLf & vbCrLf & _
           "Hojas creadas:" & vbCrLf & _
           "  - Configuracion" & vbCrLf & _
           "  - Registro_Motos" & vbCrLf & _
           "  - Gastos" & vbCrLf & _
           "  - Estado_Proyecto" & vbCrLf & _
           "  - Calculo_Ganancia" & vbCrLf & _
           "  - Dashboard", vbInformation, "Flipping de Motos"
End Sub

Private Sub CrearHojaConfiguracion()
    Dim ws As Worksheet
    Set ws = AgregarHoja(HOJA_CONFIG)

    ' Titulo
    FormatearTitulo ws, "A1:D1", "CONFIGURACION GENERAL"

    ' Datos de socios
    ws.Range("A3").Value = "Nombre Socio A:"
    ws.Range("B3").Value = "Socio A"
    ws.Range("A4").Value = "Nombre Socio B:"
    ws.Range("B4").Value = "Socio B"
    ws.Range("A6").Value = "% Ganancia Socio A:"
    ws.Range("B6").Value = 0.5
    ws.Range("B6").NumberFormat = "0%"
    ws.Range("A7").Value = "% Ganancia Socio B:"
    ws.Range("B7").Value = 0.5
    ws.Range("B7").NumberFormat = "0%"

    ' Validacion: suma = 100%
    ws.Range("A9").Value = "Validacion:"
    ws.Range("B9").Formula = "=IF(B6+B7=1,""OK"",""ERROR: debe sumar 100%"")"

    ' Formato
    ws.Range("A3:A9").Font.Bold = True
    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 25

    ' Estados posibles (para listas desplegables)
    ws.Range("D3").Value = "Estados:"
    ws.Range("D4").Value = "Desarmado"
    ws.Range("D5").Value = "Esperando repuestos"
    ws.Range("D6").Value = "En taller"
    ws.Range("D7").Value = "Lista para venta"
    ws.Range("D8").Value = "Vendida"
    ws.Range("D3").Font.Bold = True

    ' Categorias de gastos
    ws.Range("F3").Value = "Categorias Gasto:"
    ws.Range("F4").Value = "Repuestos"
    ws.Range("F5").Value = "Mano de obra"
    ws.Range("F6").Value = "Pintura"
    ws.Range("F7").Value = "Gestoria"
    ws.Range("F8").Value = "Transporte"
    ws.Range("F9").Value = "Seguro"
    ws.Range("F10").Value = "Otros"
    ws.Range("F3").Font.Bold = True

    ' Nombres definidos
    ThisWorkbook.Names.Add Name:="ListaEstados", RefersTo:="=" & HOJA_CONFIG & "!$D$4:$D$8"
    ThisWorkbook.Names.Add Name:="ListaCategorias", RefersTo:="=" & HOJA_CONFIG & "!$F$4:$F$10"
    ThisWorkbook.Names.Add Name:="PctSocioA", RefersTo:="=" & HOJA_CONFIG & "!$B$6"
    ThisWorkbook.Names.Add Name:="PctSocioB", RefersTo:="=" & HOJA_CONFIG & "!$B$7"
    ThisWorkbook.Names.Add Name:="NombreSocioA", RefersTo:="=" & HOJA_CONFIG & "!$B$3"
    ThisWorkbook.Names.Add Name:="NombreSocioB", RefersTo:="=" & HOJA_CONFIG & "!$B$4"

    ws.Tab.Color = RGB(128, 128, 128)
End Sub

Private Sub CrearHojaRegistroMotos()
    Dim ws As Worksheet
    Set ws = AgregarHoja(HOJA_MOTOS)

    ' Titulo
    FormatearTitulo ws, "A1:J1", "REGISTRO DE MOTOS"

    ' Encabezados - Fila 3
    Dim encabezados As Variant
    encabezados = Array("ID_Moto", "Marca", "Modelo", "Anio", "Cilindrada", _
                        "Precio_Compra", "Fecha_Compra", "Precio_Venta", _
                        "Fecha_Venta", "Observaciones")

    Dim i As Integer
    For i = 0 To UBound(encabezados)
        ws.Cells(3, i + 1).Value = encabezados(i)
    Next i

    FormatearEncabezados ws, "A3:J3"

    ' Anchos de columna
    ws.Columns("A").ColumnWidth = 10  ' ID
    ws.Columns("B").ColumnWidth = 14  ' Marca
    ws.Columns("C").ColumnWidth = 18  ' Modelo
    ws.Columns("D").ColumnWidth = 8   ' Anio
    ws.Columns("E").ColumnWidth = 12  ' Cilindrada
    ws.Columns("F").ColumnWidth = 16  ' Precio Compra
    ws.Columns("G").ColumnWidth = 14  ' Fecha Compra
    ws.Columns("H").ColumnWidth = 16  ' Precio Venta
    ws.Columns("I").ColumnWidth = 14  ' Fecha Venta
    ws.Columns("J").ColumnWidth = 30  ' Observaciones

    ' Formato de columnas
    ws.Columns("F:F").NumberFormat = "$#,##0.00"
    ws.Columns("H:H").NumberFormat = "$#,##0.00"
    ws.Columns("G:G").NumberFormat = "dd/mm/yyyy"
    ws.Columns("I:I").NumberFormat = "dd/mm/yyyy"

    ' Nombre definido para la tabla de motos
    ' Se actualizara dinamicamente

    ws.Tab.Color = RGB(0, 102, 204)
End Sub

Private Sub CrearHojaGastos()
    Dim ws As Worksheet
    Set ws = AgregarHoja(HOJA_GASTOS)

    ' Titulo
    FormatearTitulo ws, "A1:H1", "SEGUIMIENTO DE GASTOS POR MOTO"

    ' Encabezados - Fila 3
    Dim encabezados As Variant
    encabezados = Array("ID_Gasto", "ID_Moto", "Fecha", "Categoria", _
                        "Descripcion", "Monto", "Pagado_Por", "Comprobante")

    Dim i As Integer
    For i = 0 To UBound(encabezados)
        ws.Cells(3, i + 1).Value = encabezados(i)
    Next i

    FormatearEncabezados ws, "A3:H3"

    ' Anchos
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 30
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 18

    ' Formatos
    ws.Columns("C:C").NumberFormat = "dd/mm/yyyy"
    ws.Columns("F:F").NumberFormat = "$#,##0.00"

    ' Validacion: Pagado_Por solo acepta Socio A o Socio B
    ' Se aplicara en las primeras 500 filas de datos
    Dim rngPagado As Range
    Set rngPagado = ws.Range("G4:G503")
    With rngPagado.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=NombreSocioA&"",""&NombreSocioB"
        .ErrorMessage = "Seleccione Socio A o Socio B"
    End With

    ' Validacion: Categoria
    Dim rngCat As Range
    Set rngCat = ws.Range("D4:D503")
    With rngCat.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=ListaCategorias"
        .ErrorMessage = "Seleccione una categoria valida"
    End With

    ws.Tab.Color = RGB(255, 153, 0)
End Sub

Private Sub CrearHojaEstadoProyecto()
    Dim ws As Worksheet
    Set ws = AgregarHoja(HOJA_ESTADO)

    ' Titulo
    FormatearTitulo ws, "A1:G1", "ESTADO DEL PROYECTO - KANBAN"

    ' Encabezados
    Dim encabezados As Variant
    encabezados = Array("ID_Moto", "Moto (Marca Modelo)", "Estado_Actual", _
                        "Fecha_Cambio_Estado", "Dias_En_Estado", _
                        "Dias_Totales_Stock", "Notas")

    Dim i As Integer
    For i = 0 To UBound(encabezados)
        ws.Cells(3, i + 1).Value = encabezados(i)
    Next i

    FormatearEncabezados ws, "A3:G3"

    ' Anchos
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 25
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 30

    ' Formatos
    ws.Columns("D:D").NumberFormat = "dd/mm/yyyy"

    ' Validacion: Estado
    Dim rngEstado As Range
    Set rngEstado = ws.Range("C4:C503")
    With rngEstado.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=ListaEstados"
        .ErrorMessage = "Seleccione un estado valido"
    End With

    ' Formato condicional por estado
    AgregarFormatoCondicionalEstado ws

    ws.Tab.Color = RGB(0, 176, 80)
End Sub

Private Sub AgregarFormatoCondicionalEstado(ws As Worksheet)
    Dim rng As Range
    Set rng = ws.Range("C4:C503")
    rng.FormatConditions.Delete

    ' Desarmado - Rojo claro
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Desarmado""")
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With

    ' Esperando repuestos - Amarillo
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Esperando repuestos""")
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 101, 0)
    End With

    ' En taller - Azul claro
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""En taller""")
        .Interior.Color = RGB(189, 215, 238)
        .Font.Color = RGB(0, 51, 102)
    End With

    ' Lista para venta - Verde
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Lista para venta""")
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With

    ' Vendida - Gris
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Vendida""")
        .Interior.Color = RGB(217, 217, 217)
        .Font.Color = RGB(89, 89, 89)
    End With
End Sub

Private Sub CrearHojaCalculoGanancia()
    Dim ws As Worksheet
    Set ws = AgregarHoja(HOJA_GANANCIA)

    ' Titulo
    FormatearTitulo ws, "A1:L1", "CALCULO DE GANANCIA Y REPARTO"

    ' Encabezados
    Dim encabezados As Variant
    encabezados = Array("ID_Moto", "Moto", "Precio_Compra", _
                        "Gastos_Socio_A", "Gastos_Socio_B", "Total_Gastos", _
                        "Inversion_Total", "Precio_Venta", "Utilidad_Neta", _
                        "Devolucion_Socio_A", "Devolucion_Socio_B", "ROI_%")

    Dim i As Integer
    For i = 0 To UBound(encabezados)
        ws.Cells(3, i + 1).Value = encabezados(i)
    Next i

    FormatearEncabezados ws, "A3:L3"

    ' Anchos
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 16
    ws.Columns("J").ColumnWidth = 18
    ws.Columns("K").ColumnWidth = 18
    ws.Columns("L").ColumnWidth = 10

    ' Formatos numericos
    ws.Range("C:C,D:D,E:E,F:F,G:G,H:H,I:I,J:J,K:K").NumberFormat = "$#,##0.00"
    ws.Columns("L").NumberFormat = "0.0%"

    ' Fila de totales (fila 2 como resumen arriba de datos - se calculara)
    ' Los totales se agregan al final de cada columna cuando hay datos

    ' Sub-encabezado: Detalle del reparto
    ws.Range("N3").Value = "DETALLE DEL REPARTO"
    ws.Range("N3").Font.Bold = True
    ws.Range("N3").Font.Size = 11

    ws.Range("N4").Value = "Concepto"
    ws.Range("O4").Value = "Socio A"
    ws.Range("P4").Value = "Socio B"
    FormatearEncabezados ws, "N4:P4"

    ws.Range("N5").Value = "Inversion realizada"
    ws.Range("N6").Value = "Devolucion inversion"
    ws.Range("N7").Value = "Participacion ganancia"
    ws.Range("N8").Value = "TOTAL A RECIBIR"
    ws.Range("N8").Font.Bold = True

    ws.Columns("N").ColumnWidth = 24
    ws.Columns("O").ColumnWidth = 16
    ws.Columns("P").ColumnWidth = 16
    ws.Range("O5:P8").NumberFormat = "$#,##0.00"

    ' Formulas del resumen (se calculan sobre los datos de la tabla)
    ' Inversion realizada
    ws.Range("O5").Formula = "=SUMPRODUCT((" & HOJA_GASTOS & "!G4:G503=NombreSocioA)*(" & HOJA_GASTOS & "!F4:F503))"
    ws.Range("P5").Formula = "=SUMPRODUCT((" & HOJA_GASTOS & "!G4:G503=NombreSocioB)*(" & HOJA_GASTOS & "!F4:F503))"

    ' Devolucion inversion (igual a lo invertido)
    ws.Range("O6").Formula = "=O5"
    ws.Range("P6").Formula = "=P5"

    ' Participacion ganancia
    ws.Range("O7").Formula = "=SUM(I4:I503)*PctSocioA"
    ws.Range("P7").Formula = "=SUM(I4:I503)*PctSocioB"

    ' Total a recibir
    ws.Range("O8").Formula = "=O6+O7"
    ws.Range("P8").Formula = "=P6+P7"

    ' Formato condicional para utilidad
    Dim rngUtil As Range
    Set rngUtil = ws.Range("I4:I503")
    rngUtil.FormatConditions.Delete

    With rngUtil.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
        .Font.Color = RGB(200, 0, 0)
        .Interior.Color = RGB(255, 230, 230)
    End With
    With rngUtil.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Font.Color = RGB(0, 128, 0)
    End With

    ws.Tab.Color = RGB(112, 48, 160)
End Sub

Private Sub CrearHojaDashboard()
    Dim ws As Worksheet
    Set ws = AgregarHoja(HOJA_DASHBOARD)

    ' Titulo
    FormatearTitulo ws, "A1:M1", "DASHBOARD - FLIPPING DE MOTOS"

    ' === SECCION KPIs (Fila 3) ===
    ' KPI 1: Total Motos
    ws.Range("A3").Value = "Total Motos"
    ws.Range("A4").Formula = "=COUNTA(" & HOJA_MOTOS & "!A4:A503)"
    ws.Range("A4").Font.Size = 24
    ws.Range("A4").Font.Bold = True
    FormatearKPI ws, "A3:B4"

    ' KPI 2: Motos Activas
    ws.Range("C3").Value = "Motos Activas"
    ws.Range("C4").Formula = "=COUNTA(" & HOJA_MOTOS & "!A4:A503)-COUNTIF(" & HOJA_ESTADO & "!C4:C503,""Vendida"")"
    ws.Range("C4").Font.Size = 24
    ws.Range("C4").Font.Bold = True
    FormatearKPI ws, "C3:D4"

    ' KPI 3: Motos Vendidas
    ws.Range("E3").Value = "Motos Vendidas"
    ws.Range("E4").Formula = "=COUNTIF(" & HOJA_ESTADO & "!C4:C503,""Vendida"")"
    ws.Range("E4").Font.Size = 24
    ws.Range("E4").Font.Bold = True
    FormatearKPI ws, "E3:F4"

    ' KPI 4: Inversion Total Acumulada
    ws.Range("G3").Value = "Inversion Total"
    ws.Range("G4").Formula = "=SUM(" & HOJA_MOTOS & "!F4:F503)+SUM(" & HOJA_GASTOS & "!F4:F503)"
    ws.Range("G4").Font.Size = 18
    ws.Range("G4").Font.Bold = True
    ws.Range("G4").NumberFormat = "$#,##0"
    FormatearKPI ws, "G3:H4"

    ' KPI 5: Ganancia Neta Total
    ws.Range("I3").Value = "Ganancia Neta Total"
    ws.Range("I4").Formula = "=SUM(" & HOJA_GANANCIA & "!I4:I503)"
    ws.Range("I4").Font.Size = 18
    ws.Range("I4").Font.Bold = True
    ws.Range("I4").NumberFormat = "$#,##0"
    FormatearKPI ws, "I3:J4"

    ' KPI 6: ROI Promedio
    ws.Range("K3").Value = "ROI Promedio"
    ws.Range("K4").Formula = "=IFERROR(AVERAGE(" & HOJA_GANANCIA & "!L4:L503),0)"
    ws.Range("K4").Font.Size = 18
    ws.Range("K4").Font.Bold = True
    ws.Range("K4").NumberFormat = "0.0%"
    FormatearKPI ws, "K3:L4"

    ' === SECCION: Datos para graficos (filas 6+) ===
    ' Inversion por socio
    ws.Range("A6").Value = "Inversion por Socio"
    ws.Range("A6").Font.Bold = True
    ws.Range("A6").Font.Size = 11
    ws.Range("A7").Formula = "=NombreSocioA"
    ws.Range("B7").Formula = "=SUMPRODUCT((" & HOJA_GASTOS & "!G4:G503=NombreSocioA)*(" & HOJA_GASTOS & "!F4:F503))"
    ws.Range("A8").Formula = "=NombreSocioB"
    ws.Range("B8").Formula = "=SUMPRODUCT((" & HOJA_GASTOS & "!G4:G503=NombreSocioB)*(" & HOJA_GASTOS & "!F4:F503))"
    ws.Range("B7:B8").NumberFormat = "$#,##0.00"

    ' Dias promedio de stock
    ws.Range("A10").Value = "Dias Prom. Stock"
    ws.Range("A10").Font.Bold = True
    ws.Range("A11").Formula = "=IFERROR(AVERAGE(" & HOJA_ESTADO & "!F4:F503),0)"
    ws.Range("A11").NumberFormat = "0"
    ws.Range("A11").Font.Size = 16

    ' Flujo de caja
    ws.Range("D6").Value = "Flujo de Caja"
    ws.Range("D6").Font.Bold = True
    ws.Range("D6").Font.Size = 11
    ws.Range("D7").Value = "Ingresos (Ventas)"
    ws.Range("E7").Formula = "=SUM(" & HOJA_MOTOS & "!H4:H503)"
    ws.Range("D8").Value = "Egresos (Compras)"
    ws.Range("E8").Formula = "=SUM(" & HOJA_MOTOS & "!F4:F503)"
    ws.Range("D9").Value = "Egresos (Gastos)"
    ws.Range("E9").Formula = "=SUM(" & HOJA_GASTOS & "!F4:F503)"
    ws.Range("D10").Value = "FLUJO NETO"
    ws.Range("D10").Font.Bold = True
    ws.Range("E10").Formula = "=E7-E8-E9"
    ws.Range("E10").Font.Bold = True
    ws.Range("E7:E10").NumberFormat = "$#,##0.00"

    ' Estado de motos (conteo por estado)
    ws.Range("H6").Value = "Motos por Estado"
    ws.Range("H6").Font.Bold = True
    ws.Range("H6").Font.Size = 11
    ws.Range("H7").Value = "Desarmado"
    ws.Range("I7").Formula = "=COUNTIF(" & HOJA_ESTADO & "!C4:C503,""Desarmado"")"
    ws.Range("H8").Value = "Esperando repuestos"
    ws.Range("I8").Formula = "=COUNTIF(" & HOJA_ESTADO & "!C4:C503,""Esperando repuestos"")"
    ws.Range("H9").Value = "En taller"
    ws.Range("I9").Formula = "=COUNTIF(" & HOJA_ESTADO & "!C4:C503,""En taller"")"
    ws.Range("H10").Value = "Lista para venta"
    ws.Range("I10").Formula = "=COUNTIF(" & HOJA_ESTADO & "!C4:C503,""Lista para venta"")"
    ws.Range("H11").Value = "Vendida"
    ws.Range("I11").Formula = "=COUNTIF(" & HOJA_ESTADO & "!C4:C503,""Vendida"")"

    ' Anchos de columna
    Dim col As Integer
    For col = 1 To 13
        ws.Columns(col).ColumnWidth = 16
    Next col

    ws.Tab.Color = RGB(0, 176, 80)
End Sub

' ============================================================
' UTILIDADES DE FORMATO
' ============================================================
Private Function AgregarHoja(nombre As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombre)
    If Not ws Is Nothing Then
        ws.Delete
    End If
    On Error GoTo 0

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = nombre

    ' Config general de hoja
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10

    Set AgregarHoja = ws
End Function

Private Sub FormatearTitulo(ws As Worksheet, rango As String, titulo As String)
    With ws.Range(rango)
        .Merge
        .Value = titulo
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 40, 40)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 35
    End With
End Sub

Private Sub FormatearEncabezados(ws As Worksheet, rango As String)
    With ws.Range(rango)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 80, 120)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 25
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(0, 40, 40)
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
End Sub

Private Sub FormatearKPI(ws As Worksheet, rango As String)
    With ws.Range(rango)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    ' Titulo del KPI
    With ws.Range(Left(rango, InStr(rango, ":") - 1))
        .Font.Size = 9
        .Font.Color = RGB(100, 100, 100)
    End With
End Sub
