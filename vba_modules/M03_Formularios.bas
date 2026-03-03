Attribute VB_Name = "M03_Formularios"
'===============================================================================
' MODULO: M03_Formularios
' DESCRIPCION: Logica de los formularios UserForm para ingreso de datos
'              (Se complementa con los archivos .frm de UserForms)
' VERSION: 1.0
'===============================================================================
Option Explicit

' ============================================================
' FORMULARIO: ALTA DE MOTO
' ============================================================
Public Sub MostrarFormularioAltaMoto()
    frmAltaMoto.Show
End Sub

Public Sub RegistrarMoto(marca As String, modelo As String, anio As String, _
                         cilindrada As String, precioCompra As Double, _
                         fechaCompra As Date, observaciones As String)
    '===========================================================================
    ' Registra una nueva moto en la hoja de Registro y dispara las formulas
    '===========================================================================
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_MOTOS)

    Dim fila As Long
    fila = ObtenerSiguienteFilaVacia(ws, "A", 4)

    Dim idMoto As String
    idMoto = GenerarIDMoto()

    ws.Cells(fila, 1).Value = idMoto
    ws.Cells(fila, 2).Value = marca
    ws.Cells(fila, 3).Value = modelo
    ws.Cells(fila, 4).Value = anio
    ws.Cells(fila, 5).Value = cilindrada
    ws.Cells(fila, 6).Value = precioCompra
    ws.Cells(fila, 7).Value = fechaCompra
    ws.Cells(fila, 8).Value = ""  ' Precio venta - vacio
    ws.Cells(fila, 9).Value = ""  ' Fecha venta - vacio
    ws.Cells(fila, 10).Value = observaciones

    ' Formato de fila
    ws.Cells(fila, 6).NumberFormat = "$#,##0.00"
    ws.Cells(fila, 7).NumberFormat = "dd/mm/yyyy"

    ' Bandas de color alternas
    If fila Mod 2 = 0 Then
        ws.Range("A" & fila & ":J" & fila).Interior.Color = RGB(242, 242, 242)
    End If

    ' Insertar formulas en Estado y Ganancia
    InsertarFormulasRegistroMoto fila

    MsgBox "Moto registrada exitosamente." & vbCrLf & vbCrLf & _
           "ID: " & idMoto & vbCrLf & _
           "Moto: " & marca & " " & modelo, vbInformation, "Registro de Moto"
End Sub

' ============================================================
' FORMULARIO: REGISTRO DE GASTO
' ============================================================
Public Sub MostrarFormularioGasto()
    frmRegistroGasto.Show
End Sub

Public Sub RegistrarGasto(idMoto As String, fecha As Date, categoria As String, _
                          descripcion As String, monto As Double, pagadoPor As String, _
                          comprobante As String)
    '===========================================================================
    ' Registra un gasto asociado a una moto
    '===========================================================================
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_GASTOS)

    Dim fila As Long
    fila = ObtenerSiguienteFilaVacia(ws, "A", 4)

    ws.Cells(fila, 1).Value = GenerarIDGasto()
    ws.Cells(fila, 2).Value = idMoto
    ws.Cells(fila, 3).Value = fecha
    ws.Cells(fila, 4).Value = categoria
    ws.Cells(fila, 5).Value = descripcion
    ws.Cells(fila, 6).Value = monto
    ws.Cells(fila, 7).Value = pagadoPor
    ws.Cells(fila, 8).Value = comprobante

    ' Formatos
    ws.Cells(fila, 3).NumberFormat = "dd/mm/yyyy"
    ws.Cells(fila, 6).NumberFormat = "$#,##0.00"

    ' Bandas de color alternas
    If fila Mod 2 = 0 Then
        ws.Range("A" & fila & ":H" & fila).Interior.Color = RGB(242, 242, 242)
    End If

    MsgBox "Gasto registrado exitosamente." & vbCrLf & vbCrLf & _
           "Moto: " & idMoto & vbCrLf & _
           "Monto: $" & Format(monto, "#,##0.00") & vbCrLf & _
           "Pagado por: " & pagadoPor, vbInformation, "Registro de Gasto"
End Sub

' ============================================================
' FORMULARIO: REGISTRAR VENTA
' ============================================================
Public Sub MostrarFormularioVenta()
    frmRegistrarVenta.Show
End Sub

Public Sub RegistrarVenta(idMoto As String, precioVenta As Double, fechaVenta As Date)
    '===========================================================================
    ' Registra la venta de una moto y actualiza el estado
    '===========================================================================
    Dim wsMotos As Worksheet
    Dim wsEstado As Worksheet

    Set wsMotos = ThisWorkbook.Sheets(HOJA_MOTOS)
    Set wsEstado = ThisWorkbook.Sheets(HOJA_ESTADO)

    ' Buscar la fila de la moto en Registro_Motos
    Dim filaMotos As Long
    filaMotos = BuscarFilaPorID(wsMotos, "A", idMoto, 4)

    If filaMotos = 0 Then
        MsgBox "No se encontro la moto con ID: " & idMoto, vbExclamation
        Exit Sub
    End If

    ' Registrar venta
    wsMotos.Cells(filaMotos, 8).Value = precioVenta
    wsMotos.Cells(filaMotos, 8).NumberFormat = "$#,##0.00"
    wsMotos.Cells(filaMotos, 9).Value = fechaVenta
    wsMotos.Cells(filaMotos, 9).NumberFormat = "dd/mm/yyyy"

    ' Actualizar estado a "Vendida"
    Dim filaEstado As Long
    filaEstado = BuscarFilaPorID(wsEstado, "A", idMoto, 4)
    If filaEstado > 0 Then
        wsEstado.Cells(filaEstado, 3).Value = "Vendida"
        wsEstado.Cells(filaEstado, 4).Value = fechaVenta
    End If

    ' Actualizar graficos
    ActualizarGraficos

    MsgBox "Venta registrada exitosamente." & vbCrLf & vbCrLf & _
           "Moto: " & idMoto & vbCrLf & _
           "Precio de venta: $" & Format(precioVenta, "#,##0.00"), _
           vbInformation, "Venta Registrada"
End Sub

' ============================================================
' ACTUALIZAR ESTADO DE MOTO
' ============================================================
Public Sub MostrarFormularioEstado()
    frmCambiarEstado.Show
End Sub

Public Sub CambiarEstadoMoto(idMoto As String, nuevoEstado As String)
    Dim wsEstado As Worksheet
    Set wsEstado = ThisWorkbook.Sheets(HOJA_ESTADO)

    Dim fila As Long
    fila = BuscarFilaPorID(wsEstado, "A", idMoto, 4)

    If fila = 0 Then
        MsgBox "No se encontro la moto.", vbExclamation
        Exit Sub
    End If

    wsEstado.Cells(fila, 3).Value = nuevoEstado
    wsEstado.Cells(fila, 4).Value = Date
    wsEstado.Cells(fila, 7).Value = wsEstado.Cells(fila, 7).Value & _
        vbCrLf & Format(Date, "dd/mm/yyyy") & " - Cambio a: " & nuevoEstado

    MsgBox "Estado actualizado a: " & nuevoEstado, vbInformation, "Cambio de Estado"
End Sub

' ============================================================
' UTILIDAD: Buscar fila por ID
' ============================================================
Private Function BuscarFilaPorID(ws As Worksheet, columna As String, _
                                  idBuscado As String, filaInicio As Long) As Long
    Dim fila As Long
    Dim ultimaFila As Long
    ultimaFila = ObtenerUltimaFilaConDatos(ws, columna, filaInicio)

    For fila = filaInicio To ultimaFila
        If ws.Range(columna & fila).Value = idBuscado Then
            BuscarFilaPorID = fila
            Exit Function
        End If
    Next fila

    BuscarFilaPorID = 0  ' No encontrado
End Function
