VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegistrarVenta
   Caption         =   "Registrar Venta de Moto"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "frmRegistrarVenta.frx":0000
   StartUpPosition =   1  'CenterOwner
End

' ============================================================
' USERFORM: frmRegistrarVenta
' Formulario para registrar la venta de una moto
' ============================================================
' CONTROLES NECESARIOS:
'   lblMoto         - Label "Moto a vender:"
'   cboMoto         - ComboBox (motos no vendidas)
'   lblPrecio       - Label "Precio de Venta ($):"
'   txtPrecioVenta  - TextBox
'   lblFecha        - Label "Fecha de Venta:"
'   txtFechaVenta   - TextBox
'   lblResumen      - Label (muestra inversion total)
'   btnVender       - CommandButton "Registrar Venta"
'   btnCancelar     - CommandButton "Cancelar"
' ============================================================

Option Explicit

Private Sub UserForm_Initialize()
    ' Cargar solo motos NO vendidas
    Dim wsMotos As Worksheet
    Dim wsEstado As Worksheet
    Set wsMotos = ThisWorkbook.Sheets(HOJA_MOTOS)
    Set wsEstado = ThisWorkbook.Sheets(HOJA_ESTADO)

    Dim ultimaFila As Long
    ultimaFila = ObtenerUltimaFilaConDatos(wsMotos, "A", 4)

    Dim i As Long
    For i = 4 To ultimaFila
        ' Verificar que no este vendida
        Dim filaEstado As Long
        Dim vendida As Boolean
        vendida = False

        Dim j As Long
        Dim ultEstado As Long
        ultEstado = ObtenerUltimaFilaConDatos(wsEstado, "A", 4)

        For j = 4 To ultEstado
            If wsEstado.Cells(j, 1).Value = wsMotos.Cells(i, 1).Value Then
                If wsEstado.Cells(j, 3).Value = "Vendida" Then vendida = True
                Exit For
            End If
        Next j

        If Not vendida Then
            cboMoto.AddItem wsMotos.Cells(i, 1).Value & " - " & _
                           wsMotos.Cells(i, 2).Value & " " & wsMotos.Cells(i, 3).Value
        End If
    Next i

    txtFechaVenta.Value = Format(Date, "dd/mm/yyyy")
    Me.BackColor = RGB(245, 245, 245)
End Sub

Private Sub cboMoto_Change()
    ' Al seleccionar moto, mostrar resumen de inversion
    If cboMoto.ListIndex = -1 Then Exit Sub

    Dim idMoto As String
    idMoto = Left(cboMoto.Value, InStr(cboMoto.Value, " - ") - 1)

    Dim wsMotos As Worksheet
    Dim wsGastos As Worksheet
    Set wsMotos = ThisWorkbook.Sheets(HOJA_MOTOS)
    Set wsGastos = ThisWorkbook.Sheets(HOJA_GASTOS)

    ' Buscar precio de compra
    Dim precioCompra As Double
    Dim fila As Long
    For fila = 4 To ObtenerUltimaFilaConDatos(wsMotos, "A", 4)
        If wsMotos.Cells(fila, 1).Value = idMoto Then
            precioCompra = wsMotos.Cells(fila, 6).Value
            Exit For
        End If
    Next fila

    ' Sumar gastos
    Dim totalGastos As Double
    For fila = 4 To ObtenerUltimaFilaConDatos(wsGastos, "A", 4)
        If wsGastos.Cells(fila, 2).Value = idMoto Then
            totalGastos = totalGastos + wsGastos.Cells(fila, 6).Value
        End If
    Next fila

    lblResumen.Caption = "Inversion total: $" & Format(precioCompra + totalGastos, "#,##0.00") & _
                         "  (Compra: $" & Format(precioCompra, "#,##0.00") & _
                         " + Gastos: $" & Format(totalGastos, "#,##0.00") & ")"
End Sub

Private Sub btnVender_Click()
    If cboMoto.ListIndex = -1 Then
        MsgBox "Seleccione una moto.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtPrecioVenta.Value) Or CDbl(txtPrecioVenta.Value) <= 0 Then
        MsgBox "Ingrese un precio de venta valido.", vbExclamation
        Exit Sub
    End If

    If Not IsDate(txtFechaVenta.Value) Then
        MsgBox "Ingrese una fecha valida.", vbExclamation
        Exit Sub
    End If

    Dim idMoto As String
    idMoto = Left(cboMoto.Value, InStr(cboMoto.Value, " - ") - 1)

    Dim confirmar As VbMsgBoxResult
    confirmar = MsgBox("Confirmar venta de " & cboMoto.Value & vbCrLf & _
                       "por $" & Format(CDbl(txtPrecioVenta.Value), "#,##0.00") & "?", _
                       vbYesNo + vbQuestion, "Confirmar Venta")

    If confirmar = vbYes Then
        RegistrarVenta idMoto, CDbl(txtPrecioVenta.Value), CDate(txtFechaVenta.Value)
        Unload Me
    End If
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub
