VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCambiarEstado
   Caption         =   "Cambiar Estado de Moto"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "frmCambiarEstado.frx":0000
   StartUpPosition =   1  'CenterOwner
End

' ============================================================
' USERFORM: frmCambiarEstado
' Formulario para cambiar el estado/fase de una moto
' ============================================================
' CONTROLES NECESARIOS:
'   lblMoto         - Label "Moto:"
'   cboMoto         - ComboBox
'   lblEstadoActual - Label "Estado actual: ..."
'   lblNuevoEstado  - Label "Nuevo estado:"
'   cboEstado       - ComboBox
'   btnCambiar      - CommandButton "Cambiar Estado"
'   btnCancelar     - CommandButton "Cancelar"
' ============================================================

Option Explicit

Private Sub UserForm_Initialize()
    ' Cargar motos no vendidas
    Dim wsMotos As Worksheet
    Dim wsEstado As Worksheet
    Set wsMotos = ThisWorkbook.Sheets(HOJA_MOTOS)
    Set wsEstado = ThisWorkbook.Sheets(HOJA_ESTADO)

    Dim ultimaFila As Long
    ultimaFila = ObtenerUltimaFilaConDatos(wsMotos, "A", 4)

    Dim i As Long
    For i = 4 To ultimaFila
        Dim estado As String
        estado = ""
        Dim j As Long
        For j = 4 To ObtenerUltimaFilaConDatos(wsEstado, "A", 4)
            If wsEstado.Cells(j, 1).Value = wsMotos.Cells(i, 1).Value Then
                estado = wsEstado.Cells(j, 3).Value
                Exit For
            End If
        Next j

        If estado <> "Vendida" Then
            cboMoto.AddItem wsMotos.Cells(i, 1).Value & " - " & _
                           wsMotos.Cells(i, 2).Value & " " & wsMotos.Cells(i, 3).Value
        End If
    Next i

    ' Cargar estados
    cboEstado.AddItem "Desarmado"
    cboEstado.AddItem "Esperando repuestos"
    cboEstado.AddItem "En taller"
    cboEstado.AddItem "Lista para venta"

    Me.BackColor = RGB(245, 245, 245)
End Sub

Private Sub cboMoto_Change()
    If cboMoto.ListIndex = -1 Then Exit Sub

    Dim idMoto As String
    idMoto = Left(cboMoto.Value, InStr(cboMoto.Value, " - ") - 1)

    ' Buscar estado actual
    Dim wsEstado As Worksheet
    Set wsEstado = ThisWorkbook.Sheets(HOJA_ESTADO)

    Dim fila As Long
    For fila = 4 To ObtenerUltimaFilaConDatos(wsEstado, "A", 4)
        If wsEstado.Cells(fila, 1).Value = idMoto Then
            lblEstadoActual.Caption = "Estado actual: " & wsEstado.Cells(fila, 3).Value
            Exit For
        End If
    Next fila
End Sub

Private Sub btnCambiar_Click()
    If cboMoto.ListIndex = -1 Then
        MsgBox "Seleccione una moto.", vbExclamation
        Exit Sub
    End If

    If cboEstado.ListIndex = -1 Then
        MsgBox "Seleccione el nuevo estado.", vbExclamation
        Exit Sub
    End If

    Dim idMoto As String
    idMoto = Left(cboMoto.Value, InStr(cboMoto.Value, " - ") - 1)

    CambiarEstadoMoto idMoto, cboEstado.Value

    ' Actualizar label
    lblEstadoActual.Caption = "Estado actual: " & cboEstado.Value

    MsgBox "Estado cambiado exitosamente.", vbInformation
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub
