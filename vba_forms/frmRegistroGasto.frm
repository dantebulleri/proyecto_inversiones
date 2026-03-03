VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegistroGasto
   Caption         =   "Registrar Gasto"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "frmRegistroGasto.frx":0000
   StartUpPosition =   1  'CenterOwner
End

' ============================================================
' USERFORM: frmRegistroGasto
' Formulario para registrar gastos asociados a una moto
' ============================================================
' CONTROLES NECESARIOS:
'   lblMoto         - Label "Moto:"
'   cboMoto         - ComboBox (lista de motos)
'   lblFecha        - Label "Fecha:"
'   txtFecha        - TextBox
'   lblCategoria    - Label "Categoria:"
'   cboCategoria    - ComboBox
'   lblDescripcion  - Label "Descripcion:"
'   txtDescripcion  - TextBox
'   lblMonto        - Label "Monto ($):"
'   txtMonto        - TextBox
'   lblPagadoPor    - Label "Pagado por:"
'   cboPagadoPor    - ComboBox
'   lblComprobante  - Label "N° Comprobante:"
'   txtComprobante  - TextBox
'   btnRegistrar    - CommandButton "Registrar Gasto"
'   btnCancelar     - CommandButton "Cancelar"
' ============================================================

Option Explicit

Private Sub UserForm_Initialize()
    ' Cargar lista de motos
    Dim motos() As String
    motos = ObtenerListaMotos()

    Dim i As Long
    For i = LBound(motos) To UBound(motos)
        cboMoto.AddItem motos(i)
    Next i

    ' Cargar categorias
    cboCategoria.AddItem "Repuestos"
    cboCategoria.AddItem "Mano de obra"
    cboCategoria.AddItem "Pintura"
    cboCategoria.AddItem "Gestoria"
    cboCategoria.AddItem "Transporte"
    cboCategoria.AddItem "Seguro"
    cboCategoria.AddItem "Otros"

    ' Cargar socios
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(HOJA_CONFIG)
    cboPagadoPor.AddItem wsConfig.Range("B3").Value  ' Socio A
    cboPagadoPor.AddItem wsConfig.Range("B4").Value  ' Socio B

    ' Fecha por defecto
    txtFecha.Value = Format(Date, "dd/mm/yyyy")

    Me.BackColor = RGB(245, 245, 245)
End Sub

Private Sub btnRegistrar_Click()
    ' Validaciones
    If cboMoto.ListIndex = -1 Then
        MsgBox "Seleccione una moto.", vbExclamation
        cboMoto.SetFocus
        Exit Sub
    End If

    If Not IsDate(txtFecha.Value) Then
        MsgBox "Ingrese una fecha valida.", vbExclamation
        txtFecha.SetFocus
        Exit Sub
    End If

    If cboCategoria.ListIndex = -1 Then
        MsgBox "Seleccione una categoria.", vbExclamation
        cboCategoria.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtMonto.Value) Or CDbl(txtMonto.Value) <= 0 Then
        MsgBox "Ingrese un monto valido.", vbExclamation
        txtMonto.SetFocus
        Exit Sub
    End If

    If cboPagadoPor.ListIndex = -1 Then
        MsgBox "Seleccione quien pago.", vbExclamation
        cboPagadoPor.SetFocus
        Exit Sub
    End If

    ' Extraer solo el ID de la moto (antes del " - ")
    Dim idMoto As String
    idMoto = Left(cboMoto.Value, InStr(cboMoto.Value, " - ") - 1)

    RegistrarGasto _
        idMoto:=idMoto, _
        fecha:=CDate(txtFecha.Value), _
        categoria:=cboCategoria.Value, _
        descripcion:=Trim(txtDescripcion.Value), _
        monto:=CDbl(txtMonto.Value), _
        pagadoPor:=cboPagadoPor.Value, _
        comprobante:=Trim(txtComprobante.Value)

    ' Limpiar para otro gasto
    LimpiarFormulario
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub LimpiarFormulario()
    cboCategoria.ListIndex = -1
    txtDescripcion.Value = ""
    txtMonto.Value = ""
    txtComprobante.Value = ""
    txtFecha.Value = Format(Date, "dd/mm/yyyy")
    ' Mantener la moto y el socio seleccionados para carga rapida
End Sub
