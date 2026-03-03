VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAltaMoto
   Caption         =   "Registrar Nueva Moto"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "frmAltaMoto.frx":0000
   StartUpPosition =   1  'CenterOwner
End

' ============================================================
' USERFORM: frmAltaMoto
' Formulario para dar de alta una nueva moto
' ============================================================
' CONTROLES NECESARIOS (crear en el Editor de VBA):
'   lblMarca        - Label "Marca:"
'   txtMarca        - TextBox
'   lblModelo       - Label "Modelo:"
'   txtModelo       - TextBox
'   lblAnio         - Label "Anio:"
'   txtAnio         - TextBox
'   lblCilindrada   - Label "Cilindrada:"
'   txtCilindrada   - TextBox
'   lblPrecio       - Label "Precio de Compra ($):"
'   txtPrecio       - TextBox
'   lblFecha        - Label "Fecha de Compra:"
'   txtFecha        - TextBox
'   lblObs          - Label "Observaciones:"
'   txtObs          - TextBox (MultiLine=True)
'   btnRegistrar    - CommandButton "Registrar"
'   btnCancelar     - CommandButton "Cancelar"
' ============================================================

Option Explicit

Private Sub UserForm_Initialize()
    ' Fecha por defecto = hoy
    txtFecha.Value = Format(Date, "dd/mm/yyyy")
    txtMarca.SetFocus

    ' Estilo del formulario
    Me.BackColor = RGB(245, 245, 245)
End Sub

Private Sub btnRegistrar_Click()
    ' Validaciones
    If Trim(txtMarca.Value) = "" Then
        MsgBox "Ingrese la marca de la moto.", vbExclamation
        txtMarca.SetFocus
        Exit Sub
    End If

    If Trim(txtModelo.Value) = "" Then
        MsgBox "Ingrese el modelo de la moto.", vbExclamation
        txtModelo.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtAnio.Value) Or Len(txtAnio.Value) <> 4 Then
        MsgBox "Ingrese un anio valido (4 digitos).", vbExclamation
        txtAnio.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtPrecio.Value) Or CDbl(txtPrecio.Value) <= 0 Then
        MsgBox "Ingrese un precio de compra valido.", vbExclamation
        txtPrecio.SetFocus
        Exit Sub
    End If

    If Not IsDate(txtFecha.Value) Then
        MsgBox "Ingrese una fecha valida (dd/mm/yyyy).", vbExclamation
        txtFecha.SetFocus
        Exit Sub
    End If

    ' Registrar la moto
    RegistrarMoto _
        marca:=Trim(txtMarca.Value), _
        modelo:=Trim(txtModelo.Value), _
        anio:=Trim(txtAnio.Value), _
        cilindrada:=Trim(txtCilindrada.Value), _
        precioCompra:=CDbl(txtPrecio.Value), _
        fechaCompra:=CDate(txtFecha.Value), _
        observaciones:=Trim(txtObs.Value)

    ' Limpiar formulario para otro ingreso
    LimpiarFormulario
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub LimpiarFormulario()
    txtMarca.Value = ""
    txtModelo.Value = ""
    txtAnio.Value = ""
    txtCilindrada.Value = ""
    txtPrecio.Value = ""
    txtFecha.Value = Format(Date, "dd/mm/yyyy")
    txtObs.Value = ""
    txtMarca.SetFocus
End Sub
