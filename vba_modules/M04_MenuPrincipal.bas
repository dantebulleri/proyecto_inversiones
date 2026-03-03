Attribute VB_Name = "M04_MenuPrincipal"
'===============================================================================
' MODULO: M04_MenuPrincipal
' DESCRIPCION: Menu principal y botones de acceso rapido
' VERSION: 1.0
'===============================================================================
Option Explicit

Public Sub MostrarMenuPrincipal()
    '===========================================================================
    ' Muestra el menu principal con las opciones disponibles
    '===========================================================================
    Dim opcion As String
    opcion = InputBox("MENU PRINCIPAL - FLIPPING DE MOTOS" & vbCrLf & vbCrLf & _
                      "Ingrese el numero de la opcion:" & vbCrLf & vbCrLf & _
                      "1 - Registrar nueva moto" & vbCrLf & _
                      "2 - Registrar gasto" & vbCrLf & _
                      "3 - Cambiar estado de moto" & vbCrLf & _
                      "4 - Registrar venta" & vbCrLf & _
                      "5 - Actualizar Dashboard" & vbCrLf & _
                      "6 - Resumen rapido" & vbCrLf & vbCrLf & _
                      "0 - Cancelar", "Menu Principal")

    Select Case opcion
        Case "1": MostrarFormularioAltaMoto
        Case "2": MostrarFormularioGasto
        Case "3": MostrarFormularioEstado
        Case "4": MostrarFormularioVenta
        Case "5": ActualizarDashboard
        Case "6": MostrarResumenRapido
        Case "0", "": ' Cancelar
        Case Else
            MsgBox "Opcion no valida.", vbExclamation
    End Select
End Sub

Public Sub ActualizarDashboard()
    Application.ScreenUpdating = False
    ActualizarGraficos
    ThisWorkbook.Sheets(HOJA_DASHBOARD).Activate
    Application.ScreenUpdating = True
    MsgBox "Dashboard actualizado.", vbInformation
End Sub

Public Sub MostrarResumenRapido()
    '===========================================================================
    ' Muestra un resumen rapido del estado del negocio
    '===========================================================================
    Dim wsMotos As Worksheet
    Dim wsGastos As Worksheet
    Dim wsEstado As Worksheet
    Dim wsConfig As Worksheet

    Set wsMotos = ThisWorkbook.Sheets(HOJA_MOTOS)
    Set wsGastos = ThisWorkbook.Sheets(HOJA_GASTOS)
    Set wsEstado = ThisWorkbook.Sheets(HOJA_ESTADO)
    Set wsConfig = ThisWorkbook.Sheets(HOJA_CONFIG)

    Dim totalMotos As Long
    totalMotos = ObtenerUltimaFilaConDatos(wsMotos, "A", 4) - 3
    If totalMotos < 0 Then totalMotos = 0

    Dim motosVendidas As Long
    motosVendidas = Application.WorksheetFunction.CountIf( _
        wsEstado.Range("C4:C503"), "Vendida")

    Dim inversionTotal As Double
    inversionTotal = Application.WorksheetFunction.Sum(wsMotos.Range("F4:F503")) + _
                     Application.WorksheetFunction.Sum(wsGastos.Range("F4:F503"))

    Dim ventasTotal As Double
    ventasTotal = Application.WorksheetFunction.Sum(wsMotos.Range("H4:H503"))

    Dim nombreA As String: nombreA = wsConfig.Range("B3").Value
    Dim nombreB As String: nombreB = wsConfig.Range("B4").Value

    Dim invA As Double
    invA = Application.WorksheetFunction.SumProduct( _
        Application.WorksheetFunction.Evaluate("(" & HOJA_GASTOS & "!G4:G503=""" & nombreA & """)*(" & HOJA_GASTOS & "!F4:F503)"))

    Dim invB As Double
    invB = Application.WorksheetFunction.SumProduct( _
        Application.WorksheetFunction.Evaluate("(" & HOJA_GASTOS & "!G4:G503=""" & nombreB & """)*(" & HOJA_GASTOS & "!F4:F503)"))

    MsgBox "RESUMEN RAPIDO DEL NEGOCIO" & vbCrLf & _
           String(40, "-") & vbCrLf & vbCrLf & _
           "Total de motos: " & totalMotos & vbCrLf & _
           "Motos activas: " & (totalMotos - motosVendidas) & vbCrLf & _
           "Motos vendidas: " & motosVendidas & vbCrLf & vbCrLf & _
           "Inversion total: $" & Format(inversionTotal, "#,##0.00") & vbCrLf & _
           "Ventas total: $" & Format(ventasTotal, "#,##0.00") & vbCrLf & _
           "Flujo neto: $" & Format(ventasTotal - inversionTotal, "#,##0.00") & vbCrLf & vbCrLf & _
           "Inversion " & nombreA & ": $" & Format(invA, "#,##0.00") & vbCrLf & _
           "Inversion " & nombreB & ": $" & Format(invB, "#,##0.00"), _
           vbInformation, "Resumen Rapido"
End Sub

Public Sub CrearBotonesRapidos()
    '===========================================================================
    ' Crea botones de acceso rapido en el Dashboard
    '===========================================================================
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(HOJA_DASHBOARD)

    ' Eliminar botones existentes
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then shp.Delete
    Next shp

    Dim btn As Shape
    Dim topPos As Single: topPos = 120

    ' Boton: Nueva Moto
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, 10, topPos, 140, 30)
    btn.OnAction = "MostrarFormularioAltaMoto"
    btn.TextFrame.Characters.Text = "Nueva Moto"
    btn.TextFrame.Characters.Font.Size = 10
    btn.TextFrame.Characters.Font.Bold = True

    ' Boton: Nuevo Gasto
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, 160, topPos, 140, 30)
    btn.OnAction = "MostrarFormularioGasto"
    btn.TextFrame.Characters.Text = "Nuevo Gasto"
    btn.TextFrame.Characters.Font.Size = 10
    btn.TextFrame.Characters.Font.Bold = True

    ' Boton: Cambiar Estado
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, 310, topPos, 140, 30)
    btn.OnAction = "MostrarFormularioEstado"
    btn.TextFrame.Characters.Text = "Cambiar Estado"
    btn.TextFrame.Characters.Font.Size = 10
    btn.TextFrame.Characters.Font.Bold = True

    ' Boton: Registrar Venta
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, 460, topPos, 140, 30)
    btn.OnAction = "MostrarFormularioVenta"
    btn.TextFrame.Characters.Text = "Registrar Venta"
    btn.TextFrame.Characters.Font.Size = 10
    btn.TextFrame.Characters.Font.Bold = True

    ' Boton: Actualizar Dashboard
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, 610, topPos, 160, 30)
    btn.OnAction = "ActualizarDashboard"
    btn.TextFrame.Characters.Text = "Actualizar Dashboard"
    btn.TextFrame.Characters.Font.Size = 10
    btn.TextFrame.Characters.Font.Bold = True

    ' Boton: Menu Principal
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, 780, topPos, 140, 30)
    btn.OnAction = "MostrarMenuPrincipal"
    btn.TextFrame.Characters.Text = "Menu Principal"
    btn.TextFrame.Characters.Font.Size = 10
    btn.TextFrame.Characters.Font.Bold = True
End Sub

Public Sub InicializarTodo()
    '===========================================================================
    ' PROCEDIMIENTO MAESTRO: Ejecutar una sola vez para configurar todo
    '===========================================================================
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("Esto creara toda la estructura del libro." & vbCrLf & _
                       "Si ya tiene datos, se perderan." & vbCrLf & vbCrLf & _
                       "Desea continuar?", vbYesNo + vbExclamation, "Inicializar")

    If respuesta = vbNo Then Exit Sub

    ' 1. Crear estructura
    CrearEstructuraCompleta

    ' 2. Crear botones en Dashboard
    CrearBotonesRapidos

    ' 3. Activar Dashboard
    ThisWorkbook.Sheets(HOJA_DASHBOARD).Activate

    MsgBox "Sistema inicializado correctamente." & vbCrLf & vbCrLf & _
           "Use los botones del Dashboard o el Menu Principal para comenzar.", _
           vbInformation, "Listo!"
End Sub
