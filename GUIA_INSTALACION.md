# Herramienta de Gestion - Flipping de Motos

## Guia de Instalacion Paso a Paso

### Paso 1: Crear el archivo Excel

1. Abrir Excel
2. Guardar como **"Flipping_Motos.xlsm"** (formato: Libro de Excel habilitado para macros)
3. Habilitar macros cuando Excel lo solicite

### Paso 2: Abrir el Editor de VBA

1. Presionar **Alt + F11** para abrir el Editor de Visual Basic
2. En el panel izquierdo veras el proyecto "VBAProject (Flipping_Motos.xlsm)"

### Paso 3: Importar los Modulos (.bas)

Para cada archivo `.bas` en la carpeta `vba_modules/`:

1. En el Editor de VBA: **Archivo > Importar archivo...**
2. Seleccionar cada archivo `.bas`:
   - `M01_ConfiguracionInicial.bas`
   - `M02_FormulasYCalculos.bas`
   - `M03_Formularios.bas`
   - `M04_MenuPrincipal.bas`

Alternativamente, puedes copiar y pegar el contenido de cada archivo en modulos nuevos:
- Click derecho en el proyecto > **Insertar > Modulo**
- Pegar el codigo del archivo correspondiente

### Paso 4: Crear los UserForms

Para cada formulario, seguir estos pasos:

#### 4.1 - frmAltaMoto (Registrar Nueva Moto)

1. Click derecho en el proyecto > **Insertar > UserForm**
2. Cambiar la propiedad **Name** a `frmAltaMoto`
3. Cambiar **Caption** a `Registrar Nueva Moto`
4. Agregar los siguientes controles (usar el Cuadro de herramientas):

| Control    | Name           | Propiedad Caption/Text     | Posicion  |
|------------|----------------|----------------------------|-----------|
| Label      | lblMarca       | Marca:                     | Fila 1    |
| TextBox    | txtMarca       |                            | Fila 1    |
| Label      | lblModelo      | Modelo:                    | Fila 2    |
| TextBox    | txtModelo      |                            | Fila 2    |
| Label      | lblAnio        | Anio:                      | Fila 3    |
| TextBox    | txtAnio        |                            | Fila 3    |
| Label      | lblCilindrada  | Cilindrada:                | Fila 4    |
| TextBox    | txtCilindrada  |                            | Fila 4    |
| Label      | lblPrecio      | Precio de Compra ($):      | Fila 5    |
| TextBox    | txtPrecio      |                            | Fila 5    |
| Label      | lblFecha       | Fecha de Compra:           | Fila 6    |
| TextBox    | txtFecha       |                            | Fila 6    |
| Label      | lblObs         | Observaciones:             | Fila 7    |
| TextBox    | txtObs         | (MultiLine=True, Height=50)| Fila 7    |
| CommandBtn | btnRegistrar   | Registrar                  | Inferior  |
| CommandBtn | btnCancelar    | Cancelar                   | Inferior  |

5. Copiar el codigo VBA del archivo `vba_forms/frmAltaMoto.frm` en el modulo de codigo del UserForm

#### 4.2 - frmRegistroGasto (Registrar Gasto)

1. Insertar nuevo UserForm, **Name** = `frmRegistroGasto`
2. **Caption** = `Registrar Gasto`
3. Controles:

| Control    | Name           | Propiedad Caption/Text     |
|------------|----------------|----------------------------|
| Label      | lblMoto        | Moto:                      |
| ComboBox   | cboMoto        | (Style=fmStyleDropDownList) |
| Label      | lblFecha       | Fecha:                     |
| TextBox    | txtFecha       |                            |
| Label      | lblCategoria   | Categoria:                 |
| ComboBox   | cboCategoria   | (Style=fmStyleDropDownList) |
| Label      | lblDescripcion | Descripcion:               |
| TextBox    | txtDescripcion |                            |
| Label      | lblMonto       | Monto ($):                 |
| TextBox    | txtMonto       |                            |
| Label      | lblPagadoPor   | Pagado por:                |
| ComboBox   | cboPagadoPor   | (Style=fmStyleDropDownList) |
| Label      | lblComprobante | N° Comprobante:            |
| TextBox    | txtComprobante |                            |
| CommandBtn | btnRegistrar   | Registrar Gasto            |
| CommandBtn | btnCancelar    | Cancelar                   |

4. Copiar el codigo VBA del archivo `vba_forms/frmRegistroGasto.frm`

#### 4.3 - frmRegistrarVenta (Registrar Venta)

1. Insertar UserForm, **Name** = `frmRegistrarVenta`
2. **Caption** = `Registrar Venta de Moto`
3. Controles:

| Control    | Name           | Propiedad Caption/Text     |
|------------|----------------|----------------------------|
| Label      | lblMoto        | Moto a vender:             |
| ComboBox   | cboMoto        | (Style=fmStyleDropDownList) |
| Label      | lblPrecio      | Precio de Venta ($):       |
| TextBox    | txtPrecioVenta |                            |
| Label      | lblFecha       | Fecha de Venta:            |
| TextBox    | txtFechaVenta  |                            |
| Label      | lblResumen     | (vacio, se llena al elegir)|
| CommandBtn | btnVender      | Registrar Venta            |
| CommandBtn | btnCancelar    | Cancelar                   |

4. Copiar el codigo VBA del archivo `vba_forms/frmRegistrarVenta.frm`

#### 4.4 - frmCambiarEstado (Cambiar Estado)

1. Insertar UserForm, **Name** = `frmCambiarEstado`
2. **Caption** = `Cambiar Estado de Moto`
3. Controles:

| Control    | Name             | Propiedad Caption/Text     |
|------------|------------------|----------------------------|
| Label      | lblMoto          | Moto:                      |
| ComboBox   | cboMoto          | (Style=fmStyleDropDownList) |
| Label      | lblEstadoActual  | Estado actual: ...         |
| Label      | lblNuevoEstado   | Nuevo estado:              |
| ComboBox   | cboEstado        | (Style=fmStyleDropDownList) |
| CommandBtn | btnCambiar       | Cambiar Estado             |
| CommandBtn | btnCancelar      | Cancelar                   |

4. Copiar el codigo VBA del archivo `vba_forms/frmCambiarEstado.frm`

### Paso 5: Inicializar el Sistema

1. Cerrar el Editor de VBA (Alt + Q)
2. Presionar **Alt + F8** para ver la lista de macros
3. Seleccionar **`InicializarTodo`** y hacer click en **Ejecutar**
4. Confirmar cuando pregunte si desea crear la estructura
5. El sistema creara automaticamente las 6 hojas con sus formulas y formatos

### Paso 6: Comenzar a Usar

Una vez inicializado, usar los **botones del Dashboard** o el **Menu Principal** (Alt+F8 > MostrarMenuPrincipal).


## Estructura de Hojas

### Hoja 1: Configuracion
Parametros globales del sistema (nombres de socios, porcentajes de reparto).

### Hoja 2: Registro_Motos
| Columna | Campo          | Tipo     | Descripcion                    |
|---------|----------------|----------|--------------------------------|
| A       | ID_Moto        | Texto    | MOTO-001, MOTO-002, etc.      |
| B       | Marca          | Texto    | Honda, Yamaha, etc.            |
| C       | Modelo         | Texto    | CG 150, YBR 125, etc.         |
| D       | Anio           | Numero   | 2018, 2020, etc.               |
| E       | Cilindrada     | Texto    | 150cc, 250cc, etc.             |
| F       | Precio_Compra  | Moneda   | Precio pagado al adquirir      |
| G       | Fecha_Compra   | Fecha    | dd/mm/yyyy                     |
| H       | Precio_Venta   | Moneda   | Se llena al vender             |
| I       | Fecha_Venta    | Fecha    | Se llena al vender             |
| J       | Observaciones  | Texto    | Notas generales                |

### Hoja 3: Gastos
| Columna | Campo          | Tipo     | Descripcion                    |
|---------|----------------|----------|--------------------------------|
| A       | ID_Gasto       | Texto    | GASTO-001, GASTO-002, etc.    |
| B       | ID_Moto        | Texto    | Referencia a la moto           |
| C       | Fecha          | Fecha    | Fecha del gasto                |
| D       | Categoria      | Lista    | Repuestos/Pintura/Gestoria/etc.|
| E       | Descripcion    | Texto    | Detalle del gasto              |
| F       | Monto          | Moneda   | Importe del gasto              |
| G       | Pagado_Por     | Lista    | Socio A / Socio B              |
| H       | Comprobante    | Texto    | Numero de factura/recibo       |

### Hoja 4: Estado_Proyecto
| Columna | Campo              | Tipo     | Descripcion                  |
|---------|--------------------|----------|------------------------------|
| A       | ID_Moto            | Texto    | Referencia                   |
| B       | Moto               | Formula  | Marca + Modelo (auto)        |
| C       | Estado_Actual      | Lista    | Desarmado/Esperando/Taller/Lista/Vendida |
| D       | Fecha_Cambio       | Fecha    | Ultimo cambio de estado      |
| E       | Dias_En_Estado     | Formula  | HOY() - Fecha_Cambio         |
| F       | Dias_Totales_Stock | Formula  | Desde compra hasta venta/hoy |
| G       | Notas              | Texto    | Historial de cambios         |

### Hoja 5: Calculo_Ganancia
| Columna | Campo              | Tipo     | Formula/Descripcion              |
|---------|--------------------|----------|----------------------------------|
| A       | ID_Moto            | Texto    | Referencia                       |
| B       | Moto               | Formula  | =INDEX(Marca, MATCH(ID))+" "+Modelo |
| C       | Precio_Compra      | Formula  | =INDEX(Registro!F, MATCH(ID))    |
| D       | Gastos_Socio_A     | Formula  | =SUMPRODUCT((Gastos!B=ID)*(G=SocioA)*F) |
| E       | Gastos_Socio_B     | Formula  | =SUMPRODUCT((Gastos!B=ID)*(G=SocioB)*F) |
| F       | Total_Gastos       | Formula  | =D+E                            |
| G       | Inversion_Total    | Formula  | =C+F                            |
| H       | Precio_Venta       | Formula  | =INDEX(Registro!H, MATCH(ID))    |
| I       | Utilidad_Neta      | Formula  | =H-G (si H>0)                   |
| J       | Devolucion_Socio_A | Formula  | =(C/2)+D+(I*%A)                  |
| K       | Devolucion_Socio_B | Formula  | =(C/2)+E+(I*%B)                  |
| L       | ROI_%              | Formula  | =I/G                            |

**Panel lateral de reparto global (columnas N-P):**
- Inversion realizada por cada socio
- Devolucion de la inversion
- Participacion en ganancia segun %
- Total a recibir por cada socio

### Hoja 6: Dashboard
KPIs automaticos:
- Total Motos | Motos Activas | Motos Vendidas
- Inversion Total | Ganancia Neta Total | ROI Promedio

Graficos:
1. Inversion acumulada Socio A vs Socio B (barras)
2. ROI por moto (barras)
3. Motos por estado (torta)
4. Flujo de caja (barras)

Datos calculados:
- Dias promedio de stock
- Flujo de caja neto


## Logica de Reparto de Ganancias

El sistema implementa la siguiente logica para cada moto vendida:

```
1. Se calcula la INVERSION TOTAL = Precio Compra + Todos los Gastos
2. Se calcula la UTILIDAD NETA = Precio Venta - Inversion Total
3. El reparto se hace en dos pasos:
   a) DEVOLUCION: Cada socio recupera exactamente lo que invirtio
      - Compra: se divide 50/50 (ajustable)
      - Gastos: cada uno recupera lo que pago
   b) GANANCIA: El excedente se divide segun el % acordado
      - Por defecto: 50% / 50%
      - Configurable en la hoja "Configuracion"
```

### Ejemplo practico:
```
Moto: Honda CG 150
Precio compra: $500,000 (50% cada socio = $250,000)
Gastos Socio A: $80,000 (repuestos)
Gastos Socio B: $120,000 (pintura + gestoria)
Precio venta: $900,000

Inversion total: $500,000 + $80,000 + $120,000 = $700,000
Utilidad neta: $900,000 - $700,000 = $200,000

Reparto Socio A: $250,000 + $80,000 + ($200,000 * 50%) = $430,000
Reparto Socio B: $250,000 + $120,000 + ($200,000 * 50%) = $470,000
```


## Macros Disponibles (Alt+F8)

| Macro                          | Descripcion                                |
|--------------------------------|--------------------------------------------|
| `InicializarTodo`              | Crea toda la estructura (ejecutar 1 vez)   |
| `MostrarMenuPrincipal`         | Muestra el menu con todas las opciones     |
| `MostrarFormularioAltaMoto`    | Abre formulario para registrar nueva moto  |
| `MostrarFormularioGasto`       | Abre formulario para registrar un gasto    |
| `MostrarFormularioEstado`      | Abre formulario para cambiar estado        |
| `MostrarFormularioVenta`       | Abre formulario para registrar una venta   |
| `ActualizarDashboard`          | Regenera los graficos del Dashboard        |
| `MostrarResumenRapido`         | Muestra un popup con el resumen del negocio|
| `CrearBotonesRapidos`          | Crea los botones en el Dashboard           |
