# Referencia de Formulas Excel

## Hoja: Estado_Proyecto

### Dias en Estado Actual (Columna E)
```excel
=IF(D4<>"",TODAY()-D4,"")
```
Calcula los dias transcurridos desde el ultimo cambio de estado.

### Dias Totales de Stock (Columna F)
```excel
=IFERROR(
  IF(INDEX(Registro_Motos!I4:I503,MATCH(A4,Registro_Motos!A4:A503,0))<>"",
    INDEX(Registro_Motos!I4:I503,MATCH(A4,Registro_Motos!A4:A503,0))
      -INDEX(Registro_Motos!G4:G503,MATCH(A4,Registro_Motos!A4:A503,0)),
    TODAY()-INDEX(Registro_Motos!G4:G503,MATCH(A4,Registro_Motos!A4:A503,0))
  ),0)
```
Si la moto fue vendida: Fecha_Venta - Fecha_Compra.
Si no fue vendida: HOY() - Fecha_Compra.


## Hoja: Calculo_Ganancia

### Moto - Descripcion (Columna B)
```excel
=IFERROR(
  INDEX(Registro_Motos!B4:B503,MATCH(A4,Registro_Motos!A4:A503,0))
  &" "&
  INDEX(Registro_Motos!C4:C503,MATCH(A4,Registro_Motos!A4:A503,0)),
"")
```

### Precio de Compra (Columna C)
```excel
=IFERROR(INDEX(Registro_Motos!F4:F503,MATCH(A4,Registro_Motos!A4:A503,0)),0)
```

### Gastos Socio A (Columna D)
```excel
=SUMPRODUCT(
  (Gastos!B4:B503=A4)*
  (Gastos!G4:G503=NombreSocioA)*
  (Gastos!F4:F503)
)
```
Suma todos los gastos de esta moto pagados por Socio A.

### Gastos Socio B (Columna E)
```excel
=SUMPRODUCT(
  (Gastos!B4:B503=A4)*
  (Gastos!G4:G503=NombreSocioB)*
  (Gastos!F4:F503)
)
```

### Total Gastos (Columna F)
```excel
=D4+E4
```

### Inversion Total (Columna G)
```excel
=C4+F4
```

### Precio de Venta (Columna H)
```excel
=IFERROR(INDEX(Registro_Motos!H4:H503,MATCH(A4,Registro_Motos!A4:A503,0)),0)
```

### Utilidad Neta (Columna I)
```excel
=IF(H4>0, H4-G4, 0)
```
Solo se calcula si hay precio de venta registrado.

### Devolucion Socio A (Columna J)
```excel
=IF(H4>0,
  (C4/2) + D4 + IF(I4>0, I4*PctSocioA, I4*0.5),
0)
```
Logica:
1. Recupera la mitad del precio de compra
2. Recupera todos sus gastos
3. Recibe su % de la ganancia (o absorbe % de la perdida)

### Devolucion Socio B (Columna K)
```excel
=IF(H4>0,
  (C4/2) + E4 + IF(I4>0, I4*PctSocioB, I4*0.5),
0)
```

### ROI - Retorno de Inversion (Columna L)
```excel
=IFERROR(I4/G4, 0)
```


## Hoja: Calculo_Ganancia - Panel de Reparto Global (Columnas N-P)

### Inversion Realizada - Socio A (O5)
```excel
=SUMPRODUCT((Gastos!G4:G503=NombreSocioA)*(Gastos!F4:F503))
```

### Inversion Realizada - Socio B (P5)
```excel
=SUMPRODUCT((Gastos!G4:G503=NombreSocioB)*(Gastos!F4:F503))
```

### Participacion en Ganancia - Socio A (O7)
```excel
=SUM(I4:I503)*PctSocioA
```

### Total a Recibir - Socio A (O8)
```excel
=O6+O7
```
Devolucion de inversion + participacion en ganancia.


## Hoja: Dashboard - KPIs

### Total Motos (A4)
```excel
=COUNTA(Registro_Motos!A4:A503)
```

### Motos Activas (C4)
```excel
=COUNTA(Registro_Motos!A4:A503)-COUNTIF(Estado_Proyecto!C4:C503,"Vendida")
```

### Inversion Total (G4)
```excel
=SUM(Registro_Motos!F4:F503)+SUM(Gastos!F4:F503)
```

### Ganancia Neta Total (I4)
```excel
=SUM(Calculo_Ganancia!I4:I503)
```

### ROI Promedio (K4)
```excel
=IFERROR(AVERAGE(Calculo_Ganancia!L4:L503),0)
```

### Flujo de Caja Neto (E10)
```excel
=SUM(Registro_Motos!H4:H503) - SUM(Registro_Motos!F4:F503) - SUM(Gastos!F4:F503)
```
Ingresos por ventas - Compras - Gastos.

### Dias Promedio de Stock (A11)
```excel
=IFERROR(AVERAGE(Estado_Proyecto!F4:F503),0)
```

### Conteo por Estado (I7:I11)
```excel
=COUNTIF(Estado_Proyecto!C4:C503,"Desarmado")
=COUNTIF(Estado_Proyecto!C4:C503,"Esperando repuestos")
=COUNTIF(Estado_Proyecto!C4:C503,"En taller")
=COUNTIF(Estado_Proyecto!C4:C503,"Lista para venta")
=COUNTIF(Estado_Proyecto!C4:C503,"Vendida")
```


## Nombres Definidos (Name Manager)

| Nombre         | Referencia                    | Uso                          |
|----------------|-------------------------------|------------------------------|
| ListaEstados   | Configuracion!$D$4:$D$8      | Validacion de datos          |
| ListaCategorias| Configuracion!$F$4:$F$10     | Validacion de datos          |
| PctSocioA      | Configuracion!$B$6           | % ganancia Socio A           |
| PctSocioB      | Configuracion!$B$7           | % ganancia Socio B           |
| NombreSocioA   | Configuracion!$B$3           | Nombre del Socio A           |
| NombreSocioB   | Configuracion!$B$4           | Nombre del Socio B           |
