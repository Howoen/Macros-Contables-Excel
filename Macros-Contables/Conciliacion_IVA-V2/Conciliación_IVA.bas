Attribute VB_Name = "Módulo1"
Sub ConciliacionIva()
    ' Variables
    Dim Macro As Workbook
    Dim Macrobasedetados As Worksheet
    Dim Contable As Worksheet
    Dim Ivagenerado As Worksheet
    Dim Ivacompras As Worksheet
    Dim archivotoken As Variant
    Dim wsToken As Worksheet
    Dim wsCompras As Worksheet
    Dim wsVentas As Worksheet
    Dim wsComprasToken As Worksheet
    Dim wsbasededatos As Worksheet
    Dim wsVentasToken As Worksheet
    Dim archivoSeleccionado As Variant
    Dim NombreLibro As String
    Dim Ruta As String
    Dim ultimaFila As Long
    Dim i As Long
    Dim criterio1 As Variant
    Dim criterio2 As String
    Dim NuevoLibro As Workbook
    Dim criterioRango As Range
    Dim listaCriterios As Range
    Dim ultimaFilaDatos As Long
    Dim ultimaColumnaCriterios As Long
    Dim ultimacolumnacridebitos As Long
    Dim primeracolumnacricreditos As Long
    Dim ultimaColumnaCriterioscredito As Long
    Dim ultimafilatkc As Long
    Dim Cantidadcuentas As Long
    Dim UltimaFiladatoscuenta As Long
    Dim sumadeivasdebito As Long
    Dim sumadeivascredito As Long
    Dim inicio As Range
    Dim fin As Range
    Dim descripcionInicio As String
    Dim descripcionFin As String
    Dim rango As Range
    Dim Rangobusv As Range
    Dim Flasev As String
    Dim columnnum As Long
    Dim columnbuscav As Long
    Dim columnbuscavv As Long
    Dim columnbuscavt As Long
    Dim columnbuscavtv As Long
    Dim Columnes As Long
    Dim Columnabusvts As String
    Dim Ultimafilabuscarv As Long
    Dim g As Long
    Dim j As Long
    Dim z As Long
    Dim y As Long
    Dim h As Long
    Dim n As Long
    Dim p As String
    Dim ultimafilah As Long
    
    
    
      
       
    ' Definir y abrir hoja base de datos
    
    Worksheets("BASE DE DATOS").Activate
    
    Set Macrobasedetados = ActiveSheet
    
    ' Solicitar al usuario el nombre del libro
    NombreLibro = InputBox("Por favor, ingresa el mes a conciliar:", "Nombre del Libro")
    If NombreLibro = "" Then
        MsgBox "No se ingresó un nombre para el libro. La operación se cancelará."
        Exit Sub
    End If
    
    ' Solicitar la ubicación para guardar el libro
    Ruta = Application.GetSaveAsFilename(InitialFileName:=NombreLibro, FileFilter:="Archivos de Excel (*.xlsx), *.xlsx")
    If Ruta = "False" Then
        MsgBox "No se seleccionó una ubicación válida. La operación se cancelará."
        Exit Sub
    End If
    
    ' Crear un nuevo libro
    Set NuevoLibro = Workbooks.Add
    NuevoLibro.SaveAs Filename:=Ruta
    MsgBox "El nuevo libro ha sido creado y guardado en: " & Ruta
    
    ' Abrir el archivo del software contable
    archivoSeleccionado = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Seleccione el archivo del software contable")
    If archivoSeleccionado = False Then
        MsgBox "No se seleccionó ningún archivo. La operación ha sido cancelada."
        Exit Sub
    End If
    Workbooks.Open archivoSeleccionado
    Set Contable = ActiveSheet
    
    ' Eliminar las dos primeras filas
    Contable.Rows("1:2").Delete
    
    ' Convertir cuentas contables a número
    ultimaFila = Contable.Cells(Contable.Rows.Count, "F").End(xlUp).Row
    For i = 1 To ultimaFila
        Contable.Cells(i, 11).Formula = "=VALUE(F" & i & ")"
    Next i
    Contable.Range("K:K").Copy
    Contable.Range("F:F").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' Aplicar filtro y copiar datos a una nueva hoja llamada Iva Compras Contable
    criterio1 = "24080200"
    criterio2 = "24080299"

    ' Crear una nueva hoja para los datos filtrados
    Set Ivacompras = NuevoLibro.Sheets.Add(After:=NuevoLibro.Sheets(NuevoLibro.Sheets.Count))
    Ivacompras.Name = "Iva Compras Con"
    
    ' Aplicar el filtro
    With Contable
        .Range("A:H").AutoFilter Field:=6, Criteria1:=">" & criterio1, Operator:=xlAnd, Criteria2:="<" & criterio2
        
        ' Copiar columnas filtradas individualmente
        .Range("B:B").SpecialCells(xlCellTypeVisible).Copy Destination:=Ivacompras.Range("A1") ' Columna B a A
        .Range("C:C").SpecialCells(xlCellTypeVisible).Copy Destination:=Ivacompras.Range("B1") ' Columna C a B
        .Range("E:E").SpecialCells(xlCellTypeVisible).Copy Destination:=Ivacompras.Range("C1") ' Columna E a C
        .Range("F:F").SpecialCells(xlCellTypeVisible).Copy Destination:=Ivacompras.Range("D1") ' Columna F a D
        .Range("G:G").SpecialCells(xlCellTypeVisible).Copy Destination:=Ivacompras.Range("E1") ' Columna G a E
        .Range("H:H").SpecialCells(xlCellTypeVisible).Copy Destination:=Ivacompras.Range("F1") ' Columna H a F
        
        ' Desactivar el filtro
        .AutoFilterMode = False
    End With
    
    ' Determinar la última fila con datos en la columna C de Ivacompras
    ultimaFila = Ivacompras.Cells(Ivacompras.Rows.Count, "C").End(xlUp).Row

    ' Copiar los datos únicos de la columna A de Ivacompras a la columna H y eliminar duplicados
    Ivacompras.Range("C2:C" & ultimaFila).Copy
    Ivacompras.Range("H2").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Ivacompras.Range("H:H").RemoveDuplicates Columns:=1, Header:=xlNo
    
    ' Determinar la última fila con datos en la columna D de Ivacompras
    ultimaFila = Ivacompras.Cells(Ivacompras.Rows.Count, "D").End(xlUp).Row

    ' Copiar los datos únicos de la columna A de Ivacompras a la columna H y eliminar duplicados
    Ivacompras.Range("D2:D" & ultimaFila).Copy
    Ivacompras.Range("I2").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Ivacompras.Range("I:I").RemoveDuplicates Columns:=1, Header:=xlNo
          
   ' Determinar la última fila de la columna I con datos únicos
    ultimaFila = Ivacompras.Cells(Ivacompras.Rows.Count, "I").End(xlUp).Row

    ' Transponer los criterios a la fila 1 a partir de la columna K
    Set listaCriterios = Ivacompras.Range("I2:I" & ultimaFila)
    listaCriterios.Copy
    Ivacompras.Range("K1").PasteSpecial Paste:=xlPasteValues, Transpose:=True
    Application.CutCopyMode = False
    
    
    ' Determinar la última fila de datos en la columna H
    ultimaFilaDatos = Ivacompras.Cells(Ivacompras.Rows.Count, "H").End(xlUp).Row

    ' Determinar la última columna con criterios en la fila 1 desde K en adelante
    ultimacolumnacridebitos = Ivacompras.Cells(1, Ivacompras.Columns.Count).End(xlToLeft).Column
    columnbuscav = Ivacompras.Cells(1, Ivacompras.Columns.Count).End(xlToLeft).Column
    
    columnbuscavt = Ivacompras.Cells(1, Ivacompras.Columns.Count).End(xlToLeft).Column
    z = 2
    y = 1
    primeracolumnacricreditos = ultimacolumnacridebitos + z
    sumadeivasdebito = ultimacolumnacridebitos + y

    ' Iterar sobre cada fila de la columna H desde H2 en adelante
    For g = 2 To ultimaFilaDatos
    ' Iterar sobre cada columna con criterios desde K en adelante
    For j = 11 To ultimacolumnacridebitos ' La columna K es la 11 en numeración de Excel
            ' Agregar la fórmula SUMAR.SI.CONJUNTO en la celda correspondiente
            Ivacompras.Cells(g, j).Formula = _
                "=SUMIFS(E:E, D:D, " & Ivacompras.Cells(1, j).Address & ", C:C, " & Ivacompras.Cells(g, "H").Address & ")"
        Next j
    Next g
    
   ' Determinar la última fila con datos de SUMIFS y agregar la fórmula SUM
    For j = 11 To ultimacolumnacridebitos ' Iterar sobre las columnas de K en adelante
    UltimaFiladatoscuenta = Ivacompras.Cells(Ivacompras.Rows.Count, j).End(xlUp).Row ' Última fila con datos en la columna actual
    Ivacompras.Cells(UltimaFiladatoscuenta + 2, j).Formula = _
        "=SUM(" & Ivacompras.Cells(2, j).Address & ":" & Ivacompras.Cells(UltimaFiladatoscuenta, j).Address & ")"
    Next j
     
    
    ' Transponer los criterios a la fila 1 a partir de la Ultima columna + 1
    Set listaCriterios = Ivacompras.Range("I2:I" & ultimaFila)
    listaCriterios.Copy
    Ivacompras.Cells(1, primeracolumnacricreditos).PasteSpecial Paste:=xlPasteValues, Transpose:=True
    Application.CutCopyMode = False
    
    ' Determinar la última columna con criterios en la fila 1 desde K en adelante
    ultimaColumnaCriterios = Ivacompras.Cells(1, Ivacompras.Columns.Count).End(xlToLeft).Column
    
    ' Iterar sobre cada fila de la columna H desde H2 en adelante
    For g = 2 To ultimaFilaDatos
        ' Iterar sobre cada columna con criterios desde K en adelante
        For j = primeracolumnacricreditos To ultimaColumnaCriterios ' La columna K es la 11 en numeración de Excel
            ' Agregar la fórmula SUMAR.SI.CONJUNTO en la celda correspondiente
            Ivacompras.Cells(g, j).Formula = _
                "=SUMIFS(F:F, D:D, " & Ivacompras.Cells(1, j).Address & ", C:C, " & Ivacompras.Cells(g, "H").Address & ")"
        Next j
    Next g
    
    ' Determinar la última fila con datos de SUMIFS y agregar la fórmula SUM
    For j = primeracolumnacricreditos To ultimaColumnaCriterios ' Iterar sobre las columnas de K en adelante
    UltimaFiladatoscuenta = Ivacompras.Cells(Ivacompras.Rows.Count, j).End(xlUp).Row ' Última fila con datos en la columna actual
    Ivacompras.Cells(UltimaFiladatoscuenta + 2, j).Formula = _
        "=SUM(" & Ivacompras.Cells(2, j).Address & ":" & Ivacompras.Cells(UltimaFiladatoscuenta, j).Address & ")"
    Next j
    
    
    
    
    'suma de ivas debito
          
       
        
        
    
    ' Determinar ultima fila de h
    ultimafilah = Ivacompras.Cells(Ivacompras.Rows.Count, "H").End(xlUp).Row
    For i = 2 To ultimafilah
            ' Agregar la fórmula SUMAR.SI.CONJUNTO en la celda correspondiente
            Ivacompras.Cells(i, sumadeivasdebito).Formula = _
                "=SUM(K" & i & ":" & Cells(i, ultimacolumnacridebitos).Address(False, False) & ")"
        Next i
        
    ' ingresar descripción de columna sumas
        
    Ivacompras.Cells(1, sumadeivasdebito).Value = "Iva Debitos"
    
    ultimaColumnaCriterios = Ivacompras.Cells(1, Ivacompras.Columns.Count).End(xlToLeft).Column
    
    sumadeivascredito = ultimaColumnaCriterios + y
    
    
    
    For i = 2 To ultimafilah
    Ivacompras.Cells(i, sumadeivascredito).Formula = _
        "=SUM(" & Ivacompras.Cells(i, primeracolumnacricreditos).Address(False, False) & ":" & Ivacompras.Cells(i, ultimaColumnaCriterios).Address(False, False) & ")"
        Next i
    Ivacompras.Cells(1, sumadeivascredito).Value = "Iva Creditos"
        
           
       
    ' COMPRAS CONTABLE
        
    ' Define las descripciones a buscar
    descripcionInicio = "Factura de Compra"
    descripcionFin = "Total Factura de Compra"

    ' Busca la descripción inicial en la columna A
    Set inicio = Contable.Columns("A").Find(What:=descripcionInicio, LookIn:=xlValues, LookAt:=xlWhole)

    ' Busca la descripción final en la columna A
    Set fin = Contable.Columns("A").Find(What:=descripcionFin, LookIn:=xlValues, LookAt:=xlWhole)

    ' Verifica si se encontraron ambas descripciones
    If Not inicio Is Nothing And Not fin Is Nothing Then
    ' Define el rango utilizando las coordenadas encontradas
        Set rango = Contable.Range(Contable.Cells(inicio.Row, 2), Contable.Cells(fin.Row, 8)) ' Columnas B a H

    ' Crea una nueva hoja llamada "compras" en el libro nuevo
        Set wsCompras = NuevoLibro.Sheets.Add(After:=NuevoLibro.Sheets(NuevoLibro.Sheets.Count))
        wsCompras.Name = "compras"
    
    ' Copia y pega el rango en la nueva hoja
        rango.Copy Destination:=wsCompras.Range("A2")
        
        'Determinar ultima fila columna d
        ultimafilad = wsCompras.Cells(wsCompras.Rows.Count, "D").End(xlUp).Row

        ' Eliminar los duplicados de la columna D y copiar a la columna J
        wsCompras.Range("D2:D" & ultimafilad).Copy
        wsCompras.Range("J2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        wsCompras.Range("J:J").RemoveDuplicates Columns:=1, Header:=xlNo

        ' Agregar la fórmula SUMAR.SI en la columna K
        ultimaFila = wsCompras.Cells(Rows.Count, "J").End(xlUp).Row
        For i = 2 To ultimaFila
            wsCompras.Cells(i, 11).Formula = "=SUMIF(D2:D" & ultimafilad & ", J" & i & ", G2:G" & ultimafilad & ")"
        Next i
        
        ' Agregar la fórmula SUMA en la columna K dos filas debajo de la última celda ocupada en la columna J
        wsCompras.Cells(ultimaFila + 2, 11).Formula = "=SUM(K1:K" & ultimaFila & ")"

        
    Else
        MsgBox "No se encontraron ambas descripciones en la columna A."
        Exit Sub
    End If
    
    
    
           
        
    ' Aplicar filtro y copiar datos a una nueva hoja llamada Iva ventas Contable****** IVA VENTAS
    criterio1 = "24080100"
    criterio2 = "24080199"

    ' Crear una nueva hoja para los datos filtrados
    Set Ivagenerado = NuevoLibro.Sheets.Add(After:=NuevoLibro.Sheets(NuevoLibro.Sheets.Count))
    Ivagenerado.Name = "Iva Ventas Con"
    
    ' Aplicar el filtro
    With Contable
        .Range("A:H").AutoFilter Field:=6, Criteria1:=">" & criterio1, Operator:=xlAnd, Criteria2:="<" & criterio2
        
        ' Copiar columnas filtradas individualmente
        .Range("B:B").SpecialCells(xlCellTypeVisible).Copy Destination:=Ivagenerado.Range("A1") ' Columna B a A
        .Range("C:C").SpecialCells(xlCellTypeVisible).Copy Destination:=Ivagenerado.Range("B1") ' Columna C a B
        .Range("E:E").SpecialCells(xlCellTypeVisible).Copy Destination:=Ivagenerado.Range("C1") ' Columna E a C
        .Range("F:F").SpecialCells(xlCellTypeVisible).Copy Destination:=Ivagenerado.Range("D1") ' Columna F a D
        .Range("G:G").SpecialCells(xlCellTypeVisible).Copy Destination:=Ivagenerado.Range("E1") ' Columna G a E
        .Range("H:H").SpecialCells(xlCellTypeVisible).Copy Destination:=Ivagenerado.Range("F1") ' Columna H a F
        
        ' Desactivar el filtro
        .AutoFilterMode = False
    End With
    
    ' Determinar la última fila con datos en la columna C de Ivagenerado
    ultimaFila = Ivagenerado.Cells(Ivagenerado.Rows.Count, "C").End(xlUp).Row

    ' Copiar los datos únicos de la columna A de Ivacompras a la columna H y eliminar duplicados
    Ivagenerado.Range("C2:C" & ultimaFila).Copy
    Ivagenerado.Range("H2").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Ivagenerado.Range("H:H").RemoveDuplicates Columns:=1, Header:=xlNo
    
    ' Determinar la última fila con datos en la columna D de Ivacompras
    ultimaFila = Ivagenerado.Cells(Ivagenerado.Rows.Count, "D").End(xlUp).Row

    ' Copiar los datos únicos de la columna A de Ivacompras a la columna H y eliminar duplicados
    Ivagenerado.Range("D2:D" & ultimaFila).Copy
    Ivagenerado.Range("I2").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Ivagenerado.Range("I:I").RemoveDuplicates Columns:=1, Header:=xlNo
          
   ' Determinar la última fila de la columna I con datos únicos
    ultimaFila = Ivagenerado.Cells(Ivagenerado.Rows.Count, "I").End(xlUp).Row

    ' Transponer los criterios a la fila 1 a partir de la columna K
    Set listaCriterios = Ivagenerado.Range("I2:I" & ultimaFila)
    listaCriterios.Copy
    Ivagenerado.Range("K1").PasteSpecial Paste:=xlPasteValues, Transpose:=True
    Application.CutCopyMode = False
    
     ' Determinar la última columna con criterios en la fila 1 desde K en adelante
    ultimacolumnacridebitos = Ivagenerado.Cells(1, Ivagenerado.Columns.Count).End(xlToLeft).Column
    
      ' Determinar la última fila de la columna I con datos únicos
    ultimaFila = Ivagenerado.Cells(Ivagenerado.Rows.Count, "I").End(xlUp).Row
    
    
    ' Columna adicional
    z = 2
    primeracolumnacricreditos = ultimacolumnacridebitos + z
    
    ' Transponer los criterios a la fila 1 a partir de la Ultima columna + 1
    Set listaCriterios = Ivagenerado.Range("I2:I" & ultimaFila)
    listaCriterios.Copy
    Ivagenerado.Cells(1, primeracolumnacricreditos).PasteSpecial Paste:=xlPasteValues, Transpose:=True
    Application.CutCopyMode = False
   
    ' Determinar la última fila de datos en la columna H
    ultimaFilaDatos = Ivagenerado.Cells(Ivagenerado.Rows.Count, "H").End(xlUp).Row

    ' Determinar la última columna con criterios en la fila 1 desde K en adelante
    ultimaColumnaCriterios = Ivagenerado.Cells(1, Ivagenerado.Columns.Count).End(xlToLeft).Column
    

    ' Iterar sobre cada fila de la columna H desde H2 en adelante
    For g = 2 To ultimaFilaDatos
        ' Iterar sobre cada columna con criterios desde K en adelante
            For j = 11 To ultimacolumnacridebitos  ' La columna K es la 11 en numeración de Excel
            ' Agregar la fórmula SUMAR.SI.CONJUNTO en la celda correspondiente
            Ivagenerado.Cells(g, j).Formula = _
                "=SUMIFS(E:E, D:D, " & Ivagenerado.Cells(1, j).Address & ", C:C, " & Ivagenerado.Cells(g, "H").Address & ")"
                
                          
        Next j
    Next g
    
        ' Determinar la última fila con datos de SUMIFS y agregar la fórmula SUM
    For j = 11 To ultimacolumnacridebitos ' Iterar sobre las columnas de K en adelante
    UltimaFiladatoscuenta = Ivagenerado.Cells(Ivagenerado.Rows.Count, j).End(xlUp).Row ' Última fila con datos en la columna actual
    Ivagenerado.Cells(UltimaFiladatoscuenta + 2, j).Formula = _
        "=SUM(" & Ivagenerado.Cells(2, j).Address & ":" & Ivagenerado.Cells(UltimaFiladatoscuenta, j).Address & ")"
    Next j
  
    
    
     ' Determinar la última fila de datos en la columna H
    ultimaFilaDatos = Ivagenerado.Cells(Ivagenerado.Rows.Count, "H").End(xlUp).Row

    ' Determinar la última columna con criterios en la fila 1 desde K en adelante
    ultimaColumnaCriterios = Ivagenerado.Cells(1, Ivagenerado.Columns.Count).End(xlToLeft).Column
    columnbuscavv = Ivagenerado.Cells(1, Ivagenerado.Columns.Count).End(xlToLeft).Column
    columnbuscavtv = Ivagenerado.Cells(1, Ivagenerado.Columns.Count).End(xlToLeft).Column
    ' Iterar sobre cada fila de la columna H desde H2 en adelante
    For g = 2 To ultimaFilaDatos
        ' Iterar sobre cada columna con criterios desde K en adelante
        For j = primeracolumnacricreditos To ultimaColumnaCriterios  ' La columna K es la 11 en numeración de Excel
            ' Agregar la fórmula SUMAR.SI.CONJUNTO en la celda correspondiente
            Ivagenerado.Cells(g, j).Formula = _
                "=SUMIFS(F:F, D:D, " & Ivagenerado.Cells(1, j).Address & ", C:C, " & Ivagenerado.Cells(g, "H").Address & ")"
        Next j
    Next g
       
    ' Determinar la última fila con datos de SUMIFS y agregar la fórmula SUM
    For j = primeracolumnacricreditos To ultimaColumnaCriterios ' Iterar sobre las columnas de K en adelante
    UltimaFiladatoscuenta = Ivagenerado.Cells(Ivagenerado.Rows.Count, j).End(xlUp).Row ' Última fila con datos en la columna actual
    Ivagenerado.Cells(UltimaFiladatoscuenta + 2, j).Formula = _
        "=SUM(" & Ivagenerado.Cells(2, j).Address & ":" & Ivagenerado.Cells(UltimaFiladatoscuenta, j).Address & ")"
    Next j
    
    y = 1
    sumadeivasdebito = ultimacolumnacridebitos + y
    
    ' Determinar ultima fila de h
    ultimafilah = Ivagenerado.Cells(Ivagenerado.Rows.Count, "H").End(xlUp).Row
    For i = 2 To ultimafilah
            ' Agregar la fórmula SUMAR.SI.CONJUNTO en la celda correspondiente
            Ivagenerado.Cells(i, sumadeivasdebito).Formula = _
                "=SUM(K" & i & ":" & Cells(i, ultimacolumnacridebitos).Address(False, False) & ")"
        Next i
    
    y = 1
    ' Determinar la ultima comlumna
    ultimaColumnaCriterioscredito = Ivagenerado.Cells(1, Ivagenerado.Columns.Count).End(xlToLeft).Column
    sumadeivascredito = ultimaColumnaCriterioscredito + y
     
    For i = 2 To ultimafilah
    Ivagenerado.Cells(i, sumadeivascredito).Formula = _
        "=SUM(" & Ivagenerado.Cells(i, primeracolumnacricreditos).Address(False, False) & ":" & Ivagenerado.Cells(i, ultimaColumnaCriterioscredito).Address(False, False) & ")"
        Next i
    
     ' Ventas CONTABLE *** * ** * * * ** * ** * * **
      
     
    ' Define las descripciones a buscar para ventas
    descripcionInicio = "Factura de Venta"
    descripcionFin = "Total Factura de Venta"

    ' Busca la descripción inicial en la columna A
    Set inicio = Contable.Columns("A").Find(What:=descripcionInicio, LookIn:=xlValues, LookAt:=xlWhole)

    ' Busca la descripción final en la columna A
    Set fin = Contable.Columns("A").Find(What:=descripcionFin, LookIn:=xlValues, LookAt:=xlWhole)

    ' Verifica si se encontraron ambas descripciones
    If Not inicio Is Nothing And Not fin Is Nothing Then
        ' Define el rango utilizando las coordenadas encontradas
        Set rango = Contable.Range(Contable.Cells(inicio.Row, 2), Contable.Cells(fin.Row, 8)) ' Columnas B a H

        ' Crea una nueva hoja llamada "ventas" en el libro nuevo
        Set wsVentas = NuevoLibro.Sheets.Add(After:=NuevoLibro.Sheets(NuevoLibro.Sheets.Count))
        wsVentas.Name = "ventas"

        ' Copia y pega el rango en la nueva hoja
        rango.Copy Destination:=wsVentas.Range("A2")

        ' Eliminar los duplicados de la columna D y copiar a la columna J
        ultimafilad = wsVentas.Cells(Rows.Count, "D").End(xlUp).Row
        wsVentas.Range("D2:D" & ultimafilad).Copy
        wsVentas.Range("J2").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        wsVentas.Range("J:J").RemoveDuplicates Columns:=1, Header:=xlNo

        ' Agregar la fórmula SUMAR.SI en la columna K
        ultimaFila = wsVentas.Cells(Rows.Count, "J").End(xlUp).Row
        For i = 2 To ultimaFila
            wsVentas.Cells(i, 12).Formula = "=SUMIF(D2:D" & ultimafilad & " , J" & i & ", F2:F" & ultimafilad & ")"
        Next i
        
        ' Agregar la fórmula SUMA en la columna K dos filas debajo de la última celda ocupada en la columna J
        wsVentas.Cells(ultimaFila + 2, 12).Formula = "=SUM(L2:L" & ultimaFila & ")"

    Else
        
        Exit Sub
    End If
    
    
     ' Seleccionar el archivo Token DIAN
    archivotoken = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Seleccione el archivo del Token DIAN")

    If archivotoken = False Then
        MsgBox "No se seleccionó ningún archivo. La operación ha sido cancelada."
        Exit Sub
    End If

    ' Abre el libro Token DIAN
    Workbooks.Open archivotoken
    Set wsToken = ActiveSheet
    
    ' Convertir columnas g e i a numeros
    ' Determinar ultima fila convertir numero
    ultimafilaconvernumerotokeng = wsToken.Cells(wsToken.Rows.Count, "G").End(xlUp).Row
    For i = 2 To ultimafilaconvernumerotokeng
    wsToken.Cells(i, 17).Formula = "=VALUE(G" & i & ")"
    Next i
    wsToken.Range("Q2:Q" & ultimafilaconvernumerotokeng).Copy
    wsToken.Range("G2").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    wsToken.Range("Q2:Q" & ultimafilaconvernumerotokeng).ClearContents
    
    ' Convertir columnas g e i a numeros
    ultimafilaconvernumerotokeni = wsToken.Cells(wsToken.Rows.Count, "i").End(xlUp).Row
    For i = 2 To ultimafilaconvernumerotokeni
    wsToken.Cells(i, 18).Formula = "=VALUE(I" & i & ")"
    Next i
    wsToken.Range("R2:R" & ultimafilaconvernumerotokeni).Copy
    wsToken.Range("I2").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    wsToken.Range("R2:R" & ultimafilaconvernumerotokeng).ClearContents
    
    
    

    ' Aplicar filtro y copiar datos a una nueva hoja llamada "compras_token" en el libro nuevo
    criterio1 = Array("Factura electrónica", "Documento equivalente POS", "Nota de débito electrónica")
    criterio2 = "recibido"

    ' Crear una nueva hoja para los datos filtrados
    Set wsComprasToken = NuevoLibro.Sheets.Add(After:=NuevoLibro.Sheets(NuevoLibro.Sheets.Count))
    wsComprasToken.Name = "compras_token"
    
    On Error Resume Next ' Manejar posibles errores si los criterios no existen en los datos

    ' Aplicar el filtro
    With wsToken
        .Range("A:P").AutoFilter Field:=1, Criteria1:=criterio1, Operator:=xlFilterValues
        .Range("A:P").AutoFilter Field:=16, Criteria1:=criterio2

        ' Copiar los datos filtrados de las columnas H y N
        Set filtroRango = .Range("C:C,G:G,H:H,K:K,N:N").SpecialCells(xlCellTypeVisible)
        filtroRango.Copy Destination:=wsComprasToken.Range("A1")
        .AutoFilterMode = False
    End With
    On Error GoTo 0 ' Restablecer manejo de errores

    ' Copiar los datos de la columna B (filtrada) a la columna E y eliminar duplicados
    wsComprasToken.Range("B:C").Copy
    wsComprasToken.Range("G1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    wsComprasToken.Range("G:H").RemoveDuplicates Columns:=1, Header:=xlNo

    ' Determinar la última fila de datos en las columnas E y A
    ultimaFila = wsComprasToken.Cells(wsComprasToken.Rows.Count, "H").End(xlUp).Row
    ultimaFilarangocrite = wsComprasToken.Cells(wsComprasToken.Rows.Count, "A").End(xlUp).Row

    ' Agregar la fórmula SUMAR.SI en la columna F
    For i = 2 To ultimaFila ' Comienza desde la fila 2 para evitar el encabezado
    wsComprasToken.Cells(i, 9).Formula = _
        "=SUMIF(B2:B" & ultimaFilarangocrite & ", G" & i & ", D2:D" & ultimaFilarangocrite & ")"
        
    Next i
          
      ' Agregar la fórmula SUMAR.SI en la columna F
    For i = 2 To ultimaFila ' Comienza desde la fila 2 para evitar el encabezado
    wsComprasToken.Cells(i, 10).Formula = _
     "=SUMIF(B2:B" & ultimaFilarangocrite & ", G" & i & ", E2:E" & ultimaFilarangocrite & ")"
     Next i
          
    ' Agregar la fórmula SUMA en la columna F dos filas debajo de la última celda ocupada en la columna E
    ultimaFila = wsComprasToken.Cells(Rows.Count, "I").End(xlUp).Row
    wsComprasToken.Cells(ultimaFila + 2, 10).Formula = "=SUM(J2:J" & ultimaFila & ")"
    
      ' Agregar la fórmula SUMA en la columna F dos filas debajo de la última celda ocupada en la columna E
    
    wsComprasToken.Cells(ultimaFila + 2, 9).Formula = "=SUM(I2:I" & ultimaFila & ")"
    
    'Notas credito *************************************

    ' Determinar la ultima fila de la columna a, para pegar la información de notas credito
    ultimaFila = wsComprasToken.Cells(Rows.Count, "A").End(xlUp).Row
    
    wsComprasToken.Cells(ultimaFila + 3, 1).Value = "Notas credito"
    
    
    ' Aplicar filtro y copiar datos a una nueva hoja llamada "compras_token" en el libro nuevo
    criterio1 = "Nota de crédito electrónica"
    criterio2 = "recibido"
    
    On Error Resume Next ' Manejar posibles errores si los criterios no existen en los datos
    ' Aplicar el filtro
    With wsToken
        .Range("A:P").AutoFilter Field:=1, Criteria1:=criterio1, Operator:=xlFilterValues
        .Range("A:P").AutoFilter Field:=16, Criteria1:=criterio2

        ' Copiar los datos filtrados de las columnas H y N
        Set filtroRango = .Range("C:C,G:G,H:H,K:K,N:N").SpecialCells(xlCellTypeVisible)
        filtroRango.Copy Destination:=wsComprasToken.Cells(ultimaFila + 5, 1)
        .AutoFilterMode = False
    End With
    On Error GoTo 0 ' Restablecer manejo de errores
    ' Determinar la ultima fila de notas credito
    ultimaFilaNotasCredito = wsComprasToken.Cells(Rows.Count, "A").End(xlUp).Row
    ' Copiar los datos de la columna B (filtrada) a la columna E y eliminar duplicados
    wsComprasToken.Range("B" & ultimaFila + 5 & ":C" & ultimaFilaNotasCredito).Copy
    wsComprasToken.Cells(ultimaFila + 5, "F").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    wsComprasToken.Range("F" & ultimaFila + 5 & ":G" & ultimaFilaNotasCredito).RemoveDuplicates Columns:=1, Header:=xlNo

    ' Determinar la última fila de datos en la columna F
    ultimafilasumarsi = wsComprasToken.Cells(Rows.Count, "F").End(xlUp).Row

    ' Iterar desde la fila calculada (ultimaFila + 5) hasta la última fila con datos en la columna F
    For i = ultimaFila + 5 To ultimafilasumarsi
    wsComprasToken.Cells(i, 8).Formula = _
        "=SUMIF(B" & (ultimaFila + 5) & ":B" & ultimaFilaNotasCredito & ", G" & i & ", D" & (ultimaFila + 5) & ":D" & ultimaFilaNotasCredito & ")"
    Next i
    'Iterar desde la fila calculada (ultimaFila + 5) hasta la última fila con datos en la columna F
    For i = ultimaFila + 5 To ultimafilasumarsi
    wsComprasToken.Cells(i, 9).Formula = _
        "=SUMIF(B" & (ultimaFila + 5) & ":B" & ultimaFilaNotasCredito & ", G" & i & ", E" & (ultimaFila + 5) & ":E" & ultimaFilaNotasCredito & ")"
    Next i
    
    ' Agregar la fórmula SUMA en la columna G dos filas debajo de la última celda ocupada en la columna E
    ultimaFilasumanotascre = wsComprasToken.Cells(Rows.Count, "G").End(xlUp).Row
    wsComprasToken.Cells(ultimaFilasumanotascre + 2, 8).Formula = "=SUM(H" & ultimaFila + 5 & ":H" & ultimaFilasumanotascre & ")"
    
    ' Agregar la fórmula SUMA en la columna G dos filas debajo de la última celda ocupada en la columna E
    
    wsComprasToken.Cells(ultimaFilasumanotascre + 2, 9).Formula = "=SUM(I" & ultimaFila + 5 & ":I" & ultimaFilasumanotascre & ")"
           
   
    ' Aplicar filtro y copiar datos a una nueva hoja llamada "ventas_token" en el libro nuevo
    criterio1 = Array("Factura electrónica", "Documento equivalente POS", "Nota de débito electrónica")
    criterio2 = "emitido"

    ' Crear una nueva hoja para los datos filtrados
    Set wsVentasToken = NuevoLibro.Sheets.Add(After:=NuevoLibro.Sheets(NuevoLibro.Sheets.Count))
    wsVentasToken.Name = "ventas_token"
    
     On Error Resume Next ' Manejar posibles errores si los criterios no existen en los datos

    ' Aplicar el filtro
    With wsToken
        .Range("A:P").AutoFilter Field:=1, Criteria1:=criterio1, Operator:=xlFilterValues
        .Range("A:P").AutoFilter Field:=16, Criteria1:=criterio2

        ' Copiar los datos filtrados de las columnas J y N
        Set filtroRango = .Range("C:C,I:I,J:J,K:K,N:N").SpecialCells(xlCellTypeVisible)
        filtroRango.Copy Destination:=wsVentasToken.Range("A1")
        .AutoFilterMode = False
    End With
    
    On Error GoTo 0 ' Restablecer manejo de errores
    
    'Determinar Ultima Fila de la columna A para la formula sumar.si
    Ultimafilasumar = wsVentasToken.Cells(Rows.Count, "A").End(xlUp).Row

    ' Copiar los datos de la columna B (filtrada) a la columna E y eliminar duplicados
    wsVentasToken.Range("B:C").Copy
    wsVentasToken.Range("G1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    wsVentasToken.Range("G:H").RemoveDuplicates Columns:=1, Header:=xlNo

    ' Agregar la fórmula SUMAR.SI en la columna I
    ultimaFila = wsVentasToken.Cells(Rows.Count, "G").End(xlUp).Row
    For i = 2 To ultimaFila
        wsVentasToken.Cells(i, 9).Formula = "=SUMIF(B2:B" & Ultimafilasumar & ", G" & i & ", D2:D" & Ultimafilasumar & ")"
    Next i
    
    ' Agregar la fórmula SUMAR.SI en la columna J
    
    For i = 2 To ultimaFila
        wsVentasToken.Cells(i, 10).Formula = "=SUMIF(B2:B" & Ultimafilasumar & ", F" & i & ", E2:E" & Ultimafilasumar & ")"
    Next i
    
    ' Agregar la fórmula SUMA en la columna F dos filas debajo de la última celda ocupada en la columna E
    ultimaFila = wsVentasToken.Cells(Rows.Count, "H").End(xlUp).Row
    wsVentasToken.Cells(ultimaFila + 2, 9).Formula = "=SUM(I2:I" & ultimaFila & ")"
    
    wsVentasToken.Cells(ultimaFila + 2, 10).Formula = "=SUM(J2:J" & ultimaFila & ")"

    ' Determinar la ultima fila
    ultimaFilan = wsVentasToken.Cells(Rows.Count, "A").End(xlUp).Row
    wsVentasToken.Cells(ultimaFilan + 3, 1).Value = "Notas credito"
    
    

    ' Aplicar Filtro
    
    criterio1 = "Nota de crédito electrónica"
    criterio2 = "emitido"
    
    On Error Resume Next ' Manejar posibles errores si los criterios no existen en los datos
    
    ' Aplicar el filtro
    With wsToken
        .Range("A:P").AutoFilter Field:=1, Criteria1:=criterio1, Operator:=xlFilterValues
        .Range("A:P").AutoFilter Field:=16, Criteria1:=criterio2
        
    'Copiar los datos filtrados de las columnas J y N
        Set filtroRango = .Range("C:C,I:I,J:J,K:K,N:N").SpecialCells(xlCellTypeVisible)
        filtroRango.Copy Destination:=wsVentasToken.Cells(ultimaFilan + 5, 1)
        .AutoFilterMode = False
    End With
    
    On Error GoTo 0 ' Restablecer manejo de errores
    
    ' Determinar la ultima fila de notas credito
    ultimaFilaNotasCredito = wsVentasToken.Cells(Rows.Count, "A").End(xlUp).Row
    ' Copiar los datos de la columna B (filtrada) a la columna E y eliminar duplicados
    wsVentasToken.Range("B" & ultimaFilan + 5 & ":C" & ultimaFilaNotasCredito).Copy
    wsVentasToken.Cells(ultimaFilan + 5, "F").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    wsVentasToken.Range("F" & ultimaFilan + 5 & ":G" & ultimaFilaNotasCredito).RemoveDuplicates Columns:=1, Header:=xlNo

    ultimafilasumarsi = wsVentasToken.Cells(Rows.Count, "F").End(xlUp).Row
    
    For i = ultimaFilan + 5 To ultimafilasumarsi
    wsVentasToken.Cells(i, 8).Formula = _
    "=SUMIF(B" & (ultimaFilan + 5) & ":B" & ultimaFilaNotasCredito & ", F" & i & ", D" & (ultimaFilan + 5) & ":D" & ultimaFilaNotasCredito & ")"
    
    Next i
    
    For i = ultimaFilan + 5 To ultimafilasumarsi
    wsVentasToken.Cells(i, 9).Formula = _
    "=SUMIF(B" & (ultimaFilan + 5) & ":B" & ultimaFilaNotasCredito & ", F" & i & ", E" & (ultimaFilan + 5) & ":E" & ultimaFilaNotasCredito & ")"
    
    Next i
    
    ' Determinar la última fila ocupada en la columna G
    ultimaFilasumanotascre = wsVentasToken.Cells(wsVentasToken.Rows.Count, "H").End(xlUp).Row

    ' Agregar la fórmula SUM en la columna H, dos filas debajo de la última celda ocupada en la columna G
    wsVentasToken.Cells(ultimaFilasumanotascre + 2, 8).Formula = _
    "=SUM(H" & (ultimaFilan + 5) & ":H" & ultimaFilasumanotascre & ")"
    
    ' Agregar la fórmula SUM en la columna H, dos filas debajo de la última celda ocupada en la columna G
    wsVentasToken.Cells(ultimaFilasumanotascre + 2, 9).Formula = _
    "=SUM(I" & (ultimaFilan + 5) & ":I" & ultimaFilasumanotascre & ")"
    
     ' Agregar Buscarv en la hoja compras contables para traer el iva
       ultimaFila = wsCompras.Cells(Rows.Count, "J").End(xlUp).Row
        For i = 2 To ultimaFila
            
        Next i
        
        
    ' Crear una nueva hoja de excel en el libro de conciliacion
    Set wsbasededatos = NuevoLibro.Sheets.Add(After:=NuevoLibro.Sheets(NuevoLibro.Sheets.Count))
    wsbasededatos.Name = "BASE DE DATOS"
        
     
        
    ' Agregar formula buscarv en las hojas compras y ventas contables********* traer valor de iva en compras por tercero
               
    ' Determinar la ultima fila de la columna j
    ultimaFila = wsCompras.Cells(wsCompras.Rows.Count, "J").End(xlUp).Row
    
    Ultimafilabuscarv = Ivacompras.Cells(wsCompras.Rows.Count, "H").End(xlUp).Row
    

    Flasev = "FALSE"

y = 1
z = 7

' Definir Columnes y columnnum correctamente
Columnes = columnbuscavt + y
columnnum = Columnes - z
Columnabusvts = ColumnaLetra(Columnes)

' Bucle para aplicar la fórmula a cada fila
For i = 2 To ultimaFila

wsCompras.Cells(i, 12).Formula = _
"=IFERROR(VLOOKUP(J" & i & ",'Iva Compras Con'!H2:" & Columnabusvts & Ultimafilabuscarv & "," & columnnum & ", FALSE),0)"
Next i


    
     With Macrobasedetados
        
    ' Copiar columnas filtradas individualmente
        .Range("C:G").Copy Destination:=wsbasededatos.Range("A1")
        
        End With
       
   ' Agregar formula buscarv en hoja de compras para traer valores del iva***********
   
   ' Determinar la ultima fila de la columna j
    ultimaFila = wsCompras.Cells(wsCompras.Rows.Count, "J").End(xlUp).Row
   
   ' Bucle para aplicar la fórmula a cada fila
    For i = 2 To ultimaFila
    
wsCompras.Cells(i, 9).Formula = _
"=IFERROR(VLOOKUP(J" & i & ",'BASE DE DATOS'!A:E,5,FALSE),0)"

Next i

' Agregar formula buscarv en hoja de compras para traer valores del de la hoja token  iva***********


 ' Agregar formula buscarv en las hojas compras y ventas contables********* traer valor de iva en ventas por tercero
               
    ' Determinar la ultima fila de la columna j
    ultimaFila = wsVentas.Cells(wsVentas.Rows.Count, "J").End(xlUp).Row
    
    Ultimafilabuscarv = Ivagenerado.Cells(wsVentas.Rows.Count, "H").End(xlUp).Row
    

    Flasev = "FALSE"

y = 1
z = 7

' Definir Columnes y columnnum correctamente
Columnes = columnbuscavv + y
columnnum = Columnes - z
Columnabusvtv = ColumnaLetra(Columnes)

' Bucle para aplicar la fórmula a cada fila
For i = 2 To ultimaFila

wsVentas.Cells(i, 13).Formula = _
"=IFERROR(VLOOKUP(J" & i & ",'Iva Ventas Con'!H2:" & Columnabusvtv & Ultimafilabuscarv & "," & columnnum & ", FALSE),0)"

    
wsVentas.Cells(i, 9).Formula = _
"=IFERROR(VLOOKUP(J" & i & ",'BASE DE DATOS'!A:E,5,FALSE),0)"

Next i


' Determinar ultima fila hoja de compras token para traer valores a comparar en iva
wsCompras.Cells(1, 15).Value = "Valor Compras Token"
ultimaFila = wsCompras.Cells(wsCompras.Rows.Count, "I").End(xlUp).Row
ultimafilatkc = wsComprasToken.Cells(wsComprasToken.Rows.Count, "G").End(xlUp).Row

For i = 2 To ultimaFila

wsCompras.Cells(i, 15).Formula = _
    "=IFERROR(VLOOKUP(I" & i & ",'compras_token'!G:J,3,FALSE),0)"



Next i


    MsgBox "La conciliación de IVA se ha completado correctamente."
    
    
    
End Sub


