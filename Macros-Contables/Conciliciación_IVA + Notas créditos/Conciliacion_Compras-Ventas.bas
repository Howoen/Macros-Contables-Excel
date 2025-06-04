Attribute VB_Name = "Módulo1"
Sub EjecutarFuncionesEnArchivos()
    Dim NuevoLibro As Workbook
    Dim NombreLibro As String
    Dim Ruta As String
    Dim RutaCompleta As String
    Dim archivoSeleccionado As Variant
    Dim archivoToken As Variant
    Dim ws As Worksheet
    Dim wsToken As Worksheet
    Dim wsCompras As Worksheet
    Dim wsVentas As Worksheet
    Dim wsComprasToken As Worksheet
    Dim wsVentasToken As Worksheet
    Dim inicio As Range
    Dim fin As Range
    Dim descripcionInicio As String
    Dim descripcionFin As String
    Dim rango As Range
    Dim criterio1 As String
    Dim criterio2 As String
    Dim filtroRango As Range
    Dim i As Long
    Dim ultimaFila As Long

    ' Solicitar al usuario el nombre del libro
    NombreLibro = InputBox("Por favor, ingresa el mes a conciliar:", "Nombre del Libro")
    
    ' Verificar si el usuario ha ingresado un nombre
    If NombreLibro = "" Then
        MsgBox "No se ingresó un nombre para el libro. La operación se cancelará."
        Exit Sub
    End If
    
    ' Solicitar al usuario la ubicación para guardar el libro
    Ruta = Application.GetSaveAsFilename(InitialFileName:=NombreLibro, FileFilter:="Archivos de Excel (*.xlsx), *.xlsx")
    
    ' Verificar si el usuario ha ingresado una ruta válida
    If Ruta = "False" Then
        MsgBox "No se seleccionó una ubicación válida. La operación se cancelará."
        Exit Sub
    End If
    
    ' Crear un nuevo libro
    Set NuevoLibro = Workbooks.Add
    
    ' Guardar el nuevo libro en la ubicación especificada por el usuario
    NuevoLibro.SaveAs Filename:=Ruta
    
    ' Mostrar mensaje de confirmación
    MsgBox "El nuevo libro ha sido creado y guardado en: " & Ruta
    
    ' Preguntar al usuario en qué libro desea ejecutar las funciones
    archivoSeleccionado = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Seleccione el archivo del software contable")
    
    If archivoSeleccionado = False Then
        MsgBox "No se seleccionó ningún archivo. La operación ha sido cancelada."
        Exit Sub
    End If

    ' Abre el libro seleccionado por el usuario
    Workbooks.Open archivoSeleccionado
    Set ws = ActiveSheet

    ' Define las descripciones a buscar
    descripcionInicio = "Factura de Compra"
    descripcionFin = "Total Factura de Compra"

    ' Busca la descripción inicial en la columna A
    Set inicio = ws.Columns("A").Find(What:=descripcionInicio, LookIn:=xlValues, LookAt:=xlWhole)

    ' Busca la descripción final en la columna A
    Set fin = ws.Columns("A").Find(What:=descripcionFin, LookIn:=xlValues, LookAt:=xlWhole)

    ' Verifica si se encontraron ambas descripciones
    If Not inicio Is Nothing And Not fin Is Nothing Then
        ' Define el rango utilizando las coordenadas encontradas
        Set rango = ws.Range(ws.Cells(inicio.Row, 2), ws.Cells(fin.Row, 8)) ' Columnas B a H

        ' Crea una nueva hoja llamada "compras" en el libro nuevo
        Set wsCompras = NuevoLibro.Sheets.Add(After:=NuevoLibro.Sheets(NuevoLibro.Sheets.Count))
        wsCompras.Name = "compras"

        ' Copia y pega el rango en la nueva hoja
        rango.Copy Destination:=wsCompras.Range("A1")

        ' Eliminar los duplicados de la columna D y copiar a la columna J
        wsCompras.Range("D:D").Copy
        wsCompras.Range("J1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        wsCompras.Range("J:J").RemoveDuplicates Columns:=1, Header:=xlNo

        ' Agregar la fórmula SUMAR.SI en la columna K
        ultimaFila = wsCompras.Cells(Rows.Count, "J").End(xlUp).Row
        For i = 1 To ultimaFila
            wsCompras.Cells(i, 11).Formula = "=SUMIF(D:D, J" & i & ", G:G)"
        Next i
        
        ' Agregar la fórmula SUMA en la columna K dos filas debajo de la última celda ocupada en la columna J
        wsCompras.Cells(ultimaFila + 2, 11).Formula = "=SUM(K1:K" & ultimaFila & ")"

        MsgBox "El rango encontrado ha sido copiado y pegado en la hoja 'compras'."
    Else
        MsgBox "No se encontraron ambas descripciones en la columna A."
        Exit Sub
    End If

    ' Define las descripciones a buscar para ventas
    descripcionInicio = "Factura de Venta"
    descripcionFin = "Total Factura de Venta"

    ' Busca la descripción inicial en la columna A
    Set inicio = ws.Columns("A").Find(What:=descripcionInicio, LookIn:=xlValues, LookAt:=xlWhole)

    ' Busca la descripción final en la columna A
    Set fin = ws.Columns("A").Find(What:=descripcionFin, LookIn:=xlValues, LookAt:=xlWhole)

    ' Verifica si se encontraron ambas descripciones
    If Not inicio Is Nothing And Not fin Is Nothing Then
        ' Define el rango utilizando las coordenadas encontradas
        Set rango = ws.Range(ws.Cells(inicio.Row, 2), ws.Cells(fin.Row, 8)) ' Columnas B a H

        ' Crea una nueva hoja llamada "ventas" en el libro nuevo
        Set wsVentas = NuevoLibro.Sheets.Add(After:=NuevoLibro.Sheets(NuevoLibro.Sheets.Count))
        wsVentas.Name = "ventas"

        ' Copia y pega el rango en la nueva hoja
        rango.Copy Destination:=wsVentas.Range("A1")

        ' Eliminar los duplicados de la columna D y copiar a la columna J
        wsVentas.Range("D:D").Copy
        wsVentas.Range("J1").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        wsVentas.Range("J:J").RemoveDuplicates Columns:=1, Header:=xlNo

        ' Agregar la fórmula SUMAR.SI en la columna K
        ultimaFila = wsVentas.Cells(Rows.Count, "J").End(xlUp).Row
        For i = 1 To ultimaFila
            wsVentas.Cells(i, 11).Formula = "=SUMIF(D:D, J" & i & ", F:F)"
        Next i
        
        ' Agregar la fórmula SUMA en la columna K dos filas debajo de la última celda ocupada en la columna J
        wsVentas.Cells(ultimaFila + 2, 11).Formula = "=SUM(K1:K" & ultimaFila & ")"

        MsgBox "El rango encontrado ha sido copiado y pegado en la hoja 'ventas'."
    Else
        MsgBox "No se encontraron ambas descripciones en la columna A."
        Exit Sub
    End If
  
 

    ' Seleccionar el archivo Token DIAN
    archivoToken = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Seleccione el archivo del Token DIAN")

    If archivoToken = False Then
        MsgBox "No se seleccionó ningún archivo. La operación ha sido cancelada."
        Exit Sub
    End If

    ' Abre el libro Token DIAN
    Workbooks.Open archivoToken
    Set wsToken = ActiveSheet

    ' Aplicar filtro y copiar datos a una nueva hoja llamada "compras_token" en el libro nuevo
    criterio1 = "Factura electrónica"
    criterio2 = "recibido"

    ' Crear una nueva hoja para los datos filtrados
    Set wsComprasToken = NuevoLibro.Sheets.Add(After:=NuevoLibro.Sheets(NuevoLibro.Sheets.Count))
    wsComprasToken.Name = "compras_token"

    ' Aplicar el filtro
    With wsToken
        .Range("A:AF").AutoFilter Field:=1, Criteria1:=criterio1
        .Range("A:AF").AutoFilter Field:=32, Criteria1:=criterio2

        ' Copiar los datos filtrados de las columnas K y AD
        Set filtroRango = .Range("C:C,K:K,AD:AD").SpecialCells(xlCellTypeVisible)
        filtroRango.Copy Destination:=wsComprasToken.Range("A1")
        .AutoFilterMode = False
    End With

    ' Copiar los datos de la columna B (filtrada) a la columna E y eliminar duplicados
    wsComprasToken.Range("B:B").Copy
    wsComprasToken.Range("E1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    wsComprasToken.Range("E:E").RemoveDuplicates Columns:=1, Header:=xlNo

    ' Agregar la fórmula SUMAR.SI en la columna F
    ultimaFila = wsComprasToken.Cells(Rows.Count, "E").End(xlUp).Row
    For i = 1 To ultimaFila
        wsComprasToken.Cells(i, 6).Formula = "=SUMIF(B:B, E" & i & ", C:C)"
    Next i
     
    ' Agregar la fórmula SUMA en la columna F dos filas debajo de la última celda ocupada en la columna E
    ultimaFila = wsComprasToken.Cells(Rows.Count, "E").End(xlUp).Row
    wsComprasToken.Cells(ultimaFila + 2, 6).Formula = "=SUM(F1:F" & ultimaFila & ")"

    MsgBox "Los datos filtrados han sido copiados y las fórmulas agregadas en la hoja 'compras_token'."

    ' Aplicar filtro y copiar datos a una nueva hoja llamada "ventas_token" en el libro nuevo
    criterio1 = "Factura electrónica"
    criterio2 = "emitido"

    ' Crear una nueva hoja para los datos filtrados
    Set wsVentasToken = NuevoLibro.Sheets.Add(After:=NuevoLibro.Sheets(NuevoLibro.Sheets.Count))
    wsVentasToken.Name = "ventas_token"

    ' Aplicar el filtro
    With wsToken
        .Range("A:AF").AutoFilter Field:=1, Criteria1:=criterio1
        .Range("A:AF").AutoFilter Field:=32, Criteria1:=criterio2

        ' Copiar los datos filtrados de las columnas J y N
        Set filtroRango = .Range("C:C,M:M,AD:AD").SpecialCells(xlCellTypeVisible)
        filtroRango.Copy Destination:=wsVentasToken.Range("A1")
        .AutoFilterMode = False
    End With

    ' Copiar los datos de la columna B (filtrada) a la columna E y eliminar duplicados
    wsVentasToken.Range("B:B").Copy
    wsVentasToken.Range("E1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    wsVentasToken.Range("E:E").RemoveDuplicates Columns:=1, Header:=xlNo

    ' Agregar la fórmula SUMAR.SI en la columna F
    ultimaFila = wsVentasToken.Cells(Rows.Count, "E").End(xlUp).Row
    For i = 1 To ultimaFila
        wsVentasToken.Cells(i, 6).Formula = "=SUMIF(B:B, E" & i & ", C:C)"
    Next i
    
    ' Agregar la fórmula SUMA en la columna F dos filas debajo de la última celda ocupada en la columna E
    ultimaFila = wsVentasToken.Cells(Rows.Count, "E").End(xlUp).Row
    wsVentasToken.Cells(ultimaFila + 2, 6).Formula = "=SUM(F1:F" & ultimaFila & ")"

    MsgBox "Los datos filtrados han sido copiados y las fórmulas agregadas en la hoja 'ventas_token'."
    
    ' --- Notas crédito recibidas (Compras)---
    criterio1 = "Nota de crédito electrónica"
    criterio2 = "recibido"
    
    Set wsComprasToken = NuevoLibro.Sheets.Add(After:=NuevoLibro.Sheets(NuevoLibro.Sheets.Count))
    wsComprasToken.Name = "notas_credito_compras"
    
    With wsToken
        .Range("A:AF").AutoFilter Field:=1, Criteria1:=criterio1
        .Range("A:AF").AutoFilter Field:=32, Criteria1:=criterio2
        
        Set filtroRango = .Range("C:C,K:K,AD:AD").SpecialCells(xlCellTypeVisible)
        filtroRango.Copy Destination:=wsComprasToken.Range("A1")
        .AutoFilterMode = False
    End With
    
    wsComprasToken.Range("B:B").Copy
    wsComprasToken.Range("E1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    wsComprasToken.Range("E:E").RemoveDuplicates Columns:=1, Header:=xlNo
    
    ultimaFila = wsComprasToken.Cells(Rows.Count, "E").End(xlUp).Row
    For i = 1 To ultimaFila
        wsComprasToken.Cells(i, 6).Formula = "=-1 * SUMIF(B:B, E" & i & ", C:C)"
    Next i
    
    wsComprasToken.Cells(ultimaFila + 2, 6).Formula = "=SUM(F1:F" & ultimaFila & ")"
    MsgBox "Notas crédito recibidos copiados y procesados en 'notas_crédito_compras'."
    
    
    ' --- Notas crédito emitidas (vetas) ---
    criterio2 = "emitido"
    
    Set wsVentasToken = NuevoLibro.Sheets.Add(After:=NuevoLibro.Sheets(NuevoLibro.Sheets.Count))
    wsVentasToken.Name = "notas_credito_ventas"
    
    With wsToken
        .Range("A:AF").AutoFilter Field:=1, Criteria1:=criterio1
        .Range("A:AF").AutoFilter Field:=32, Criteria1:=criterio2
        
        Set filtroRango = .Range("C:C,M:M,AD:AD").SpecialCells(xlCellTypeVisible)
        filtroRango.Copy Destination:=wsVentasToken.Range("A1")
        .AutoFilterMode = False
    End With
    
    wsVentasToken.Range("B:B").Copy
    wsVentasToken.Range("E1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    wsVentasToken.Range("E:E").RemoveDuplicates Columns:=1, Header:=xlNo
    
    ultimaFila = wsVentasToken.Cells(Rows.Count, "E").End(xlUp).Row
    For i = 1 To ultimaFila
        wsVentasToken.Cells(i, 6).Formula = "=-1 * SUMIF(B:B, E" & i & ", C:C)"
    Next i
    
    wsVentasToken.Cells(ultimaFila + 2, 6).Formula = "=SUM(F1:F" & ultimaFila & ")"
    MsgBox "Notas crédito emitidas copiadas y procesadas en 'notas_credito_ventas'."
    
    

    ' Guardar y cerrar los libros abiertos
    NuevoLibro.Save
    ActiveWorkbook.Close
End Sub

