Attribute VB_Name = "Módulo1"
Sub EliminarCampos()
Attribute EliminarCampos.VB_ProcData.VB_Invoke_Func = " \n14"

' Asignación de variables
Dim wbOrigen As Workbook
Dim wsOrigen As Worksheet
Dim rngOrigen As Range
Dim Archivo As String
Dim hoja As Worksheet
Dim nombreEmpresa As String
Dim fechaFiltro As String
Dim ultimaFila As Long
Dim wbNuevo As Workbook
Dim rutaDestino As String
Dim nombreArchivo As String

  
'cuadro de dialogo para seleccioanr el archivo
Archivo = Application.GetOpenFilename("Archivos de Excel (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", , "Selecciona un archivo")


'Verificar si se selecciono un archivo
If Archivo = "Falso" Then
    MsgBox "Archivo no seleccionado.", vbExclamation
    Exit Sub
End If


'Desactivar actualización de pantalla para mejor velocidad
Application.ScreenUpdating = False
  
'Abrir el archivo seleccionado
Set wbOrigen = Workbooks.Open(Archivo)
Set wsOrigen = wbOrigen.Sheets(1)
Set rngOrigen = wsOrigen.UsedRange

'Copiar y pegar datos
rngOrigen.Copy
ThisWorkbook.Sheets(1).Range("A1").PasteSpecial Paste:=xlPasteValues

'Obtener valores de A4(empresa) y A8(fecha)
nombreEmpresa = wsOrigen.Range("A4").Value
fechaFiltro = wsOrigen.Range("A8").Value

'cerrar el archivo de origen sin guardar
wbOrigen.Close SaveChanges:=False

'Limpiar el portapapeles
Application.CutCopyMode = False
Application.ScreenUpdating = True

MsgBox "Datos copiados correctamente." & vbCrLf & _
           "Empresa: " & nombreEmpresa & vbCrLf & "Fecha: " & fechaFiltro, vbInformation

Set hoja = ThisWorkbook.Sheets("Hoja1") 'Cambiar "Hoja1" por el nombre de tu hoja si es diferente



With hoja
    If .AutoFilterMode Then .AutoFilterMode = False

    ' Aplicar filtro con criterios
    .Range("A14:F14").AutoFilter
    .Range("A14:F14").AutoFilter Field:=1, _
        Criteria1:=Array("", "CLIENTE", "DESDE", "FECHA", "Información cliente:", _
                             "Información General:", "Movimientos:", nombreEmpresa, fechaFiltro), _
        Operator:=xlFilterValues

' Eliminar las filas visibles (filtradas)
    On Error Resume Next
    .Range("A15", .Cells(.Rows.Count, 1).End(xlUp)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    On Error GoTo 0

    If .AutoFilterMode Then .AutoFilterMode = False
End With


' Crear nuevo archivo Excel sin macros
Set wbNuevo = Workbooks.Add
ultimaFila = hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row

' Copiar los datos restantes al nuevo libro
hoja.Range("A1:F" & ultimaFila).Copy Destination:=wbNuevo.Sheets(1).Range("A1")

' Guardar archivo nuevo como .xlsx
rutaDestino = Left(Archivo, InStrRev(Archivo, "\")) ' carpeta donde estaba el archivo original


' Mostrar cuadro de diálogo para que el usuario elija el nombre del archivo
nombreArchivo = Application.GetSaveAsFilename( _
    InitialFileName:=rutaDestino & "Nombre_personalizado.xlsx", _
    FileFilter:="Archivos de Excel (*.xlsx), *.xlsx", _
    Title:="Guardar archivo como")
    
    
' Verificar si el usuario canceló
If nombreArchivo = "Falso" Then
    MsgBox "Guardado cancelado.", vbExclamation
    wbNuevo.Close SaveChanges:=False
    Exit Sub
End If

' Guardar archivo
wbNuevo.SaveAs Filename:=nombreArchivo, FileFormat:=xlOpenXMLWorkbook
wbNuevo.Close SaveChanges:=True

MsgBox "Datos exportados correctamente a:" & vbCrLf & rutaDestino & nombreArchivo, vbInformation


End Sub
