VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormReporte 
   Caption         =   "FORMULARIO DE REPORTE"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11610
   OleObjectBlob   =   "FormReporte.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' FORMULARIO: FrmReportes
' DESCRIPCIÓN: Generador de reportes con filtros combinados y exportación a PDF.
' VERSIÓN: 2.0 - Filtros combinados (Agencia + Estado + Tipo + Garantía)
' ==============================================================================

' Constante para nombre de hoja temporal de reporte
Private Const HOJA_REPORTE_TEMP As String = "ReporteTEMP"

' ==============================================================================
' INICIALIZACIÓN
' ==============================================================================
Private Sub UserForm_Initialize()
    ' 1. Configurar el ListBox de vista previa
    With Me.lstPreview
        .ColumnCount = 6
        .ColumnWidths = "70;140;100;100;100;80"
    End With
    
    ' 2. Todos los filtros visibles desde el inicio
    Me.cmbAgencia.Visible = True
    Me.cmbEstado.Visible = True
    Me.cmbGarantia.Visible = True
    Me.cmbTipo.Visible = True
    Me.lblAgencia.Visible = True
    Me.lblEstado.Visible = True
    Me.lblGarantia.Visible = True
    Me.lblTipo.Visible = True
    
    ' 3. Estado inicial de botones y etiqueta
    Me.btnPDF.Enabled = False
    Me.lblConteo.Caption = "Configurá los filtros y presioná Vista Previa."
    
    ' 4. Cargar todos los combos desde Config (Agencias, Estados, Tipos, Garantía)
    CargarFiltrosDesdeConfig
    
    ' 5. Tooltips informativos
    Me.btnVista.ControlTipText = "Previsualizar registros según los filtros seleccionados"
    Me.btnPDF.ControlTipText = "Exportar el reporte actual a PDF"
    Me.btnCerrar.ControlTipText = "Cerrar reportes y volver al formulario principal"
End Sub

' ==============================================================================
' BOTÓN: Vista Previa
' ==============================================================================
Private Sub btnVista_Click()
    EjecutarFiltroReporte
    Me.btnPDF.Enabled = (Me.lstPreview.ListCount > 0)
End Sub

' ==============================================================================
' LÓGICA CENTRAL: EjecutarFiltroReporte
' Todos los filtros son opcionales y se combinan entre sí.
' Si un filtro queda en "(Todos...)" se ignora y no restringe resultados.
' ==============================================================================
Private Sub EjecutarFiltroReporte()
    Dim datos      As Variant
    Dim totalFilas As Long
    Dim i          As Long
    Dim conteo     As Long
    
    ' --- 1. OBTENER DATOS EN ARRAY ---
    datos = ObtenerDatosTablaArray("Inventario", "TablaInventario")
    If IsEmpty(datos) Then
        MsgBox "No se encontraron datos en TablaInventario.", vbInformation
        Exit Sub
    End If
    
    totalFilas = UBound(datos, 1)
    conteo = 0
    Me.lstPreview.Clear
    
    ' --- 2. LEER FILTROS SELECCIONADOS ---
    Dim filtroAgencia  As String
    Dim filtroEstado   As String
    Dim filtroTipo     As String
    Dim filtroGarantia As String
    
    filtroAgencia = Me.cmbAgencia.Value
    filtroEstado = Me.cmbEstado.Value
    filtroTipo = Me.cmbTipo.Value
    filtroGarantia = Me.cmbGarantia.Value
    
    ' Si eligieron la opción "Todos..." la tratamos como sin filtro
    If InStr(filtroAgencia, "(Todas") > 0 Then filtroAgencia = ""
    If InStr(filtroEstado, "(Todos") > 0 Then filtroEstado = ""
    If InStr(filtroTipo, "(Todos") > 0 Then filtroTipo = ""
    If InStr(filtroGarantia, "(Todas") > 0 Then filtroGarantia = ""
    
    ' --- 3. RECORRER ARRAY Y APLICAR FILTROS COMBINADOS ---
    For i = 1 To totalFilas
        
        ' Saltar filas vacías
        If Trim(CStr(datos(i, 1))) = "" Then GoTo SiguienteFilaReporte
        
        Dim incluir As Boolean
        incluir = True  ' Asumimos que incluir; cada filtro activo puede rechazarlo
        
        ' FILTRO 1: AGENCIA (Col 2)
        If filtroAgencia <> "" Then
            If UCase(Trim(CStr(datos(i, 2)))) <> UCase(filtroAgencia) Then
                incluir = False
            End If
        End If
        
        ' FILTRO 2: ESTADO (Col 9)
        If incluir And filtroEstado <> "" Then
            If UCase(Trim(CStr(datos(i, 9)))) <> UCase(filtroEstado) Then
                incluir = False
            End If
        End If
        
        ' FILTRO 3: TIPO DE EQUIPO (Col 6)
        If incluir And filtroTipo <> "" Then
            If UCase(Trim(CStr(datos(i, 6)))) <> UCase(filtroTipo) Then
                incluir = False
            End If
        End If
        
        ' FILTRO 4: GARANTÍA (Col 36 = Fecha Fin Garantía)
        If incluir And filtroGarantia <> "" Then
            Dim fechaFin As Variant
            fechaFin = datos(i, 36)
            
            If IsDate(fechaFin) Then
                Dim diasRestantes As Long
                diasRestantes = DateDiff("d", Date, CDate(fechaFin))
                
                Select Case filtroGarantia
                    Case "Garantía Activa"
                        incluir = (diasRestantes > 30)
                    Case "Por vencer (30 días)"
                        incluir = (diasRestantes >= 0 And diasRestantes <= 30)
                    Case "Por vencer (60 días)"
                        incluir = (diasRestantes >= 0 And diasRestantes <= 60)
                    Case "Por vencer (90 días)"
                        incluir = (diasRestantes >= 0 And diasRestantes <= 90)
                    Case "Ya vencida"
                        incluir = (diasRestantes < 0)
                End Select
            Else
                ' Sin fecha de garantía cargada ? no cumple ningún filtro de garantía
                incluir = False
            End If
        End If
        
        ' --- AGREGAR AL LISTBOX SI PASÓ TODOS LOS FILTROS ---
        If incluir Then
            conteo = conteo + 1
            With Me.lstPreview
                .AddItem CStr(datos(i, 1))                     ' Código
                .List(.ListCount - 1, 1) = CStr(datos(i, 5))  ' Usuario
                .List(.ListCount - 1, 2) = CStr(datos(i, 2))  ' Agencia
                .List(.ListCount - 1, 3) = CStr(datos(i, 6))  ' Tipo Equipo
                .List(.ListCount - 1, 4) = CStr(datos(i, 10)) ' Serie
                .List(.ListCount - 1, 5) = CStr(datos(i, 9))  ' Estado
            End With
        End If
        
SiguienteFilaReporte:
    Next i
    
    ' --- 4. ACTUALIZAR CONTADOR ---
    If conteo = 0 Then
        Me.lblConteo.Caption = "No se encontraron registros con esos filtros."
        Me.lblConteo.ForeColor = RGB(200, 0, 0)
    Else
        Me.lblConteo.Caption = conteo & " registro(s) encontrado(s)."
        Me.lblConteo.ForeColor = RGB(0, 120, 0)
    End If
    
    Me.btnPDF.Enabled = (conteo > 0)
End Sub

' ==============================================================================
' FUNCIÓN: ConstruirHojaReporte
' Crea hoja temporal formateada ? se exporta a PDF ? se elimina.
' ==============================================================================
Private Function ConstruirHojaReporte() As Worksheet
    Dim ws   As Worksheet
    Dim fila As Long
    Dim col  As Integer
    
    ' Eliminar hoja temporal anterior si existe
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(HOJA_REPORTE_TEMP).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    On Error GoTo ErrorHoja
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = HOJA_REPORTE_TEMP
    
    ' --- Construir descripción de filtros activos para el subtítulo ---
    Dim filtroDesc As String
    filtroDesc = ""
    If Me.cmbAgencia.Value <> "" And InStr(Me.cmbAgencia.Value, "(Todas") = 0 Then
        filtroDesc = filtroDesc & "Agencia: " & Me.cmbAgencia.Value & "   "
    End If
    If Me.cmbEstado.Value <> "" And InStr(Me.cmbEstado.Value, "(Todos") = 0 Then
        filtroDesc = filtroDesc & "Estado: " & Me.cmbEstado.Value & "   "
    End If
    If Me.cmbTipo.Value <> "" And InStr(Me.cmbTipo.Value, "(Todos") = 0 Then
        filtroDesc = filtroDesc & "Tipo: " & Me.cmbTipo.Value & "   "
    End If
    If Me.cmbGarantia.Value <> "" And InStr(Me.cmbGarantia.Value, "(Todas") = 0 Then
        filtroDesc = filtroDesc & "Garantía: " & Me.cmbGarantia.Value
    End If
    If Trim(filtroDesc) = "" Then filtroDesc = "Inventario Completo (sin filtros)"
    
    With ws
        ' --- TÍTULO PRINCIPAL ---
        .Cells(1, 1).Value = "REPORTE DE INVENTARIO IT"
        .Cells(1, 1).Font.Size = 16
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = RGB(31, 78, 121)
        .Range("A1:F1").Merge
        .Range("A1:F1").HorizontalAlignment = xlCenter
        
        ' --- SUBTÍTULO CON FILTROS APLICADOS ---
        .Cells(2, 1).Value = filtroDesc
        .Cells(2, 1).Font.Size = 10
        .Cells(2, 1).Font.Italic = True
        .Cells(2, 1).Font.Color = RGB(68, 68, 68)
        .Range("A2:F2").Merge
        .Range("A2:F2").HorizontalAlignment = xlCenter
        
        ' --- FECHA, HORA Y USUARIO ---
        .Cells(3, 1).Value = "Generado: " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & _
                                  "   |   Usuario: " & Environ("USERNAME")
        .Cells(3, 1).Font.Size = 9
        .Cells(3, 1).Font.Color = RGB(120, 120, 120)
        .Range("A3:F3").Merge
        .Range("A3:F3").HorizontalAlignment = xlCenter
        
        ' --- LÍNEA SEPARADORA ---
        .Range("A4:F4").Interior.Color = RGB(31, 78, 121)
        .Rows(4).RowHeight = 3
        
        ' --- ENCABEZADOS DE COLUMNAS ---
        Dim encabezados As Variant
        encabezados = Array("Código", "Usuario", "Agencia", "Tipo Equipo", "Serie", "Estado")
        For col = 0 To 5
            .Cells(5, col + 1).Value = encabezados(col)
        Next col
        With .Range("A5:F5")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        
        ' --- DATOS DESDE EL LISTBOX ---
        fila = 6
        Dim colorPar   As Long
        Dim colorImpar As Long
        colorPar = RGB(235, 241, 250)
        colorImpar = RGB(255, 255, 255)
        
        Dim j As Long
        For j = 0 To Me.lstPreview.ListCount - 1
            Dim colorFila As Long
            If j Mod 2 = 0 Then colorFila = colorImpar Else colorFila = colorPar
            
            For col = 0 To 5
                .Cells(fila, col + 1).Value = Me.lstPreview.List(j, col)
                .Cells(fila, col + 1).Interior.Color = colorFila
                .Cells(fila, col + 1).Font.Size = 9
            Next col
            
            With .Range(.Cells(fila, 1), .Cells(fila, 6))
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(200, 200, 200)
            End With
            
            fila = fila + 1
        Next j
        
        ' --- PIE: TOTAL DE REGISTROS ---
        fila = fila + 1
        .Cells(fila, 1).Value = "Total de registros: " & Me.lstPreview.ListCount
        .Cells(fila, 1).Font.Bold = True
        .Cells(fila, 1).Font.Size = 10
        .Range(.Cells(fila, 1), .Cells(fila, 3)).Merge
        
        ' --- ANCHO DE COLUMNAS ---
        .Columns("A").ColumnWidth = 10
        .Columns("B").ColumnWidth = 22
        .Columns("C").ColumnWidth = 18
        .Columns("D").ColumnWidth = 14
        .Columns("E").ColumnWidth = 18
        .Columns("F").ColumnWidth = 16
        
        ' --- CONFIGURACIÓN DE PÁGINA ---
        With .PageSetup
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.5)
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .CenterHorizontally = True
            .RightFooter = "Página &P de &N"
            .LeftFooter = "Inventario IT - Confidencial"
        End With
        
    End With
    
    Set ConstruirHojaReporte = ws
    Exit Function
    
ErrorHoja:
    Set ConstruirHojaReporte = Nothing
End Function

' ==============================================================================
' LIMPIEZA: Eliminar hoja temporal
' ==============================================================================
Private Sub LimpiarHojaTemp(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

' ==============================================================================
' BOTÓN: Exportar — elige formato y delega
' ==============================================================================
Private Sub btnPDF_Click()
    If Me.lstPreview.ListCount = 0 Then
        MsgBox "No hay datos para generar el reporte." & vbCrLf & _
               "Presioná primero 'Vista Previa'.", vbExclamation
        Exit Sub
    End If

    Dim respFormato As VbMsgBoxResult
    respFormato = MsgBox( _
        "¿En qué formato querés exportar el reporte?" & vbCrLf & vbCrLf & _
        "     SÍ   ?  PDF" & vbCrLf & _
        "     NO   ?  Excel (.xlsx)", _
        vbYesNoCancel + vbQuestion, "Formato de exportación")

    Select Case respFormato
        Case vbYes:    ExportarPDF
        Case vbNo:     ExportarExcel
        Case vbCancel: Exit Sub
    End Select
End Sub

' ==============================================================================
' EXPORTAR PDF  — lógica original extraída sin cambios funcionales
' ==============================================================================
Private Sub ExportarPDF()
    ' 1. Nombre sugerido con filtros activos
    Dim nombreSugerido As String
    nombreSugerido = "Reporte_Inventario"
    If Me.cmbAgencia.Value <> "" And InStr(Me.cmbAgencia.Value, "(Todas") = 0 Then
        nombreSugerido = nombreSugerido & "_" & Me.cmbAgencia.Value
    End If
    If Me.cmbEstado.Value <> "" And InStr(Me.cmbEstado.Value, "(Todos") = 0 Then
        nombreSugerido = nombreSugerido & "_" & Me.cmbEstado.Value
    End If
    If Me.cmbTipo.Value <> "" And InStr(Me.cmbTipo.Value, "(Todos") = 0 Then
        nombreSugerido = nombreSugerido & "_" & Me.cmbTipo.Value
    End If
    nombreSugerido = nombreSugerido & "_" & Format(Now(), "yyyymmdd_hhmmss")
    nombreSugerido = Replace(nombreSugerido, "/", "-")
    nombreSugerido = Replace(nombreSugerido, " ", "_")

    ' 2. Diálogo guardar
    Dim rutaPDF As String
    rutaPDF = Application.GetSaveAsFilename( _
        InitialFileName:=nombreSugerido, _
        FileFilter:="Archivo PDF (*.pdf), *.pdf", _
        Title:="Guardar Reporte como PDF")

    If VarType(rutaPDF) = vbBoolean Then Exit Sub
    If LCase(Right(rutaPDF, 4)) <> ".pdf" Then rutaPDF = rutaPDF & ".pdf"

    ' 3. Construir hoja temporal y exportar
    Dim wsReporte As Worksheet
    Set wsReporte = ConstruirHojaReporte()

    If wsReporte Is Nothing Then
        MsgBox "Error al construir el reporte. Intentalo de nuevo.", vbCritical
        Exit Sub
    End If

    On Error GoTo ErrorPDF
    wsReporte.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=rutaPDF, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    LimpiarHojaTemp wsReporte
    MsgBox "? Reporte generado exitosamente:" & vbCrLf & rutaPDF, vbInformation, "PDF Generado"
    Exit Sub

ErrorPDF:
    LimpiarHojaTemp wsReporte
    MsgBox "Error al generar el PDF: " & Err.Description & vbCrLf & _
           "Verificá que el archivo no esté abierto en otro programa.", vbCritical
End Sub

' ==============================================================================
' EXPORTAR EXCEL — orquesta selección de hojas, advertencia, creación del libro
' ==============================================================================
Private Sub ExportarExcel()

    ' --- PASO 1: Seleccionar hojas ---
    Dim opcion As String
    opcion = InputBox( _
        "¿Qué hojas querés exportar?" & vbCrLf & vbCrLf & _
        "  1  ?  Inventario     (con filtros aplicados)" & vbCrLf & _
        "  2  ?  Préstamos      (completa)" & vbCrLf & _
        "  3  ?  Fuera de Uso   (completa)" & vbCrLf & _
        "  4  ?  Las 3 hojas" & vbCrLf & vbCrLf & _
        "Ingresá el número (1 / 2 / 3 / 4):", _
        "Seleccionar hojas")

    If opcion = "" Then Exit Sub  ' Canceló

    If opcion <> "1" And opcion <> "2" And opcion <> "3" And opcion <> "4" Then
        MsgBox "Opción inválida. Ingresá 1, 2, 3 o 4.", vbExclamation
        Exit Sub
    End If

    ' --- PASO 2: Advertencia de filtros si Inventario está incluida ---
    If opcion = "1" Or opcion = "4" Then
        Dim msgAviso As String
        msgAviso = "ATENCIÓN — Filtros aplicados sobre Inventario:" & vbCrLf & vbCrLf & _
                   ConstruirDescFiltros() & vbCrLf & _
                   "Los filtros afectan ÚNICAMENTE la hoja Inventario exportada."

        If opcion = "4" Then
            msgAviso = msgAviso & vbCrLf & _
                       "Préstamos y Fuera de Uso se exportan completas."
        End If

        msgAviso = msgAviso & vbCrLf & vbCrLf & "¿Continuar?"

        If MsgBox(msgAviso, vbOKCancel + vbInformation, "Confirmar exportación") = vbCancel Then
            Exit Sub
        End If
    End If

    ' --- PASO 3: Diálogo guardar ---
    Dim nombreSugerido As String
    nombreSugerido = "Reporte_Inventario_" & Format(Now(), "yyyymmdd_hhmmss")
    nombreSugerido = Replace(nombreSugerido, " ", "_")

    Dim rutaXLS As Variant
    rutaXLS = Application.GetSaveAsFilename( _
        InitialFileName:=nombreSugerido, _
        FileFilter:="Libro Excel (*.xlsx), *.xlsx", _
        Title:="Guardar Reporte como Excel")

    If VarType(rutaXLS) = vbBoolean Then Exit Sub
    If LCase(Right(rutaXLS, 5)) <> ".xlsx" Then rutaXLS = rutaXLS & ".xlsx"

    ' --- PASO 4: Crear workbook nuevo con 1 sola hoja default ---
    Dim sheetsAntes As Integer
    sheetsAntes = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = 1

    Dim wbNuevo As Workbook
    Set wbNuevo = Workbooks.Add

    Application.SheetsInNewWorkbook = sheetsAntes  ' Restaurar preferencia del usuario

    ' Guardar referencia a la hoja vacía default para eliminarla al final
    Dim wsDefault As Worksheet
    Set wsDefault = wbNuevo.Sheets(1)

    On Error GoTo ErrorExcel

    ' --- PASO 5: Agregar las hojas solicitadas ---
    If opcion = "1" Or opcion = "4" Then AgregarHojaInventario wbNuevo
    If opcion = "2" Or opcion = "4" Then AgregarHojaDesde wbNuevo, "Prestamos"
    If opcion = "3" Or opcion = "4" Then AgregarHojaDesde wbNuevo, "Fuera de Uso"

    ' Eliminar la hoja vacía default (ahora que ya hay otras hojas)
    Application.DisplayAlerts = False
    wsDefault.Delete
    Application.DisplayAlerts = True

    ' --- PASO 6: Guardar y cerrar ---
    wbNuevo.SaveAs Filename:=rutaXLS, FileFormat:=xlOpenXMLWorkbook
    wbNuevo.Close SaveChanges:=False

    MsgBox "? Excel generado exitosamente:" & vbCrLf & rutaXLS, vbInformation, "Excel Generado"
    Exit Sub

ErrorExcel:
    Application.DisplayAlerts = False
    If Not wbNuevo Is Nothing Then wbNuevo.Close SaveChanges:=False
    Application.DisplayAlerts = True
    MsgBox "Error al generar el Excel: " & Err.Description & vbCrLf & _
           "Verificá que el archivo no esté abierto en otro programa.", vbCritical
End Sub

' ==============================================================================
' HELPER: Construir hoja "Inventario" desde lstPreview (filtros ya aplicados)
' ==============================================================================
Private Sub AgregarHojaInventario(wbDestino As Workbook)
    Dim ws As Worksheet
    Set ws = wbDestino.Sheets.Add(After:=wbDestino.Sheets(wbDestino.Sheets.Count))
    ws.Name = "Inventario"

    ' Encabezados (deben coincidir con las columnas del lstPreview)
    Dim headers As Variant
    headers = Array("Código", "Usuario", "Agencia", "Tipo Equipo", "Serie", "Estado")
    Dim col As Integer
    For col = 0 To 5
        ws.Cells(1, col + 1).Value = headers(col)
    Next col

    ' Datos fila por fila desde el ListBox
    Dim j As Long
    For j = 0 To Me.lstPreview.ListCount - 1
        For col = 0 To 5
            ws.Cells(j + 2, col + 1).Value = Me.lstPreview.List(j, col)
        Next col
    Next j

    ws.Columns("A:F").AutoFit

    ' Aplicar formato tabla si hay al menos 1 fila de datos
    Dim totalFilas As Long
    totalFilas = Me.lstPreview.ListCount + 1  ' +1 = encabezado

    If totalFilas > 1 Then
        Dim tbl As ListObject
        Set tbl = ws.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=ws.Range("A1").Resize(totalFilas, 6), _
            XlListObjectHasHeaders:=xlYes)
        tbl.Name = "Tabla_Inventario"
        tbl.TableStyle = "TableStyleMedium2"
    End If
End Sub

' ==============================================================================
' HELPER: Construir hoja desde una hoja existente (Prestamos / Fuera de Uso)
' Solo se copian valores — sin fórmulas ni formatos de origen.
' ==============================================================================
Private Sub AgregarHojaDesde(wbDestino As Workbook, nombreHoja As String)
    ' Verificar que la hoja origen existe
    Dim wsOrigen As Worksheet
    On Error Resume Next
    Set wsOrigen = ThisWorkbook.Sheets(nombreHoja)
    On Error GoTo 0

    If wsOrigen Is Nothing Then
        MsgBox "No se encontró la hoja '" & nombreHoja & "'." & vbCrLf & _
               "Se omitirá en el Excel generado.", vbExclamation
        Exit Sub
    End If

    ' Determinar rango: si tiene tabla propia la usa, sino UsedRange
    Dim rngOrigen As Range
    If wsOrigen.ListObjects.Count > 0 Then
        Set rngOrigen = wsOrigen.ListObjects(1).Range  ' Incluye encabezados
    Else
        Set rngOrigen = wsOrigen.UsedRange
    End If

    If rngOrigen Is Nothing Then Exit Sub

    Dim nFilas As Long
    Dim nCols  As Long
    nFilas = rngOrigen.Rows.Count
    nCols = rngOrigen.Columns.Count

    If nFilas < 1 Then Exit Sub

    ' Agregar hoja destino y pegar solo valores
    Dim wsNueva As Worksheet
    Set wsNueva = wbDestino.Sheets.Add(After:=wbDestino.Sheets(wbDestino.Sheets.Count))
    wsNueva.Name = nombreHoja

    wsNueva.Range("A1").Resize(nFilas, nCols).Value = rngOrigen.Value
    wsNueva.Columns.AutoFit

    ' Aplicar formato tabla si hay al menos 1 fila de datos
    If nFilas > 1 Then
        Dim tbl As ListObject
        Set tbl = wsNueva.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=wsNueva.Range("A1").Resize(nFilas, nCols), _
            XlListObjectHasHeaders:=xlYes)
        tbl.Name = "Tabla_" & Replace(nombreHoja, " ", "_")
        tbl.TableStyle = "TableStyleMedium2"
    End If
End Sub

' ==============================================================================
' HELPER: Armar texto descriptivo de los filtros activos (para la advertencia)
' ==============================================================================
Private Function ConstruirDescFiltros() As String
    Dim desc As String
    desc = ""

    If Me.cmbAgencia.Value <> "" And InStr(Me.cmbAgencia.Value, "(Todas") = 0 Then
        desc = desc & "  • Agencia:   " & Me.cmbAgencia.Value & vbCrLf
    End If
    If Me.cmbEstado.Value <> "" And InStr(Me.cmbEstado.Value, "(Todos") = 0 Then
        desc = desc & "  • Estado:    " & Me.cmbEstado.Value & vbCrLf
    End If
    If Me.cmbTipo.Value <> "" And InStr(Me.cmbTipo.Value, "(Todos") = 0 Then
        desc = desc & "  • Tipo:      " & Me.cmbTipo.Value & vbCrLf
    End If
    If Me.cmbGarantia.Value <> "" And InStr(Me.cmbGarantia.Value, "(Todas") = 0 Then
        desc = desc & "  • Garantía:  " & Me.cmbGarantia.Value & vbCrLf
    End If

    If Trim(desc) = "" Then
        desc = "  (Sin filtros — inventario completo)" & vbCrLf
    End If

    ConstruirDescFiltros = desc
End Function

' ==============================================================================
' CARGA DE FILTROS DESDE CONFIG
' Todos los combos se cargan aquí. Un solo punto de mantenimiento.
' ==============================================================================
Private Sub CargarFiltrosDesdeConfig()
    Dim wsConfig As Worksheet
    Dim ultFila  As Long
    Dim arr      As Variant
    Dim k        As Long
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Config")
    On Error GoTo 0
    If wsConfig Is Nothing Then Exit Sub
    
    ' --- AGENCIAS (Columna A de Config) ---
    Me.cmbAgencia.Clear
    Me.cmbAgencia.AddItem "(Todas las Agencias)"
    ultFila = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).Row
    If ultFila > 1 Then
        arr = wsConfig.Range("A2:A" & ultFila).Value
        For k = 1 To UBound(arr, 1)
            If Trim(CStr(arr(k, 1))) <> "" Then Me.cmbAgencia.AddItem arr(k, 1)
        Next k
    End If
    Me.cmbAgencia.ListIndex = 0  ' Seleccionar "(Todas las Agencias)" por defecto
    
    ' --- ESTADOS (Columna D de Config) ---
    Me.cmbEstado.Clear
    Me.cmbEstado.AddItem "(Todos los Estados)"
    ultFila = wsConfig.Cells(wsConfig.Rows.Count, "D").End(xlUp).Row
    If ultFila > 1 Then
        arr = wsConfig.Range("D2:D" & ultFila).Value
        For k = 1 To UBound(arr, 1)
            If Trim(CStr(arr(k, 1))) <> "" Then Me.cmbEstado.AddItem arr(k, 1)
        Next k
    End If
    Me.cmbEstado.ListIndex = 0  ' Seleccionar "(Todos los Estados)" por defecto
    
    ' --- TIPOS DE EQUIPO (manual) ---
    Me.cmbTipo.Clear
    Me.cmbTipo.AddItem "(Todos los Tipos)"
    Me.cmbTipo.AddItem "Desktop"
    Me.cmbTipo.AddItem "Laptop"
    Me.cmbTipo.ListIndex = 0  ' Seleccionar "(Todos los Tipos)" por defecto
    
    ' --- GARANTÍA (opciones fijas) ---
    Me.cmbGarantia.Clear
    Me.cmbGarantia.AddItem "(Todas)"
    Me.cmbGarantia.AddItem "Garantía Activa"
    Me.cmbGarantia.AddItem "Por vencer (30 días)"
    Me.cmbGarantia.AddItem "Por vencer (60 días)"
    Me.cmbGarantia.AddItem "Por vencer (90 días)"
    Me.cmbGarantia.AddItem "Ya vencida"
    Me.cmbGarantia.ListIndex = 0  ' Seleccionar "(Todas)" por defecto
End Sub

' ==============================================================================
' BOTÓN: Cerrar - volver al formulario principal
' ==============================================================================
Private Sub btnCerrar_Click()
    Unload Me
End Sub



