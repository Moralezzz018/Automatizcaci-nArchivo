VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} INVENTARIOCRUD 
   Caption         =   "UserForm1"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17760
   OleObjectBlob   =   "INVENTARIOCRUD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "INVENTARIOCRUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bthNuevoArticulo_Click()
AGGARTICULO.Show
End Sub

Private Sub btnAgregarStock_Click()
Dim ws As Worksheet
    Dim fila As Variant
    Dim cantidadIngresada As Double, cantidadNormalizada As Double
    Dim unidadIngresada As String, unidadProducto As String
    Dim totalActual As Double, capacidad As Double
    Dim nuevoTotal As Double, nuevaCantidadEnvases As Double

    If txtID.Value = "" Then MsgBox "Seleccione un reactivo primero", vbExclamation: Exit Sub

    ' 1. Pedir datos
    cantidadIngresada = Val(InputBox("Ingrese cantidad que LLEGÓ al inventario:", "Entrada de Stock"))
    If cantidadIngresada <= 0 Then Exit Sub
    
    unidadIngresada = InputBox("Ingrese unidad (g, kg, ml, L):", "Unidad de entrada")
    If unidadIngresada = "" Then Exit Sub

    Set ws = Hoja1
    fila = Application.Match(txtID.Value, ws.Columns(1), 0)
    If IsError(fila) Then Exit Sub

    ' 2. Leer datos actuales
    unidadProducto = ws.Cells(fila, 7).Value
    totalActual = ws.Cells(fila, 11).Value
    capacidad = ws.Cells(fila, 10).Value

    ' 3. Convertir (Usamos la misma función que ya tienes para Rebajar)
    cantidadNormalizada = ConvertirUnidades(cantidadIngresada, unidadIngresada, unidadProducto)

    If cantidadNormalizada = -1 Then
        MsgBox "No se puede sumar " & unidadIngresada & " con " & unidadProducto, vbCritical
        Exit Sub
    End If

    ' 4. SUMAR (Aquí está la diferencia con Rebajar)
    nuevoTotal = totalActual + cantidadNormalizada
    nuevaCantidadEnvases = nuevoTotal / capacidad

    ' 5. Guardar en Excel
    ws.Cells(fila, 11).Value = nuevoTotal
    ws.Cells(fila, 3).Value = Round(nuevaCantidadEnvases, 2)
    
    ' Opcional: Actualizar fecha de pedido a HOY porque acabó de llegar
    ws.Cells(fila, 5).Value = Date

    ' 6. Actualizar Estado (Si estaba vencido o sin stock, ahora cambia)
    ws.Cells(fila, 8).Value = CalcularEstado(ws.Cells(fila, 4).Value)

    CargarDatos
    MsgBox "Stock agregado. Nuevo total: " & nuevoTotal & " " & unidadProducto, vbInformation
End Sub

Private Sub btnLimpiar_Click()
' 1. Borrar los textos de las cajas
    LimpiarCampos
    
    ' 2. Borrar el buscador (para que aparezca toda la lista de nuevo)
    txtBuscar.Value = ""
    
    ' 3. Recargar la lista original desde Excel
    CargarDatos
    
    ' 4. Poner el cursor listo para escribir un nuevo reactivo
    On Error Resume Next
    txtReactivo.SetFocus
    On Error GoTo 0
End Sub

'========================================
' RUTINA AUXILIAR PARA BORRAR CAJAS
' (Asegúrate de tener esta parte también)
'========================================

Private Sub btnReporte_Click()
Dim wbNuevo As Workbook
    Dim wsTemp As Worksheet
    Dim ruta As Variant
    Dim i As Integer, j As Integer
    
    If lstInventario.ListCount = 0 Then
        MsgBox "No hay datos para exportar.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 1. Crear hoja temporal
    Set wbNuevo = Workbooks.Add
    Set wsTemp = wbNuevo.Sheets(1)
    
    ' 2. Encabezados
    Dim encabezados As Variant
    encabezados = Array("ID", "REACTIVO", "CANT", "VENCIMIENTO", "PEDIDO", "APERTURA", "UNIDAD", "ESTADO", "OBSERVACIONES", "CAPACIDAD", "TOTAL BASE", "ENVASE")
    wsTemp.Range("A1:L1").Value = encabezados
    
    With wsTemp.Range("A1:L1")
        .Interior.Color = RGB(50, 50, 50)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' 3. Volcar datos
    For i = 0 To lstInventario.ListCount - 1
        For j = 0 To 11
            wsTemp.Cells(i + 2, j + 1).Value = lstInventario.List(i, j)
        Next j
    Next i
    
    ' 4. AUTOAJUSTAR COLUMNAS (Solución punto 2)
    wsTemp.Columns("A:L").EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    
    ' 5. Guardar
    If MsgBox("¿Guardar como PDF?" & vbCrLf & "(Sí = PDF, No = Excel)", vbYesNo + vbQuestion) = vbYes Then
        ruta = Application.GetSaveAsFilename("Reporte_" & Format(Date, "dd-mm-yy") & ".pdf", "PDF Files (*.pdf), *.pdf")
        If ruta <> False Then
            
            ' CONFIGURACIÓN PARA QUE QUEPA EN UNA HOJA (Solución punto 1)
            With wsTemp.PageSetup
                .Orientation = xlLandscape      ' Horizontal
                .Zoom = False                   ' Desactivar zoom fijo
                .FitToPagesWide = 1             ' Forzar a 1 página de ancho
                .FitToPagesTall = False         ' El largo puede ser infinito
                .LeftMargin = Application.InchesToPoints(0.2)
                .RightMargin = Application.InchesToPoints(0.2)
            End With
            
            wsTemp.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ruta, OpenAfterPublish:=True
            wbNuevo.Close SaveChanges:=False
        End If
    Else
        ' Excel normal
        ruta = Application.GetSaveAsFilename("Reporte_" & Format(Date, "dd-mm-yy") & ".xlsx", "Excel Files (*.xlsx), *.xlsx")
        If ruta <> False Then
            wbNuevo.SaveAs ruta
            MsgBox "Excel guardado.", vbInformation
        Else
            wbNuevo.Close SaveChanges:=False
        End If
    End If
End Sub


'========================================
' INICIALIZAR FORMULARIO
'========================================
Private Sub UserForm_Initialize()
    ' 1. Configuración Visual
    Me.StartUpPosition = 2 ' Centrar
    
    ' 2. Verificar si la hoja está vacía y crear encabezados
    VerificarEncabezados
    
    ' 3. Cargar Combos
    CargarCombos
    
    ' 4. Configurar ListBox (Ahora 12 columnas por la fecha extra)
    With lstInventario
        .ColumnCount = 12
        ' Ajusta estos anchos según tu preferencia visual
        .ColumnWidths = "55;150;30;65;65;70;50;80;150;50;50;60"
    End With

    ' 5. Bloquear campos automáticos
    txtID.Enabled = False
    txtEstado.Enabled = False
    
    ' 6. Cargar datos a la lista
    CargarDatos
End Sub

'========================================
' AUTO-CREAR ENCABEZADOS (NUEVO)
'========================================
Sub VerificarEncabezados()
    Dim ws As Worksheet
    Set ws = Hoja1
    
    ' Si la celda A1 está vacía, creamos la estructura
    If ws.Range("A1").Value = "" Then
        Dim encabezados As Variant
        ' Estructura de 12 Columnas
        encabezados = Array("ID", "REACTIVO", "CANTIDAD", "FECHA VENCIMIENTO", "FECHA PEDIDO", "FECHA APERTURA", "UNIDAD", "ESTADO", "OBSERVACIONES", "CAPACIDAD", "TOTAL BASE", "ENVASE")
        
        ws.Range("A1:L1").Value = encabezados
        
        ' Formato profesional
        With ws.Range("A1:L1")
            .Font.Bold = True
            .Interior.Color = RGB(0, 51, 102) ' Azul oscuro
            .Font.Color = vbWhite
            .HorizontalAlignment = xlCenter
        End With
    End If
End Sub

'========================================
' BOTÓN: GUARDAR (CON ID INTELIGENTE)
'========================================
Private Sub btnGuardar_Click()
    Dim ws As Worksheet
    Dim nuevaFila As Long
    Dim estado As String
    Dim capacidad As Double
    Dim totalBase As Double
    Dim cantidadEnvases As Double

    Set ws = Hoja1
    
    ' --- VALIDACIONES ---
    If txtReactivo.Value = "" Then MsgBox "Ingrese el reactivo", vbExclamation: Exit Sub
    If Not IsNumeric(txtCantidad.Value) Or Val(txtCantidad.Value) <= 0 Then MsgBox "Cantidad inválida", vbExclamation: Exit Sub
    If Not IsDate(txtFechaVencimiento.Value) Then MsgBox "Fecha vencimiento inválida", vbExclamation: Exit Sub
    If cmbUnidad.Value = "" Then MsgBox "Seleccione unidad", vbExclamation: Exit Sub

    ' Pedir capacidad para el cálculo base
    capacidad = Val(InputBox("Ingrese la capacidad por envase (Ej: 500 para 500ml)", "Capacidad del Envase"))
    If capacidad <= 0 Then MsgBox "Capacidad inválida", vbExclamation: Exit Sub

    ' Cálculos matemáticos
    cantidadEnvases = CDbl(txtCantidad.Value)
    totalBase = cantidadEnvases * capacidad
    estado = CalcularEstado(CDate(txtFechaVencimiento.Value))

    ' --- GENERAR NUEVO ID (RCT 001) ---
    txtID.Value = GenerarCodigoInteligente("RCT")

    ' --- GUARDAR EN EXCEL ---
    nuevaFila = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    
    ws.Cells(nuevaFila, 1).Value = txtID.Value
    ws.Cells(nuevaFila, 2).Value = txtReactivo.Value
    ws.Cells(nuevaFila, 3).Value = cantidadEnvases
    ws.Cells(nuevaFila, 4).Value = CDate(txtFechaVencimiento.Value)
    ws.Cells(nuevaFila, 5).Value = txtFechaPedido.Value
    ws.Cells(nuevaFila, 6).Value = txtFechaApertura.Value ' Nuevo Campo
    ws.Cells(nuevaFila, 7).Value = cmbUnidad.Value
    ws.Cells(nuevaFila, 8).Value = estado
    ws.Cells(nuevaFila, 9).Value = txtObservaciones.Value
    ws.Cells(nuevaFila, 10).Value = capacidad
    ws.Cells(nuevaFila, 11).Value = totalBase
    ws.Cells(nuevaFila, 12).Value = cmbEnvase.Value

    MsgBox "Reactivo registrado con éxito: " & txtID.Value, vbInformation
    
    CargarDatos
    LimpiarCampos
End Sub

'========================================
' BOTÓN: ACTUALIZAR (CORREGIDO)
'========================================
Private Sub btnActualizar_Click()
    Dim ws As Worksheet
    Dim fila As Variant
    Dim capacidadUnit As Double
    Dim nuevoTotalBase As Double

    If txtID.Value = "" Then Exit Sub
    Set ws = Hoja1

    ' Buscar por texto exacto (RCT 001)
    fila = Application.Match(txtID.Value, ws.Columns(1), 0)

    If IsError(fila) Then
        MsgBox "Registro no encontrado. Verifique el ID.", vbExclamation
        Exit Sub
    End If

    ' --- LÓGICA DE RECALCULO ---
    ' Leemos la capacidad original para no perder la referencia matemática
    capacidadUnit = ws.Cells(fila, 10).Value ' Columna 10 es Capacidad
    
    ' Si por error la capacidad en celda es 0, pedimos confirmación (seguridad)
    If capacidadUnit = 0 Then
        capacidadUnit = Val(InputBox("El sistema no detecta capacidad previa. Ingrésela:", "Corrección"))
        ws.Cells(fila, 10).Value = capacidadUnit
    End If

    nuevoTotalBase = CDbl(txtCantidad.Value) * capacidadUnit

    ' Actualizar Celdas
    ws.Cells(fila, 2).Value = txtReactivo.Value
    ws.Cells(fila, 3).Value = CDbl(txtCantidad.Value)
    ws.Cells(fila, 4).Value = CDate(txtFechaVencimiento.Value)
    ws.Cells(fila, 5).Value = txtFechaPedido.Value
    ws.Cells(fila, 6).Value = txtFechaApertura.Value
    ws.Cells(fila, 7).Value = cmbUnidad.Value
    ws.Cells(fila, 8).Value = CalcularEstado(CDate(txtFechaVencimiento.Value)) ' Recalcular estado
    ws.Cells(fila, 9).Value = txtObservaciones.Value
    ' La columna 10 (Capacidad) no se toca a menos que sea necesario
    ws.Cells(fila, 11).Value = nuevoTotalBase ' Actualizamos el stock total real
    ws.Cells(fila, 12).Value = cmbEnvase.Value

    CargarDatos
    MsgBox "Registro actualizado y stock recalculado.", vbInformation
End Sub

'========================================
' BOTÓN: ELIMINAR
'========================================
Private Sub btnEliminar_Click()
    Dim ws As Worksheet
    Dim fila As Variant

    If txtID.Value = "" Then Exit Sub
    If MsgBox("¿Está seguro de eliminar " & txtID.Value & "?", vbYesNo + vbCritical) = vbNo Then Exit Sub

    Set ws = Hoja1
    fila = Application.Match(txtID.Value, ws.Columns(1), 0)

    If IsError(fila) Then Exit Sub
    
    ws.Rows(fila).Delete
    CargarDatos
    LimpiarCampos
    MsgBox "Registro eliminado.", vbInformation
End Sub

'========================================
' BOTÓN: REBAJAR INVENTARIO
'========================================
Private Sub btnRebajar_Click()
    Dim ws As Worksheet
    Dim fila As Variant
    Dim cantidadIngresada As Double, cantidadNormalizada As Double
    Dim unidadIngresada As String, unidadProducto As String
    Dim totalActual As Double, capacidad As Double
    Dim nuevoTotal As Double, nuevaCantidadEnvases As Double

    If txtID.Value = "" Then MsgBox "Seleccione un reactivo", vbExclamation: Exit Sub

    ' 1. Obtener datos ingresados
    cantidadIngresada = Val(InputBox("Ingrese cantidad a consumir:", "Rebaja de Stock"))
    If cantidadIngresada <= 0 Then Exit Sub
    
    unidadIngresada = InputBox("Ingrese unidad (g, kg, ml, L):", "Unidad de consumo")
    If unidadIngresada = "" Then Exit Sub

    Set ws = Hoja1
    fila = Application.Match(txtID.Value, ws.Columns(1), 0)
    If IsError(fila) Then Exit Sub

    ' 2. Leer datos del Excel
    unidadProducto = ws.Cells(fila, 7).Value   ' La unidad en la que está guardado (ej: kg)
    totalActual = ws.Cells(fila, 11).Value     ' El Total Base actual (ej: 1)
    capacidad = ws.Cells(fila, 10).Value       ' Capacidad del envase

    ' 3. CONVERSIÓN DE UNIDADES (El corazón de la corrección)
    ' Convertimos lo que ingresaste a la unidad del producto
    cantidadNormalizada = ConvertirUnidades(cantidadIngresada, unidadIngresada, unidadProducto)

    If cantidadNormalizada = -1 Then
        MsgBox "No se puede convertir de " & unidadIngresada & " a " & unidadProducto, vbCritical
        Exit Sub
    End If

    ' 4. Validar Stock
    If cantidadNormalizada > totalActual Then
        MsgBox "Stock insuficiente." & vbCrLf & _
               "Tienes: " & totalActual & " " & unidadProducto & vbCrLf & _
               "Intentas sacar: " & cantidadNormalizada & " " & unidadProducto, vbCritical
        Exit Sub
    End If

    ' 5. Cálculos Finales
    nuevoTotal = totalActual - cantidadNormalizada
    nuevaCantidadEnvases = nuevoTotal / capacidad

    ' 6. Guardar y Actualizar
    ws.Cells(fila, 11).Value = nuevoTotal
    ws.Cells(fila, 3).Value = Round(nuevaCantidadEnvases, 2)

    ' Actualizar Estado
    If nuevoTotal <= 0 Then
        ws.Cells(fila, 8).Value = "SIN EXISTENCIA"
    ElseIf nuevaCantidadEnvases < 1 Then
        ws.Cells(fila, 8).Value = "STOCK BAJO"
        MsgBox "ALERTA: Queda menos de 1 envase.", vbExclamation
    Else
        ws.Cells(fila, 8).Value = CalcularEstado(ws.Cells(fila, 4).Value)
    End If

    CargarDatos
    MsgBox "Rebaja aplicada. Restan: " & nuevoTotal & " " & unidadProducto, vbInformation
End Sub

' Función auxiliar necesaria para que la rebaja funcione
Function ConvertirUnidades(cant As Double, uOrigen As String, uDestino As String) As Double
    uOrigen = LCase(Trim(uOrigen))
    uDestino = LCase(Trim(uDestino))
    
    If uOrigen = uDestino Then
        ConvertirUnidades = cant
        Exit Function
    End If
    
    ' Lógica cruzada
    If uOrigen = "g" And uDestino = "kg" Then ConvertirUnidades = cant / 1000
    If uOrigen = "kg" And uDestino = "g" Then ConvertirUnidades = cant * 1000
    If uOrigen = "ml" And uDestino = "l" Then ConvertirUnidades = cant / 1000
    If uOrigen = "l" And uDestino = "ml" Then ConvertirUnidades = cant * 1000
    ' Si no cae en ninguno de estos, error (-1)
    If uOrigen <> "g" And uOrigen <> "kg" And uOrigen <> "ml" And uOrigen <> "l" Then ConvertirUnidades = -1
End Function

'========================================
' CIERRE SEMANAL (CORREGIDO Y ROBUSTO)
'========================================
Private Sub btnCierreSemanal_Click()
    Dim nombreCierre As String
    Dim wsActual As Worksheet
    Dim wsCheck As Worksheet ' Variable auxiliar para verificar
    
    ' 1. Pedir nombre
    nombreCierre = InputBox("Nombre para el cierre:" & vbCrLf & _
                            "(Ej: Semana 24-Feb)", "Cierre de Inventario")
    
    If Trim(nombreCierre) = "" Then Exit Sub
    
    ' Limpiar caracteres prohibidos
    nombreCierre = Replace(nombreCierre, "/", "-")
    nombreCierre = Replace(nombreCierre, "\", "-")
    nombreCierre = Replace(nombreCierre, ":", "")
    
    ' 2. VERIFICACIÓN ROBUSTA (Aquí estaba el error)
    Set wsCheck = Nothing
    On Error Resume Next
    Set wsCheck = Worksheets(nombreCierre) ' Intenta "agarrar" la hoja
    On Error GoTo 0
    
    ' Si logró "agarrar" la hoja (no es Nothing), es que YA EXISTE
    If Not wsCheck Is Nothing Then
        MsgBox "Error: Ya existe una hoja llamada '" & nombreCierre & "' en este archivo (quizás está oculta).", vbCritical
        Exit Sub
    End If
    
    ' 3. Copiar hoja
    Application.ScreenUpdating = False
    Set wsActual = Hoja1 ' Asegúrate que Hoja1 sea tu inventario
    
    wsActual.Copy After:=Worksheets(Worksheets.count)
    
    With ActiveSheet
        .Name = nombreCierre
        .Tab.Color = vbRed
        ' Opcional: Proteger y romper vínculos con botones
        On Error Resume Next
        .Shapes("btnGuardar").Delete
        .Shapes("btnActualizar").Delete
        .Shapes("btnEliminar").Delete
        On Error GoTo 0
    End With
    
    wsActual.Activate
    Application.ScreenUpdating = True
    
    MsgBox "Cierre generado exitosamente: " & nombreCierre, vbInformation
End Sub

'========================================
' BOTÓN: BUSCAR (FILTRO ARRAY)
'========================================
Private Sub txtBuscar_Change()
    Dim ws As Worksheet
    Dim datosOriginales As Variant
    Dim datosFiltrados() As Variant
    Dim i As Long, j As Long, count As Long
    Dim texto As String, ultFila As Long
    
    Set ws = Hoja1
    texto = LCase(Trim(txtBuscar.Value))
    ultFila = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    If ultFila < 2 Then Exit Sub
    
    ' Si está vacío, cargar todo normal
    If texto = "" Then
        CargarDatos
        Exit Sub
    End If
    
    datosOriginales = ws.Range("A2:L" & ultFila).Value
    
    ' 1. Contar coincidencias primero
    count = 0
    For i = 1 To UBound(datosOriginales, 1)
        If InStr(LCase(datosOriginales(i, 1)), texto) > 0 Or _
           InStr(LCase(datosOriginales(i, 2)), texto) > 0 Then
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        lstInventario.Clear
        Exit Sub
    End If
    
    ' 2. Crear matriz del tamaño exacto (Filas, Columnas)
    ReDim datosFiltrados(0 To count - 1, 0 To 11)
    
    ' 3. Llenar la matriz
    Dim filaActual As Long
    filaActual = 0
    
    For i = 1 To UBound(datosOriginales, 1)
        If InStr(LCase(datosOriginales(i, 1)), texto) > 0 Or _
           InStr(LCase(datosOriginales(i, 2)), texto) > 0 Then
            
            For j = 0 To 11
                datosFiltrados(filaActual, j) = datosOriginales(i, j + 1)
            Next j
            filaActual = filaActual + 1
        End If
    Next i
    
    ' 4. Asignar directamente (sin Transpose)
    lstInventario.ColumnWidths = "50;150;50;60;60;60;40;80;150;50;50;60" ' Reafirmar anchos
    lstInventario.List = datosFiltrados
End Sub

'========================================
' GENERAR CÓDIGO INTELIGENTE (RCT 001)
'========================================
Function GenerarCodigoInteligente(prefijo As String) As String
    Dim ws As Worksheet
    Dim i As Long, ultimaFila As Long
    Dim maxNum As Long, numActual As Long
    Dim valorCelda As String
    
    Set ws = Hoja1
    ultimaFila = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    maxNum = 0
    
    If ultimaFila < 2 Then
        GenerarCodigoInteligente = prefijo & " 001"
        Exit Function
    End If
    
    For i = 2 To ultimaFila
        valorCelda = ws.Cells(i, 1).Value
        If UCase(Left(valorCelda, 3)) = UCase(prefijo) Then
            ' Extrae los ultimos 3 digitos
            numActual = Val(Right(valorCelda, 3))
            If numActual > maxNum Then maxNum = numActual
        End If
    Next i
    
    GenerarCodigoInteligente = prefijo & " " & Format(maxNum + 1, "000")
End Function

'========================================
' CARGAR DATOS AL LISTBOX
'========================================
Sub CargarDatos()
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim matrizDatos() As Variant

    Set ws = Hoja1
    ultimaFila = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    If ultimaFila < 2 Then
        lstInventario.Clear
        Exit Sub
    End If

    ' Matriz dinámica para 12 columnas (0 a 11)
    ReDim matrizDatos(0 To ultimaFila - 2, 0 To 11)
    Dim i As Long, j As Long

    For i = 2 To ultimaFila
        For j = 0 To 11
            matrizDatos(i - 2, j) = ws.Cells(i, j + 1).Value
        Next j
    Next i

    lstInventario.Clear
    lstInventario.List = matrizDatos
End Sub

'========================================
' EVENTOS DE FECHAS (AUTOCOMPLETAR)
'========================================
Private Sub txtFechaVencimiento_Change()
    AplicarFormatoFecha txtFechaVencimiento
End Sub
Private Sub txtFechaPedido_Change()
    AplicarFormatoFecha txtFechaPedido
End Sub
Private Sub txtFechaApertura_Change()
    AplicarFormatoFecha txtFechaApertura
End Sub

Sub AplicarFormatoFecha(txt As MSForms.TextBox)
    Dim texto As String
    Static bloqueado As Boolean
    If bloqueado Then Exit Sub
    bloqueado = True
    
    texto = Replace(txt.Text, "//", "/")
    If Len(texto) = 2 Or Len(texto) = 5 Then
        If Right(texto, 1) <> "/" Then
            txt.Text = texto & "/"
            txt.SelStart = Len(txt.Text)
        End If
    End If
    bloqueado = False
End Sub

'========================================
' SELECCIONAR ITEM DEL LISTBOX
'========================================
Private Sub lstInventario_Click()
    If lstInventario.ListIndex = -1 Then Exit Sub

    ' Mapeo de columnas a TextBoxes
    txtID.Value = lstInventario.List(lstInventario.ListIndex, 0)
    txtReactivo.Value = lstInventario.List(lstInventario.ListIndex, 1)
    txtCantidad.Value = lstInventario.List(lstInventario.ListIndex, 2)
    txtFechaVencimiento.Value = lstInventario.List(lstInventario.ListIndex, 3)
    txtFechaPedido.Value = lstInventario.List(lstInventario.ListIndex, 4)
    txtFechaApertura.Value = lstInventario.List(lstInventario.ListIndex, 5)
    
    AsignarComboSeguro cmbUnidad, lstInventario.List(lstInventario.ListIndex, 6)
    txtEstado.Value = lstInventario.List(lstInventario.ListIndex, 7)
    txtObservaciones.Value = lstInventario.List(lstInventario.ListIndex, 8)
    txtCantidadBase.Value = lstInventario.List(lstInventario.ListIndex, 10)
    
    ' Campos ocultos o secundarios
    ' Columna 9 es Capacidad, Col 10 es TotalBase, Col 11 es Envase
    AsignarComboSeguro cmbEnvase, lstInventario.List(lstInventario.ListIndex, 11)
    
    MostrarAlerta txtEstado.Value
End Sub

'========================================
' FUNCIONES AUXILIARES
'========================================
Sub CargarCombos()
    cmbUnidad.Clear
    cmbUnidad.List = Array("ml", "L", "g", "kg", "lb")
    
    cmbEnvase.Clear
    cmbEnvase.List = Array("Bote", "Caja", "Paquete", "Frasco", "Saco", "Galón")
    
    cmbUnidadBase.Clear ' Se llena automático, pero inicializamos
    cmbUnidadBase.List = Array("g", "kg", "lb", "ml", "L")
End Sub

Sub LimpiarCampos()
    txtID.Value = ""
    txtReactivo.Value = ""
    txtCantidad.Value = ""
    txtFechaVencimiento.Value = ""
    txtFechaPedido.Value = ""
    txtFechaApertura.Value = ""
    txtEstado.Value = ""
    txtObservaciones.Value = ""
    cmbUnidad.Value = ""
    cmbEnvase.Value = ""
    lblAlerta.Caption = ""
End Sub

Function CalcularEstado(fechaVenc As Date) As String
    Dim dias As Long
    dias = DateDiff("d", Date, fechaVenc)
    If Date > fechaVenc Then
        CalcularEstado = "VENCIDO"
    ElseIf dias >= 0 And dias <= 3 Then
        CalcularEstado = "PROXIMO A VENCER"
    Else
        CalcularEstado = "DISPONIBLE"
    End If
End Function

Sub MostrarAlerta(estado As String)
    estado = LCase(estado)
    lblAlerta.Caption = UCase(estado)
    If InStr(estado, "vencido") > 0 Then
        lblAlerta.ForeColor = vbRed
    ElseIf InStr(estado, "proximo") > 0 Then
        lblAlerta.ForeColor = RGB(255, 140, 0)
    ElseIf InStr(estado, "stock") > 0 Then
        lblAlerta.ForeColor = RGB(255, 0, 255)
    Else
        lblAlerta.ForeColor = RGB(0, 150, 0)
    End If
End Sub

Function NormalizarTexto(texto As String) As String
    texto = LCase(Trim(texto))
    texto = Replace(texto, "á", "a")
    texto = Replace(texto, "é", "e")
    texto = Replace(texto, "í", "i")
    texto = Replace(texto, "ó", "o")
    texto = Replace(texto, "ú", "u")
    NormalizarTexto = texto
End Function

Function ConvertirAUnidadBase(valor As Double, unidad As String) As Double
    Select Case LCase(Trim(unidad))
        Case "g", "ml": ConvertirAUnidadBase = valor
        Case "kg", "l": ConvertirAUnidadBase = valor * 1000
        Case "lb": ConvertirAUnidadBase = valor * 453.592
        Case Else: ConvertirAUnidadBase = 0: MsgBox "Unidad no reconocida", vbExclamation
    End Select
End Function

Sub AsignarComboSeguro(cmb As ComboBox, valor As String)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If LCase(cmb.List(i)) = LCase(Trim(valor)) Then
            cmb.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Sub cmbUnidad_Change()
    Select Case cmbUnidad.Value
        Case "L", "ml": cmbUnidadBase.Value = "ml"
        Case "kg", "g", "lb": cmbUnidadBase.Value = "g"
        Case Else: cmbUnidadBase.Value = ""
    End Select
End Sub

Private Sub txtCantidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57, 8 ' Números y borrar
        Case 46 ' Punto
            If InStr(txtCantidad.Text, ".") > 0 Then KeyAscii = 0
        Case Else: KeyAscii = 0
    End Select
End Sub
