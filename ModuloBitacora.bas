Attribute VB_Name = "ModuloBitacora"
Option Explicit

' ==============================================================================
' MÓDULO: ModuloBitacora
' DESCRIPCIÓN: Sistema de Bitácora de cambios y Control de Acceso Multiusuario
' VERSIÓN: 1.0
' ==============================================================================
'
' INSTRUCCIONES DE USO:
' 1. En el VBE: Archivo > Importar archivo > seleccionar este .bas
' 2. En UserForm_Initialize del formulario principal, agregar al inicio:
'       If Not VerificarAcceso() Then Unload Me : Exit Sub
' 3. En UserForm_QueryClose del formulario principal, agregar:
'       LiberarAcceso
' 4. En cada botón crítico (btnGuardar, btnEliminar, btnGuardarPrestamo), agregar:
'       Call RegistrarBitacora("ACCIÓN", codigo, "Descripción del cambio")
' ==============================================================================

' --- CONSTANTES GLOBALES ---
Private Const HOJA_BITACORA     As String = "Bitacora"
Private Const HOJA_CONFIG       As String = "Config"
Private Const CELDA_USUARIO_EN_USO As String = "Z1"  ' Celda oculta en Config para control de acceso
Private Const CELDA_TIMESTAMP   As String = "Z2"      ' Timestamp de cuando abrió
Private Const TIMEOUT_MINUTOS   As Long = 120          ' Si lleva más de 2h sin actividad, se libera automático

' ==============================================================================
' SECCIÓN 1: BITÁCORA DE CAMBIOS
' ==============================================================================

' ------------------------------------------------------------------------------
' SUB PRINCIPAL: RegistrarBitacora
' Llamar desde cualquier acción crítica del formulario.
' Parámetros:
'   accion  : "NUEVO" | "EDICIÓN" | "ELIMINAR" | "PRÉSTAMO" | "BAJA" | "IMPORTAR"
'   codigo  : El código del registro afectado (COD001, FDU003, etc.)
'   detalle : Texto libre describiendo qué cambió
' ------------------------------------------------------------------------------
Public Sub RegistrarBitacora(accion As String, codigo As String, detalle As String)
    Dim ws As Worksheet
    Dim sigFila As Long
    
    ' 1. Garantizar que la hoja Bitácora exista (se crea automáticamente si no existe)
    Set ws = ObtenerOCrearHojaBitacora()
    If ws Is Nothing Then Exit Sub  ' Si falla la creación, no bloqueamos el flujo principal
    
    ' 2. Encontrar la siguiente fila disponible
    sigFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' 3. Escribir el registro
    With ws
        .Cells(sigFila, 1).Value = Now()                    ' Fecha y Hora exacta
        .Cells(sigFila, 2).Value = ObtenerUsuarioActual()   ' Usuario del sistema
        .Cells(sigFila, 3).Value = UCase(accion)            ' Acción realizada
        .Cells(sigFila, 4).Value = codigo                   ' Código del registro
        .Cells(sigFila, 5).Value = detalle                  ' Detalle del cambio
        
        ' 4. Formato visual por tipo de acción
        Dim colorFila As Long
        Select Case UCase(accion)
            Case "NUEVO"
                colorFila = RGB(198, 239, 206)   ' Verde claro
            Case "EDICIÓN"
                colorFila = RGB(255, 235, 156)   ' Amarillo claro
            Case "ELIMINAR", "BAJA"
                colorFila = RGB(255, 199, 206)   ' Rojo claro
            Case "PRÉSTAMO"
                colorFila = RGB(189, 215, 238)   ' Azul claro
            Case "IMPORTAR"
                colorFila = RGB(226, 198, 239)   ' Morado claro
            Case Else
                colorFila = RGB(242, 242, 242)   ' Gris claro
        End Select
        
        .Range(.Cells(sigFila, 1), .Cells(sigFila, 5)).Interior.Color = colorFila
        .Cells(sigFila, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    End With
    
End Sub

' ------------------------------------------------------------------------------
' FUNCIÓN PRIVADA: ObtenerOCrearHojaBitacora
' Devuelve la hoja Bitacora. Si no existe, la crea con sus encabezados.
' ------------------------------------------------------------------------------
Private Function ObtenerOCrearHojaBitacora() As Worksheet
    Dim ws As Worksheet
    Dim wsExistente As Worksheet
    
    ' Intentamos obtener la hoja existente
    On Error Resume Next
    Set wsExistente = ThisWorkbook.Sheets(HOJA_BITACORA)
    On Error GoTo 0
    
    If Not wsExistente Is Nothing Then
        ' Ya existe, la devolvemos directamente
        Set ObtenerOCrearHojaBitacora = wsExistente
        Exit Function
    End If
    
    ' No existe: la creamos al final del libro
    On Error GoTo ErrorCreacion
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = HOJA_BITACORA
    
    ' Encabezados
    With ws
        .Cells(1, 1).Value = "Fecha / Hora"
        .Cells(1, 2).Value = "Usuario"
        .Cells(1, 3).Value = "Acción"
        .Cells(1, 4).Value = "Código Registro"
        .Cells(1, 5).Value = "Detalle del Cambio"
        
        ' Formato del encabezado
        With .Range("A1:E1")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)   ' Azul oscuro corporativo
            .HorizontalAlignment = xlCenter
        End With
        
        ' Ancho de columnas
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 16
        .Columns("E").ColumnWidth = 50
        
        ' Proteger la hoja para que nadie borre registros accidentalmente
        ' (sin contraseña para no complicar, solo dificulta edición accidental)
        .Protect UserInterfaceOnly:=True
    End With
    
    Set ObtenerOCrearHojaBitacora = ws
    Exit Function
    
ErrorCreacion:
    ' Si falla la creación, no interrumpimos el flujo principal del sistema
    Set ObtenerOCrearHojaBitacora = Nothing
End Function

' ------------------------------------------------------------------------------
' FUNCIÓN PRIVADA: ObtenerUsuarioActual
' Obtiene el nombre del usuario de Windows. Si falla, usa "DESCONOCIDO".
' ------------------------------------------------------------------------------
Private Function ObtenerUsuarioActual() As String
    Dim usuario As String
    On Error Resume Next
    usuario = Environ("USERNAME")
    On Error GoTo 0
    
    If Trim(usuario) = "" Then
        usuario = "DESCONOCIDO"
    End If
    
    ObtenerUsuarioActual = usuario
End Function


' ==============================================================================
' SECCIÓN 2: CONTROL DE ACCESO MULTIUSUARIO
' ==============================================================================
'
' CÓMO FUNCIONA:
' - Al abrir el formulario, se escribe en la celda Z1 de Config:
'   "USUARIO_WINDOWS | dd/mm/yyyy hh:mm"
' - Si otro usuario abre y Z1 ya tiene un valor, recibe un aviso con
'   el nombre de quién lo tiene abierto y desde cuándo.
' - El comportamiento es: AVISAR y dejar entrar (modo informativo).
' - Si el usuario que "bloqueó" lleva más de TIMEOUT_MINUTOS sin actividad,
'   se libera automáticamente (protección ante cierres forzados/cortes de luz).
' ==============================================================================

' ------------------------------------------------------------------------------
' FUNCIÓN PRINCIPAL: VerificarAcceso
' Retorna TRUE si puede entrar, FALSE si debe bloquearse.
' Llamar al inicio de UserForm_Initialize.
' ------------------------------------------------------------------------------
Public Function VerificarAcceso() As Boolean
    Dim wsConfig As Worksheet
    Dim valorActual As String
    Dim usuarioEnUso As String
    Dim timestampStr As String
    Dim timestampApertura As Date
    Dim minutosTranscurridos As Long
    Dim usuarioActual As String
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets(HOJA_CONFIG)
    On Error GoTo 0
    
    If wsConfig Is Nothing Then
        ' Si no hay hoja Config, dejamos entrar sin control
        VerificarAcceso = True
        Exit Function
    End If
    
    usuarioActual = ObtenerUsuarioActual()
    valorActual = Trim(wsConfig.Range(CELDA_USUARIO_EN_USO).Value)
    timestampStr = Trim(wsConfig.Range(CELDA_TIMESTAMP).Value)
    
    ' Si la celda está vacía, nadie lo tiene abierto → registramos y entramos
    If valorActual = "" Then
        Call MarcarArchivoEnUso(wsConfig, usuarioActual)
        VerificarAcceso = True
        Exit Function
    End If
    
    ' Si el mismo usuario lo tiene "abierto" (misma PC, reabrió el form), lo dejamos
    If UCase(valorActual) = UCase(usuarioActual) Then
        Call MarcarArchivoEnUso(wsConfig, usuarioActual) ' Refrescamos el timestamp
        VerificarAcceso = True
        Exit Function
    End If
    
    ' Verificar timeout: si lleva más de TIMEOUT_MINUTOS, liberamos automáticamente
    If IsDate(timestampStr) Then
        timestampApertura = CDate(timestampStr)
        minutosTranscurridos = DateDiff("n", timestampApertura, Now())
        
        If minutosTranscurridos > TIMEOUT_MINUTOS Then
            ' Timeout alcanzado: liberamos y dejamos entrar
            Call MarcarArchivoEnUso(wsConfig, usuarioActual)
            VerificarAcceso = True
            Exit Function
        End If
    End If
    
    ' Otro usuario lo tiene abierto y está dentro del timeout
    ' COMPORTAMIENTO: Avisar pero dejar entrar (modo informativo)
    Dim mensaje As String
    mensaje = "⚠️ AVISO: El sistema ya está siendo utilizado por:" & vbCrLf & vbCrLf & _
              "   Usuario: " & valorActual & vbCrLf & _
              "   Desde:   " & timestampStr & vbCrLf & vbCrLf & _
              "Podés entrar, pero TEN CUIDADO:" & vbCrLf & _
              "Si ambos guardan al mismo tiempo en OneDrive," & vbCrLf & _
              "uno podría sobreescribir los cambios del otro." & vbCrLf & vbCrLf & _
              "¿Deseás entrar de todas formas?"
    
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox(mensaje, vbYesNo + vbExclamation, "Sistema en Uso")
    
    If respuesta = vbYes Then
        Call MarcarArchivoEnUso(wsConfig, usuarioActual)
        VerificarAcceso = True
    Else
        VerificarAcceso = False
    End If
    
End Function

' ------------------------------------------------------------------------------
' SUB: MarcarArchivoEnUso
' Escribe el usuario actual y timestamp en las celdas de control.
' ------------------------------------------------------------------------------
Private Sub MarcarArchivoEnUso(wsConfig As Worksheet, usuario As String)
    On Error Resume Next
    wsConfig.Range(CELDA_USUARIO_EN_USO).Value = usuario
    wsConfig.Range(CELDA_TIMESTAMP).Value = Format(Now(), "dd/mm/yyyy hh:mm:ss")
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' SUB PÚBLICO: LiberarAcceso
' Limpiar las celdas de control al cerrar el formulario.
' Llamar en UserForm_QueryClose.
' ------------------------------------------------------------------------------
Public Sub LiberarAcceso()
    Dim wsConfig As Worksheet
    Dim usuarioActual As String
    Dim usuarioRegistrado As String
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets(HOJA_CONFIG)
    On Error GoTo 0
    
    If wsConfig Is Nothing Then Exit Sub
    
    usuarioActual = ObtenerUsuarioActual()
    usuarioRegistrado = Trim(wsConfig.Range(CELDA_USUARIO_EN_USO).Value)
    
    ' Solo limpiamos si somos nosotros los que estamos registrados
    ' (Evita que un segundo usuario limpie el registro del primero al cerrar)
    If UCase(usuarioActual) = UCase(usuarioRegistrado) Then
        On Error Resume Next
        wsConfig.Range(CELDA_USUARIO_EN_USO).Value = ""
        wsConfig.Range(CELDA_TIMESTAMP).Value = ""
        On Error GoTo 0
    End If
End Sub

' ------------------------------------------------------------------------------
' SUB PÚBLICO: RefrescarTimestamp
' Llamar periódicamente o en acciones del usuario para renovar el timeout.
' Opcional: llamar desde btnGuardar_Click para indicar actividad reciente.
' ------------------------------------------------------------------------------
Public Sub RefrescarTimestamp()
    Dim wsConfig As Worksheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets(HOJA_CONFIG)
    If Not wsConfig Is Nothing Then
        wsConfig.Range(CELDA_TIMESTAMP).Value = Format(Now(), "dd/mm/yyyy hh:mm:ss")
    End If
    On Error GoTo 0
End Sub
