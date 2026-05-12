VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormDisponibles 
   Caption         =   "EQUIPOS DISPONIBLES PARA PRESTAMO"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9345.001
   OleObjectBlob   =   "FormDisponibles.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormDisponibles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' FORMULARIO: FormDisponibles
' Muestra equipos disponibles y permite seleccionar uno para el préstamo.
' Se comunica con FormInventario llenando los campos txtNombrePRE, etc.
'
' VALIDACIONES:
'   [1] Verifica que haya un usuario cargado en FormInventario antes de
'       permitir la selección del equipo.
'   [2] El equipo seleccionado se elimina del ListBox inmediatamente después
'       de ser asignado, para evitar asignarlo dos veces en la misma sesión.
' ==============================================================================

Private datosCompletos() As Variant  ' Array en memoria para filtrado rápido


' ==============================================================================
' INICIALIZACIÓN
' ==============================================================================

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 1  ' Centrado en pantalla
    Call CargarDatosEnMemoria
    Call AplicarFiltro("")
End Sub


' ==============================================================================
' CARGA DE DATOS EN MEMORIA
' Una sola lectura de toda la tabla para filtrado rápido sin tocar la hoja.
' ==============================================================================

Private Sub CargarDatosEnMemoria()
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Sheets("Inventario")
    Set tbl = Nothing
    On Error Resume Next
    Set tbl = ws.ListObjects("TablaInventario")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "No se encontró TablaInventario.", vbCritical
        Unload Me
        Exit Sub
    End If
    
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "La tabla de inventario está vacía.", vbCritical
        Unload Me
        Exit Sub
    End If
    
    datosCompletos = tbl.DataBodyRange.Value
End Sub


' ==============================================================================
' FILTRADO EN MEMORIA
' Aplica el criterio de búsqueda sobre el array sin releer la hoja.
' ==============================================================================

Private Sub AplicarFiltro(ByVal criterio As String)
    Dim i As Long
    Dim sEstado As String
    Dim sUsuario As String
    Dim sNombre As String
    Dim sMarca As String
    Dim sSerie As String
    Dim sTipo As String
    Dim sCriterio As String
    Dim conteo As Long
    
    sCriterio = UCase(Trim(criterio))
    conteo = 0
    
    Me.lstDisponibles.Clear
    
    If Not IsArray(datosCompletos) Then Exit Sub
    
    For i = 1 To UBound(datosCompletos, 1)
        sEstado = UCase(Trim(CStr(datosCompletos(i, 9))))
        sUsuario = UCase(Trim(CStr(datosCompletos(i, 5))))
        
        ' Solo equipos disponibles por estado O por nombre de usuario
        If sEstado = "DISPONIBLE" Or InStr(sUsuario, "EQUIPO DISPONIBLE") > 0 Then
            
            sNombre = UCase(Trim(CStr(datosCompletos(i, 4))))
            sMarca = UCase(Trim(CStr(datosCompletos(i, 7))))
            sSerie = UCase(Trim(CStr(datosCompletos(i, 10))))
            sTipo = UCase(Trim(CStr(datosCompletos(i, 6))))
            
            ' Aplicar filtro de búsqueda si hay criterio escrito
            If sCriterio = "" Or InStr(sNombre, sCriterio) > 0 Or _
               InStr(sMarca, sCriterio) > 0 Or InStr(sSerie, sCriterio) > 0 Or _
               InStr(sTipo, sCriterio) > 0 Then
                
                With Me.lstDisponibles
                    .AddItem Trim(CStr(datosCompletos(i, 1)))         ' Código
                    .List(.ListCount - 1, 1) = sNombre                ' Nombre
                    .List(.ListCount - 1, 2) = sTipo                  ' Tipo
                    .List(.ListCount - 1, 3) = sMarca                 ' Marca
                    .List(.ListCount - 1, 4) = Trim(CStr(datosCompletos(i, 10))) ' Serie
                End With
                conteo = conteo + 1
            End If
        End If
    Next i
    
    Me.lblConteo.Caption = conteo & " equipo(s) disponible(s)"
End Sub


' ==============================================================================
' FILTRO EN TIEMPO REAL - Se activa con cada tecla escrita en txtFiltroDisponibles
' ==============================================================================

Private Sub txtFiltroDisponibles_Change()
    Call AplicarFiltro(Me.txtFiltroDisponibles.Value)
End Sub


' ==============================================================================
' DOBLE CLIC EN LISTA - Atajo para seleccionar sin usar el botón
' ==============================================================================

Private Sub lstDisponibles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call btnSeleccionar_Click
End Sub


' ==============================================================================
' BOTÓN SELECCIONAR
'
' VALIDACIÓN 1: Verifica que haya un usuario cargado en FormInventario.
'               Si txtUsuarioVista está vacío, no permite continuar.
'
' VALIDACIÓN 2: Verifica en tiempo real que el equipo sigue disponible
'               en la tabla (pudo cambiar desde que se cargó la lista).
'
' Al confirmar: llena los campos PRE en FormInventario y elimina el equipo
' del ListBox para que no pueda asignarse dos veces en la misma sesión.
' ==============================================================================

Private Sub btnSeleccionar_Click()
    
    ' --- VALIDACIÓN 1: Debe haber un usuario cargado en préstamos ---
    ' [MEJORA 1] Sin usuario no tiene sentido asignar un equipo prestado.
    If Trim(frmGeneralCorregido.txtUsuarioVista.Value) = "" Then
        MsgBox "Primero debes buscar y seleccionar un usuario en la pestaña Préstamos." & _
               vbCrLf & vbCrLf & _
               "Usa el buscador de la pestaña Préstamos, selecciona al usuario afectado" & _
               vbCrLf & "y luego vuelve aquí para asignar el equipo.", _
               vbExclamation, "Usuario No Seleccionado"
        Exit Sub
    End If
    
    ' --- Verificar que hay un equipo seleccionado en la lista ---
    If Me.lstDisponibles.ListIndex = -1 Then
        MsgBox "Selecciona un equipo de la lista antes de continuar.", _
               vbExclamation, "Sin Selección"
        Exit Sub
    End If
    
    Dim indiceSeleccionado As Long
    Dim serieSeleccionada As String
    indiceSeleccionado = Me.lstDisponibles.ListIndex
    serieSeleccionada = Me.lstDisponibles.List(indiceSeleccionado, 4)
    
    ' --- VALIDACIÓN 2: Verificar que el equipo sigue disponible en la tabla ---
    If Not EquipoSigueDisponible(serieSeleccionada) Then
        MsgBox "El equipo con serie " & serieSeleccionada & " ya no está disponible." & _
               vbCrLf & "La lista será actualizada automáticamente.", _
               vbExclamation, "Equipo No Disponible"
        ' Recargar datos y refrescar lista
        Call CargarDatosEnMemoria
        Call AplicarFiltro(Me.txtFiltroDisponibles.Value)
        Exit Sub
    End If
    
    ' --- Buscar fila completa en el array y enviar datos al formulario principal ---
    Dim i As Long
    Dim encontrado As Boolean
    encontrado = False
    
    For i = 1 To UBound(datosCompletos, 1)
        If UCase(Trim(CStr(datosCompletos(i, 10)))) = UCase(Trim(serieSeleccionada)) Then
            
            ' Llenar campos PRE en FormInventario
            With frmGeneralCorregido
                .txtNombrePRE.Value = Trim(CStr(datosCompletos(i, 4)))     ' Nombre
                .txtTipoPRE.Value = Trim(CStr(datosCompletos(i, 6)))       ' Tipo
                .txtMarcaPRE.Value = Trim(CStr(datosCompletos(i, 7)))      ' Marca
                .txtModeloPRE.Value = Trim(CStr(datosCompletos(i, 8)))     ' Modelo
                .txtSeriePRE.Value = Trim(CStr(datosCompletos(i, 10)))     ' Serie
                ' Campos que el usuario debe completar manualmente
                .txtFechaEntrega.Value = ""
                .txtFechaRetorno.Value = ""
                .txtTarea.Value = ""
                .txtJustificacion.Value = ""
                .chkFueraUso.Value = False
            End With
            
            encontrado = True
            Exit For
        End If
    Next i
    
    If Not encontrado Then
        MsgBox "No se encontraron los datos completos del equipo." & _
               vbCrLf & "Intenta recargar la lista.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' --- [MEJORA 2] Eliminar el equipo del ListBox inmediatamente ---
    ' Así no puede asignarse dos veces en la misma sesión sin recargar.
    Me.lstDisponibles.RemoveItem indiceSeleccionado
    
    ' Actualizar conteo en el label
    Me.lblConteo.Caption = Me.lstDisponibles.ListCount & " equipo(s) disponible(s)"
    
    MsgBox "Equipo asignado correctamente." & vbCrLf & _
           "Serie: " & serieSeleccionada & vbCrLf & vbCrLf & _
           "Completa las fechas y demás datos en la pestaña Préstamos.", _
           vbInformation, "Equipo Seleccionado"
    
    Unload Me
End Sub


' ==============================================================================
' VERIFICAR DISPONIBILIDAD EN TIEMPO REAL
' Consulta directamente la tabla para confirmar el estado actual del equipo.
' Independiente del array en memoria, que puede estar desactualizado.
' ==============================================================================

Private Function EquipoSigueDisponible(ByVal serie As String) As Boolean
    EquipoSigueDisponible = False
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim filaRel As Long
    Dim sEstado As String
    Dim sUsuario As String
    
    Set ws = ThisWorkbook.Sheets("Inventario")
    Set tbl = Nothing
    On Error Resume Next
    Set tbl = ws.ListObjects("TablaInventario")
    On Error GoTo 0
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    Set rng = tbl.ListColumns(10).DataBodyRange.Find( _
        What:=Trim(serie), LookIn:=xlValues, LookAt:=xlWhole)
    If rng Is Nothing Then Exit Function
    
    filaRel = rng.Row - tbl.HeaderRowRange.Row
    sEstado = UCase(Trim(tbl.ListRows(filaRel).Range(9).Value))
    sUsuario = UCase(Trim(tbl.ListRows(filaRel).Range(5).Value))
    
    If sEstado = "DISPONIBLE" Or InStr(sUsuario, "EQUIPO DISPONIBLE") > 0 Then
        EquipoSigueDisponible = True
    End If
End Function


' ==============================================================================
' BOTÓN CERRAR
' ==============================================================================

Private Sub btnCerrarDisponible_Click()
    Unload Me
End Sub



