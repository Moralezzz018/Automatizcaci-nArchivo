VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AGGARTICULO 
   Caption         =   "UserForm2"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6885
   OleObjectBlob   =   "AGGARTICULO.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AGGARTICULO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================================
' CÓDIGO DEL FORMULARIO: AGREGAR ARTÍCULO SIMPLE
'======================================================

Private Sub UserForm_Initialize()
    ' Al abrir, generamos un ID con prefijo UTN (Utensilio)
    ' Asegúrate de que la función GenerarCodigoInteligente esté en un Módulo o accesible
    ' Si la función está en el otro formulario, tendríamos que moverla a un Módulo.
    ' POR AHORA, usaremos una lógica simple aquí mismo para no complicarte:
    
    txtID.Value = GenerarNuevoID_UTN()
    txtID.Enabled = False ' Bloqueado para que no se toque
End Sub

Private Sub btnGuardar_Click()
    Dim ws As Worksheet
    Dim siguienteFila As Long
    
    ' 1. Validar que haya nombre
    If Trim(txtNombre.Value) = "" Then
        MsgBox "Por favor, escribe el nombre del artículo.", vbExclamation
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets("Hoja1") ' Tu base de datos compartida
    
    ' 2. Buscar la siguiente fila vacía
    siguienteFila = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    
    ' 3. Guardar SOLO los datos básicos
    ' Columna A (1): ID
    ws.Cells(siguienteFila, 1).Value = txtID.Value
    
    ' Columna B (2): Nombre del Artículo
    ws.Cells(siguienteFila, 2).Value = UCase(Trim(txtNombre.Value))
    
    ' Columna C (3): Cantidad (Ponemos 1 por defecto para que no quede vacío)
    ws.Cells(siguienteFila, 3).Value = ""
    
    ' Columna G (7): Unidad (Ponemos "Und" por defecto)
    ws.Cells(siguienteFila, 7).Value = ""
    
    ' Columna H (8): Estado
    ws.Cells(siguienteFila, 8).Value = ""
    
    ' Columna I (9): Observaciones
    ws.Cells(siguienteFila, 9).Value = StrConv(Trim(txtObservaciones.Value), vbProperCase)
    
    ' Columna J y K (10 y 11): Datos numéricos base (Ponemos 1 para evitar errores de cálculo)
    ws.Cells(siguienteFila, 10).Value = ""
    ws.Cells(siguienteFila, 11).Value = ""
    
    MsgBox "Artículo guardado correctamente.", vbInformation
    
    ' 4. Limpiar y generar nuevo ID por si quieres agregar otro seguido
    txtNombre.Value = ""
    txtObservaciones.Value = ""
    txtID.Value = GenerarNuevoID_UTN()
    txtNombre.SetFocus
    
    ' Si quieres que el formulario se cierre al guardar, descomenta la siguiente línea:
    ' Unload Me
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

' --- FUNCIÓN INTERNA PARA GENERAR ID UTN ---
Function GenerarNuevoID_UTN() As String
    Dim ws As Worksheet
    Dim ultFila As Long, i As Long
    Dim maxNum As Long, numActual As Long
    Dim codigo As String
    
    Set ws = ThisWorkbook.Sheets("Hoja1")
    ultFila = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    maxNum = 0
    
    ' Recorre la columna A buscando códigos que empiecen con "UTN"
    For i = 2 To ultFila
        codigo = ws.Cells(i, 1).Value
        If Left(codigo, 3) = "UTN" Then
            numActual = Val(Mid(codigo, 5)) ' Extrae el número después de "UTN "
            If numActual > maxNum Then maxNum = numActual
        End If
    Next i
    
    ' Devuelve el siguiente número formateado (ej: UTN 005)
    GenerarNuevoID_UTN = "UTN " & Format(maxNum + 1, "000")
End Function
