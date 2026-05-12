# Aporte 2: Automatización de procesos mediante VBA y soluciones web

## Antecedentes

Durante la práctica profesional supervisada en Alimentos Maravilla de Honduras S.A., se identificó que el área de soporte e infraestructura tecnológica realizaba múltiples actividades administrativas mediante hojas de cálculo manipuladas de forma manual. Entre estas tareas se encontraban el registro de equipos, la actualización de estados de inventario, el control de préstamos temporales, la consulta de garantías, la generación de reportes y la revisión de equipos disponibles para sustitución o soporte.

Estas actividades eran propensas a inconsistencias por digitación manual, duplicidad de registros, pérdida de trazabilidad y retrasos en la atención de tickets de soporte. En una empresa industrial y alimentaria, donde los equipos informáticos apoyan operaciones administrativas, producción, distribución y control interno, la disponibilidad de información confiable sobre activos tecnológicos resulta necesaria para tomar decisiones oportunas y reducir interrupciones operativas.

Las oportunidades de mejora se identificaron a partir de la observación directa de tareas repetitivas, el análisis de solicitudes de usuarios y la revisión de tickets relacionados con inventario, préstamos de equipo, verificación de garantía y elaboración de reportes. A partir de esa necesidad, el practicante implementó automatizaciones en Microsoft Excel mediante VBA, apoyadas por formularios de usuario, tablas estructuradas y enlaces hacia servicios web externos.

## Objetivo

El objetivo técnico principal fue automatizar y estandarizar la gestión del inventario tecnológico, los préstamos de equipos y la generación de reportes, reduciendo la intervención manual del personal de IT, minimizando errores humanos en la manipulación de planillas y facilitando la consulta centralizada de información mediante formularios VBA e integraciones web de apoyo.

De forma específica, la solución buscó:

- Registrar y actualizar equipos de inventario con validaciones de datos obligatorios.
- Controlar préstamos de equipos disponibles y actualizar automáticamente sus estados.
- Generar reportes filtrados por agencia, estado, tipo de equipo y garantía.
- Exportar información a PDF o Excel para evidencia administrativa.
- Mantener una bitácora de cambios que permitiera auditar acciones críticas.
- Facilitar la consulta de garantías y tableros mediante enlaces web desde el formulario.

## Alcance

El alcance funcional de la solución abarca una automatización local en Excel desarrollada con VBA, estructurada alrededor de formularios (`UserForms`), módulos estándar y tablas de datos. Los archivos analizados corresponden a:

- `frmGeneralCorregido.frm`: formulario principal para inventario, préstamos, búsqueda, validación de fechas, gestión de equipos fuera de uso, consulta de garantías y acceso a dashboard.
- `FormDisponibles.frm`: formulario para listar, filtrar y seleccionar equipos disponibles para préstamo.
- `FormReporte.frm`: formulario de reportes con filtros combinados y exportación a PDF o Excel.
- `ModuloBitacora.bas`: módulo estándar para bitácora de cambios, control informativo de acceso multiusuario y actualización de actividad.

La solución beneficia principalmente al departamento de IT o soporte técnico, ya que centraliza el control de activos tecnológicos utilizados por diferentes agencias, departamentos y usuarios. De manera indirecta, también beneficia a usuarios administrativos y operativos de la empresa, al agilizar la atención de incidentes relacionados con reemplazo de equipos, préstamos temporales y disponibilidad de activos.

La solución se ejecuta localmente en Excel, utilizando hojas estructuradas como `Inventario`, `Prestamos`, `Fuera de uso`, `Config` y `Bitacora`. Asimismo, incorpora componentes web de apoyo mediante apertura de enlaces externos para verificación de garantía de fabricantes y consulta de un dashboard en Power BI. En el código suministrado no se observa un backend web propio ni una base de datos externa; la persistencia se realiza en tablas de Excel y la integración web se consume mediante hipervínculos.

## Metodología

El desarrollo siguió un enfoque incremental orientado a necesidades reales del área de soporte. Primero se analizaron las tareas repetitivas y los puntos de error en la actualización manual de inventarios. Posteriormente, se diseñaron formularios para capturar datos, se programaron rutinas modulares para búsquedas y guardado, se agregaron validaciones y se incorporaron reportes exportables. Finalmente, se realizaron pruebas funcionales sobre los flujos de inventario, préstamos y reportes.

La metodología puede describirse en cinco etapas: análisis de requisitos a partir de tickets y observación directa; prototipado de formularios en el editor de VBA; programación modular separando formularios, reportes y bitácora; pruebas unitarias/manuales sobre funciones críticas como validación de fechas, generación de códigos, filtros y exportación; y despliegue mediante la importación de archivos `.frm` y `.bas` dentro de un libro de Excel habilitado para macros.

La estructura del proyecto VBA utiliza:

- **UserForms** para presentar interfaces gráficas y reducir la manipulación directa de hojas.
- **Módulos estándar** para funciones transversales, como bitácora y control de acceso.
- **ListObject** para operar sobre tablas estructuradas de Excel.
- **ListBox y ComboBox** para búsqueda, filtrado y selección de registros.
- **Arreglos en memoria (`Variant`)** para optimizar búsquedas y filtrados.
- **Manejo de errores** mediante `On Error` y mensajes controlados al usuario.
- **Validaciones determinísticas** para fechas y campos obligatorios.
- **Protección y desprotección de hojas** en operaciones críticas de escritura.

El uso de arreglos en memoria permitió mejorar el rendimiento de las búsquedas, evitando recorrer celda por celda la hoja de cálculo. Esto es relevante en inventarios con crecimiento continuo, donde la respuesta del formulario debe mantenerse fluida para no afectar la atención de tickets.

Respecto a los componentes web, se identificó una arquitectura de consumo cliente-servidor desde Excel: el cliente VBA abre recursos externos mediante `FollowHyperlink`, mientras que los servicios remotos, como Power BI y los portales de garantía de fabricantes, atienden la consulta desde el navegador. No se evidencia conexión directa a una API ni a una base de datos web desde el código analizado; la fuente principal de datos permanece en tablas locales de Excel.

**Fragmento técnico 1. Registro auditable de cambios**

```vb
Public Sub RegistrarBitacora(accion As String, codigo As String, detalle As String)
    Dim ws As Worksheet
    Dim sigFila As Long
    Set ws = ObtenerOCrearHojaBitacora()
    If ws Is Nothing Then Exit Sub
    sigFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    With ws
        .Cells(sigFila, 1).Value = Now()
        .Cells(sigFila, 2).Value = ObtenerUsuarioActual()
        .Cells(sigFila, 3).Value = UCase(accion)
```

Este fragmento muestra cómo se registra una acción crítica con fecha, usuario, tipo de operación y código del registro afectado. Su finalidad es proporcionar trazabilidad ante cambios en el inventario o préstamos.

**Fragmento técnico 2. Validación estricta de fechas**

```vb
Private Function EsFechaValida(ByVal sFecha As String) As Boolean
    EsFechaValida = False
    If Len(Trim(sFecha)) <> 10 Then Exit Function
    If Mid(sFecha, 3, 1) <> "/" Then Exit Function
    If Mid(sFecha, 6, 1) <> "/" Then Exit Function
    Dim sDia As String, sMes As String, sAnio As String
    Dim nDia As Integer, nMes As Integer, nAnio As Integer
    sDia = Left(sFecha, 2)
    sMes = Mid(sFecha, 4, 2)
```

El código valida el formato `dd/mm/yyyy` antes de guardar fechas de garantía o préstamo. Esta decisión reduce errores causados por configuraciones regionales distintas en los equipos de los usuarios.

**Fragmento técnico 3. Filtros combinados para reportes**

```vb
If filtroAgencia <> "" Then
    If UCase(Trim(CStr(datos(i, 2)))) <> UCase(filtroAgencia) Then
        incluir = False
    End If
End If
If incluir And filtroEstado <> "" Then
    If UCase(Trim(CStr(datos(i, 9)))) <> UCase(filtroEstado) Then
        incluir = False
    End If
End If
```

Este fragmento evidencia la aplicación de filtros acumulativos. Si un registro no cumple el criterio seleccionado, se excluye del reporte, permitiendo generar informes específicos por agencia o estado.

**Fragmento técnico 4. Integración con dashboard web**

```vb
Private Sub btnDashboard_Click()
    MsgBox "Se abrirá el Dashboard en tu navegador." & vbCrLf & _
           "Recuerda presionar 'Actualizar' en Power BI si guardaste datos nuevos.", _
           vbInformation, "Abriendo Dashboard"
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:= _
        "https://app.powerbi.com/groups/me/reports/61acaaae-1d6c-40a8-9a6b-d70f1147a215/9fa9b6b9312b307ea30b?experience=power-bi"
```

El procedimiento integra la herramienta de Excel con una solución web de visualización en Power BI, facilitando que el usuario consulte métricas desde el navegador.

## Procedimiento

### Fase 1: Inicialización del formulario principal

El usuario abre el formulario principal `frmGeneralCorregido`. Durante la inicialización, el formulario se ajusta al tamaño de la ventana de Excel, protege la interacción con hojas, verifica el acceso, carga listas desde la hoja `Config`, prepara los listados de inventario y préstamos, y genera un nuevo código consecutivo para posibles registros.

El flujo inicia con la carga de catálogos como agencias, departamentos, marcas, estados y tipos de equipo. Esto evita que el usuario escriba valores libres y contribuye a mantener consistencia en los datos.

### Fase 2: Registro y actualización de inventario

El usuario ingresa la información del equipo en el formulario: agencia, departamento, usuario asignado, tipo de equipo, marca, modelo, serie, periféricos, fechas de garantía y comentarios. Antes de guardar, la macro valida campos obligatorios y verifica reglas específicas; por ejemplo, para equipos de escritorio exige mouse y teclado.

Luego se determina si el código ya existe en `TablaInventario`. Si existe, se solicita confirmación para actualizar; si no existe, se agrega una nueva fila. La macro escribe los valores en las columnas correspondientes, registra la acción en bitácora como `NUEVO` o `EDICIÓN`, actualiza el timestamp de actividad y guarda el libro de forma segura.

### Fase 3: Búsqueda y selección de equipos para préstamo

El usuario utiliza el buscador de inventario para localizar un usuario o serie. La búsqueda se ejecuta en memoria mediante arreglos, lo que reduce el tiempo de respuesta en comparación con la lectura directa de celdas. Una vez seleccionado un registro, el formulario muestra usuario, departamento, agencia y serie del equipo afectado.

Después, el usuario presiona el botón para revisar equipos disponibles. El formulario `FormDisponibles` carga `TablaInventario` en memoria y filtra únicamente equipos con estado `DISPONIBLE` o asociados al texto `EQUIPO DISPONIBLE`. El usuario puede buscar por nombre, marca, serie o tipo de equipo.

Antes de asignar un equipo disponible, el formulario valida que ya exista un usuario seleccionado y confirma en tiempo real que la serie aún siga disponible en la tabla. Esta validación evita asignaciones duplicadas cuando varios usuarios consultan el archivo o cuando la información cambia durante la sesión.

### Fase 4: Registro del préstamo y actualización de estados

Al guardar un préstamo, la macro valida que exista usuario afectado, equipo PRE seleccionado y fechas con formato correcto. Después calcula el estado del préstamo:

- `En préstamo`, cuando no existe fecha de retorno.
- `Retornado`, cuando se registra fecha de retorno.
- `Fuera de uso`, cuando el usuario marca el equipo como no utilizable.

El registro se guarda en `TablaPrestamos` y se actualizan estados en `TablaInventario`. Si el equipo dañado todavía cuenta con garantía, se marca como `EN TALLER / GARANTÍA`; si no es retornable, se traslada a la hoja `Fuera de uso` con un código consecutivo `FDU`. De esta forma, la automatización relaciona inventario, préstamos y bajas dentro del mismo flujo operativo.

### Fase 5: Generación de reportes

El usuario abre `FormReporte`, selecciona filtros por agencia, estado, tipo de equipo y condición de garantía, y presiona `Vista Previa`. La macro lee la tabla de inventario en un arreglo y evalúa cada filtro. Los registros que cumplen los criterios se muestran en un `ListBox`.

Si existen resultados, el botón de exportación se habilita. El usuario puede generar un archivo PDF o Excel. Para PDF, la macro crea una hoja temporal formateada con título, subtítulo de filtros, fecha, usuario, encabezados, datos y configuración de página; luego exporta y elimina la hoja temporal. Para Excel, se permite exportar inventario filtrado, préstamos, equipos fuera de uso o las tres hojas.

### Fase 6: Bitácora, control de acceso e integraciones web

Cada operación crítica, como nuevo registro, edición, eliminación o préstamo, se registra en la hoja `Bitacora`. El módulo también contempla un control informativo de acceso multiusuario mediante celdas de la hoja `Config`, con usuario y fecha de apertura. Esto permite advertir cuando otra persona utiliza el archivo y reduce el riesgo de sobrescrituras en ambientes compartidos como OneDrive o carpetas de red.

Como apoyo a procesos de soporte, el formulario abre portales web de garantía de HP y Dell copiando la serie al portapapeles, y también abre un tablero de Power BI. Estas integraciones no sustituyen el registro local, pero permiten conectar el flujo administrativo de Excel con consultas externas necesarias para toma de decisiones.

## Conclusión

La automatización desarrollada permitió transformar una gestión manual y dispersa del inventario tecnológico en un flujo guiado por formularios, validaciones y reportes. En el contexto de Alimentos Maravilla de Honduras S.A., esto contribuyó a mejorar la consistencia de los datos, reducir errores de digitación, acelerar la atención de tickets de soporte y liberar carga operativa del personal de IT.

El impacto principal se observa en la reducción de pasos manuales para registrar equipos, localizar activos, asignar préstamos, controlar garantías y generar informes. Actividades que anteriormente podían requerir revisar varias hojas, copiar datos y construir reportes manualmente quedaron integradas en botones y formularios. Además, la bitácora fortaleció la trazabilidad de cambios, aspecto importante cuando varias personas consultan o actualizan información sensible del inventario.

Desde la perspectiva de Ingeniería en Ciencias de la Computación, el practicante aplicó competencias de programación estructurada, análisis de sistemas, diseño de interfaces, validación de datos, manejo de eventos, optimización con estructuras en memoria, pruebas funcionales y documentación técnica. También se evidenció criterio para integrar herramientas locales con recursos web, como dashboards y portales de fabricantes, ampliando la utilidad del sistema para el entorno empresarial.

## Recursos visuales sugeridos

### Diagrama de flujo

Se recomienda elaborar un diagrama de flujo del proceso "Registro de préstamo y actualización de inventario" con los siguientes elementos:

1. **Inicio**: El usuario abre el formulario principal de inventario.
2. **Proceso**: El sistema carga listas desde `Config`, inventario y préstamos.
3. **Proceso**: El usuario busca un usuario o serie afectada.
4. **Decisión**: ¿Se seleccionó un usuario válido?
   - No: Mostrar advertencia y regresar a búsqueda.
   - Sí: Continuar.
5. **Proceso**: El usuario abre el formulario de equipos disponibles.
6. **Proceso**: El sistema filtra equipos con estado disponible.
7. **Decisión**: ¿El equipo seleccionado sigue disponible?
   - No: Actualizar lista y solicitar nueva selección.
   - Sí: Cargar datos del equipo PRE en el formulario principal.
8. **Proceso**: El usuario ingresa fechas, tarea y justificación.
9. **Decisión**: ¿Las fechas tienen formato válido?
   - No: Mostrar error y solicitar corrección.
   - Sí: Continuar.
10. **Proceso**: Guardar préstamo en `TablaPrestamos`.
11. **Decisión**: ¿El equipo dañado queda fuera de uso?
    - Sí: Evaluar garantía y enviar a taller o a `Fuera de uso`.
    - No: Actualizar estado del equipo PRE como prestado o disponible.
12. **Proceso**: Registrar acción en bitácora y guardar libro.
13. **Fin**: Mostrar mensaje de éxito y refrescar listados.

### Fragmentos de código como Figuras

**Figura 6. Función de registro de bitácora para auditoría de cambios.**  
Se recomienda utilizar el fragmento `RegistrarBitacora` del archivo `ModuloBitacora.bas`, debido a que evidencia trazabilidad, usuario del sistema y fecha de ejecución.

**Figura 7. Validación estricta de fecha en formato dd/mm/yyyy.**  
Se recomienda utilizar la función `EsFechaValida` del archivo `frmGeneralCorregido.frm`, ya que muestra control de calidad de datos antes de registrar garantías o préstamos.

**Figura 8. Aplicación de filtros combinados para generación de reportes.**  
Se recomienda utilizar el bloque de filtros de `EjecutarFiltroReporte` en `FormReporte.frm`, porque representa la lógica de consulta del inventario por criterios administrativos.

**Figura 9. Integración con dashboard web de Power BI.**  
Se recomienda utilizar el procedimiento `btnDashboard_Click` de `frmGeneralCorregido.frm`, debido a que demuestra la conexión entre la herramienta local en Excel y un recurso web corporativo.

### Sugerencia de captura de pantalla

Se recomienda tomar una captura de pantalla en el momento en que el usuario abre el formulario `FormReporte`, selecciona filtros y visualiza resultados en la vista previa antes de exportar. Esta captura permite evidenciar la interfaz de usuario, los criterios de búsqueda, el conteo de registros y la preparación del reporte final.

Como segunda evidencia visual, se sugiere capturar el formulario `FormDisponibles` después de filtrar equipos disponibles y antes de seleccionar uno para préstamo. Esta imagen demostraría el proceso de asignación controlada de equipos y la validación operativa previa al registro del préstamo.
