# Sistema de inventario de reactivos mediante VBA en Excel

## Antecedentes

Durante la práctica profesional supervisada en Alimentos Maravilla de Honduras S.A., se identificó que el departamento de control de calidad requería un mecanismo más ordenado para administrar reactivos, utensilios y materiales utilizados en análisis internos. En una empresa industrial y alimentaria, el control de reactivos es un proceso sensible porque se relaciona con inspecciones, pruebas de laboratorio, seguimiento de vencimientos y disponibilidad de insumos para verificar la calidad de materias primas, producto en proceso y producto terminado.

Antes de la automatización, el registro manual de reactivos podía generar inconsistencias en nombres, unidades de medida, cantidades disponibles, fechas de vencimiento y estados de inventario. Estas inconsistencias podían ocasionar compras no planificadas, uso de reactivos vencidos, dificultad para justificar consumos y pérdida de tiempo al elaborar reportes semanales. Además, el personal debía revisar manualmente hojas de Excel para conocer si un reactivo estaba disponible, próximo a vencer, vencido, con stock bajo o sin existencia.

La oportunidad de mejora se identificó mediante la observación directa de tareas repetitivas y solicitudes del área usuaria, especialmente aquellas relacionadas con actualización de stock, búsqueda de reactivos, generación de cierres y elaboración de reportes. A partir de estas necesidades, el practicante desarrolló un segundo sistema en Excel con formularios VBA, orientado a facilitar operaciones CRUD, cálculo de estados, conversión de unidades y generación de reportes para control de calidad.

## Objetivo

El objetivo técnico principal fue automatizar la gestión del inventario de reactivos mediante formularios VBA en Excel, permitiendo registrar, actualizar, eliminar, buscar, rebajar y agregar existencias de forma controlada, con validaciones de datos y cálculo automático del estado de cada reactivo.

De forma específica, la solución buscó:

- Registrar reactivos con identificadores consecutivos bajo el prefijo `RCT`.
- Registrar artículos o utensilios simples con identificadores bajo el prefijo `UTN`.
- Controlar cantidades por envase, capacidad y total disponible en unidad base.
- Calcular estados como `DISPONIBLE`, `PROXIMO A VENCER`, `VENCIDO`, `STOCK BAJO` o `SIN EXISTENCIA`.
- Rebajar inventario por consumo y agregar stock por ingreso de nuevos reactivos.
- Convertir unidades compatibles, como gramos a kilogramos o mililitros a litros.
- Generar reportes en PDF o Excel desde los datos visibles en el formulario.
- Crear cierres semanales mediante copias de la hoja principal.

## Alcance

El alcance funcional del sistema se concentra en dos formularios VBA:

- `INVENTARIOCRUD.frm`: formulario principal para administrar reactivos, ejecutar operaciones de creación, lectura, actualización y eliminación, controlar entradas y salidas de stock, buscar registros, calcular estados, generar reportes y crear cierres semanales.
- `AGGARTICULO.frm`: formulario complementario para agregar artículos simples o utensilios, con generación automática de código `UTN` y registro de observaciones.

La solución utiliza `Hoja1` como base de datos local en Excel. En esta hoja se crean y administran doce columnas: `ID`, `REACTIVO`, `CANTIDAD`, `FECHA VENCIMIENTO`, `FECHA PEDIDO`, `FECHA APERTURA`, `UNIDAD`, `ESTADO`, `OBSERVACIONES`, `CAPACIDAD`, `TOTAL BASE` y `ENVASE`. Estas columnas permiten almacenar información operativa suficiente para el seguimiento básico del inventario del área.

Los usuarios beneficiados son principalmente colaboradores del departamento de control de calidad, responsables de registrar reactivos, consultar disponibilidad, controlar vencimientos y preparar reportes de inventario. De manera indirecta, también se beneficia el área administrativa o de compras, ya que la información generada puede apoyar decisiones sobre reposición de insumos y prevención de desabastecimiento.

La solución se ejecuta localmente en Excel y no evidencia componentes web, conexión a bases de datos externas ni consumo de APIs. Esta característica la diferencia del primer sistema documentado, que incluía enlaces hacia Power BI y portales de garantía. En este segundo sistema, la funcionalidad se concentra en formularios, macros, cálculos locales y exportación de archivos.

## Metodología

El desarrollo siguió un enfoque incremental basado en necesidades operativas del departamento de control de calidad. Primero se definió la estructura mínima de datos requerida para controlar reactivos y utensilios. Luego se implementó un formulario principal para operaciones CRUD y un formulario secundario para artículos simples. Posteriormente, se agregaron validaciones, funciones auxiliares de conversión de unidades, cálculo automático de estados, búsqueda optimizada y exportación de reportes.

La metodología puede describirse en cinco etapas: análisis de requisitos mediante observación de tareas repetitivas; prototipado de formularios en el editor de VBA; programación modular dentro de los eventos de cada `UserForm`; pruebas funcionales sobre registro, actualización, rebaja, entrada de stock, búsqueda y reportes; y despliegue mediante importación de los archivos `.frm` en un libro de Excel habilitado para macros.

La estructura técnica del proyecto utiliza:

- **UserForms** para encapsular la interacción del usuario y evitar edición directa de la hoja.
- **Eventos de botones y cajas de texto** para ejecutar operaciones específicas.
- **`Application.Match`** para ubicar registros por ID dentro de la hoja.
- **Arreglos dinámicos (`Variant`)** para cargar y filtrar datos en el `ListBox`.
- **Funciones auxiliares** para generación de códigos, conversión de unidades y cálculo de estado.
- **Validaciones de campos obligatorios** para reactivo, cantidad, unidad y fecha de vencimiento.
- **Manejo básico de errores** mediante verificación de valores, `IsError`, `IsDate`, `IsNumeric` y mensajes al usuario.

Entre las técnicas de optimización se observa el uso de matrices para asignar resultados directamente al `ListBox`, evitando actualizaciones celda por celda durante la búsqueda. También se automatiza la creación de encabezados cuando la hoja está vacía, lo que reduce tareas iniciales de configuración y estandariza la estructura de datos.

**Fragmento técnico 1. Creación automática de encabezados**

```vb
Sub VerificarEncabezados()
    Dim ws As Worksheet
    Set ws = Hoja1
    If ws.Range("A1").Value = "" Then
        Dim encabezados As Variant
        encabezados = Array("ID", "REACTIVO", "CANTIDAD", "FECHA VENCIMIENTO", _
            "FECHA PEDIDO", "FECHA APERTURA", "UNIDAD", "ESTADO", "OBSERVACIONES", _
            "CAPACIDAD", "TOTAL BASE", "ENVASE")
```

Este fragmento muestra cómo el sistema prepara la estructura de la hoja cuando aún no existen encabezados. La automatización evita errores de configuración inicial y asegura que los datos se registren en columnas estandarizadas.

**Fragmento técnico 2. Registro de reactivos con validaciones**

```vb
If txtReactivo.Value = "" Then MsgBox "Ingrese el reactivo", vbExclamation: Exit Sub
If Not IsNumeric(txtCantidad.Value) Or Val(txtCantidad.Value) <= 0 Then MsgBox "Cantidad inválida", vbExclamation: Exit Sub
If Not IsDate(txtFechaVencimiento.Value) Then MsgBox "Fecha vencimiento inválida", vbExclamation: Exit Sub
If cmbUnidad.Value = "" Then MsgBox "Seleccione unidad", vbExclamation: Exit Sub
capacidad = Val(InputBox("Ingrese la capacidad por envase (Ej: 500 para 500ml)", "Capacidad del Envase"))
If capacidad <= 0 Then MsgBox "Capacidad inválida", vbExclamation: Exit Sub
cantidadEnvases = CDbl(txtCantidad.Value)
totalBase = cantidadEnvases * capacidad
estado = CalcularEstado(CDate(txtFechaVencimiento.Value))
```

El bloque valida los datos mínimos antes de guardar un reactivo y calcula el total disponible a partir de la cantidad de envases y la capacidad por envase. Esto reduce errores de digitación y mejora la confiabilidad del inventario.

**Fragmento técnico 3. Rebaja de stock con conversión de unidades**

```vb
cantidadNormalizada = ConvertirUnidades(cantidadIngresada, unidadIngresada, unidadProducto)
If cantidadNormalizada = -1 Then
    MsgBox "No se puede convertir de " & unidadIngresada & " a " & unidadProducto, vbCritical
    Exit Sub
End If
If cantidadNormalizada > totalActual Then
    MsgBox "Stock insuficiente." & vbCrLf & _
           "Tienes: " & totalActual & " " & unidadProducto, vbCritical
    Exit Sub
End If
```

Este fragmento evidencia el control aplicado antes de rebajar inventario. La macro convierte la unidad ingresada a la unidad del producto y evita consumos superiores al stock disponible.

**Fragmento técnico 4. Búsqueda optimizada mediante arreglo**

```vb
datosOriginales = ws.Range("A2:L" & ultFila).Value
count = 0
For i = 1 To UBound(datosOriginales, 1)
    If InStr(LCase(datosOriginales(i, 1)), texto) > 0 Or _
       InStr(LCase(datosOriginales(i, 2)), texto) > 0 Then
        count = count + 1
    End If
Next i
ReDim datosFiltrados(0 To count - 1, 0 To 11)
```

El uso de arreglos permite filtrar registros por ID o nombre de reactivo sin consultar repetidamente la hoja de cálculo. Esta técnica mejora la experiencia del usuario cuando el inventario crece.

## Procedimiento

### Fase 1: Inicialización del formulario principal

El usuario abre el formulario `INVENTARIOCRUD`. Durante la inicialización, la macro centra la ventana, verifica si `Hoja1` tiene encabezados, carga los combos de unidades y tipos de envase, configura el `ListBox` con doce columnas, bloquea campos automáticos como `txtID` y `txtEstado`, y carga los datos existentes en la lista.

Si la hoja no tiene estructura previa, el sistema crea los encabezados y aplica formato visual. Esta fase permite que el archivo quede preparado para operar sin requerir configuración manual avanzada por parte del usuario.

### Fase 2: Registro de un reactivo

El usuario ingresa el nombre del reactivo, cantidad de envases, fecha de vencimiento, fecha de pedido, fecha de apertura, unidad, envase y observaciones. Al presionar el botón de guardar, el sistema valida campos obligatorios y solicita la capacidad por envase mediante un cuadro de entrada.

Posteriormente, el sistema calcula el total base multiplicando la cantidad de envases por la capacidad, determina el estado según la fecha de vencimiento y genera un ID consecutivo con el prefijo `RCT`. Finalmente, escribe los datos en la siguiente fila disponible de `Hoja1`, actualiza el `ListBox` y limpia los campos para un nuevo registro.

### Fase 3: Actualización y eliminación de registros

Cuando el usuario selecciona un elemento del `ListBox`, el sistema carga los datos del reactivo en los controles del formulario. Esto permite modificar cantidades, fechas, unidad, envase u observaciones. Al actualizar, la macro localiza la fila mediante `Application.Match`, recalcula el total base y actualiza el estado de acuerdo con la fecha de vencimiento.

Para eliminar un registro, el usuario debe seleccionar un reactivo y confirmar la acción. La macro busca el ID en la primera columna de `Hoja1`, elimina la fila correspondiente, recarga los datos y limpia los controles.

### Fase 4: Rebaja de inventario por consumo

El usuario selecciona un reactivo y presiona el botón de rebaja. El sistema solicita la cantidad consumida y la unidad correspondiente. Luego identifica la unidad registrada del producto, convierte la cantidad ingresada a una unidad compatible y valida que no se consuma más de lo disponible.

Si la operación es válida, se descuenta la cantidad del total base, se recalcula la cantidad de envases y se actualiza el estado. Si el total llega a cero, el estado cambia a `SIN EXISTENCIA`; si queda menos de un envase, cambia a `STOCK BAJO` y se muestra una alerta al usuario.

### Fase 5: Agregar stock al inventario

Cuando llegan nuevos reactivos, el usuario selecciona el registro existente y utiliza el botón de agregar stock. La macro solicita la cantidad recibida y su unidad, convierte el valor a la unidad registrada, suma el resultado al total base, recalcula la cantidad de envases y actualiza la fecha de pedido con la fecha actual.

Este flujo permite registrar entradas de inventario sin crear duplicados del mismo reactivo, manteniendo continuidad en el historial operativo del registro.

### Fase 6: Búsqueda, reportes y cierre semanal

El campo de búsqueda filtra el inventario por ID o nombre de reactivo. La búsqueda se ejecuta con arreglos en memoria y actualiza el `ListBox` con los resultados coincidentes.

Para reportes, el usuario puede exportar los datos visibles a PDF o Excel. La macro crea un libro temporal, copia los encabezados y registros del `ListBox`, aplica formato básico y solicita la ruta de guardado. En el caso de PDF, configura la página en orientación horizontal y ajusta el ancho a una página.

El cierre semanal permite copiar `Hoja1` a una nueva hoja con el nombre indicado por el usuario. Antes de crearla, se limpian caracteres no permitidos y se verifica que no exista una hoja con el mismo nombre. Esta función permite conservar cortes históricos del inventario de reactivos.

### Fase 7: Registro de artículos o utensilios simples

El formulario `AGGARTICULO` se utiliza para registrar artículos simples, como utensilios o materiales auxiliares que no requieren el mismo nivel de control de vencimiento que un reactivo. Al abrirse, genera automáticamente un ID con prefijo `UTN`, bloquea el campo de ID y permite ingresar nombre y observaciones.

Al guardar, valida que exista nombre, escribe el ID, nombre y observaciones en `Hoja1`, y genera un nuevo ID para continuar registrando artículos sin cerrar el formulario.

## Conclusión

La automatización del inventario de reactivos permitió trasladar un proceso manual y propenso a inconsistencias hacia una herramienta guiada por formularios, validaciones y cálculos automáticos. En el departamento de control de calidad, esto favorece el seguimiento de vencimientos, disponibilidad de reactivos, control de consumos, entradas de stock y generación de reportes para la toma de decisiones.

El impacto principal se refleja en la reducción de tiempo dedicado a revisar hojas manualmente, la disminución de errores por unidades incompatibles, la detección oportuna de reactivos próximos a vencer y la posibilidad de conservar cierres periódicos del inventario. Para una empresa alimentaria, estos controles contribuyen indirectamente a la continuidad de los análisis de calidad y a la disponibilidad de insumos necesarios para verificar condiciones de producción.

Desde la perspectiva de Ingeniería en Ciencias de la Computación, el practicante aplicó competencias de programación orientada a eventos, análisis de procesos, diseño de interfaces, validación de datos, manejo de estructuras de almacenamiento en Excel, optimización mediante arreglos, generación de reportes y pruebas funcionales. Además, se evidencia capacidad para adaptar una misma tecnología, VBA en Excel, a necesidades distintas: primero para inventario tecnológico y luego para inventario de reactivos del área de control de calidad.

## Recursos visuales sugeridos

### Diagrama de flujo

Se recomienda elaborar un diagrama de flujo del proceso "Rebaja de inventario por consumo de reactivo" con los siguientes elementos:

1. **Inicio**: El usuario abre el formulario `INVENTARIOCRUD`.
2. **Proceso**: El sistema carga encabezados, combos y datos de `Hoja1`.
3. **Proceso**: El usuario busca y selecciona un reactivo.
4. **Decisión**: ¿Existe un reactivo seleccionado?
   - No: Mostrar mensaje de advertencia y regresar a la selección.
   - Sí: Continuar.
5. **Proceso**: El usuario ingresa cantidad consumida y unidad.
6. **Decisión**: ¿La cantidad es mayor que cero y la unidad fue ingresada?
   - No: Cancelar operación.
   - Sí: Continuar.
7. **Proceso**: Convertir la cantidad ingresada a la unidad registrada del producto.
8. **Decisión**: ¿La conversión es válida?
   - No: Mostrar error de unidad incompatible.
   - Sí: Continuar.
9. **Decisión**: ¿Existe stock suficiente?
   - No: Mostrar mensaje de stock insuficiente.
   - Sí: Continuar.
10. **Proceso**: Restar la cantidad consumida al total base.
11. **Proceso**: Recalcular cantidad de envases.
12. **Decisión**: ¿El nuevo total es cero o menor?
    - Sí: Actualizar estado a `SIN EXISTENCIA`.
    - No: Evaluar si queda menos de un envase.
13. **Decisión**: ¿Queda menos de un envase?
    - Sí: Actualizar estado a `STOCK BAJO` y mostrar alerta.
    - No: Calcular estado según vencimiento.
14. **Proceso**: Guardar cambios en `Hoja1` y recargar el listado.
15. **Fin**: Mostrar confirmación de rebaja aplicada.

### Fragmentos de código como Figuras

**Figura 10. Creación automática de estructura de inventario de reactivos.**  
Se recomienda utilizar el fragmento `VerificarEncabezados` del archivo `INVENTARIOCRUD.frm`, porque evidencia cómo el sistema estandariza las columnas requeridas para operar.

**Figura 11. Validaciones y cálculo inicial para registro de reactivos.**  
Se recomienda utilizar el bloque de validaciones del procedimiento `btnGuardar_Click`, ya que muestra el control previo sobre cantidad, fecha, unidad y capacidad de envase.

**Figura 12. Rebaja de stock con conversión de unidades y validación de existencia.**  
Se recomienda utilizar el fragmento de `btnRebajar_Click`, debido a que representa el proceso más crítico del sistema: descontar consumo sin permitir unidades incompatibles ni cantidades superiores al stock.

**Figura 13. Filtro de búsqueda mediante arreglo en memoria.**  
Se recomienda utilizar el bloque de `txtBuscar_Change`, ya que evidencia una técnica de optimización para filtrar registros por ID o nombre sin recorrer visualmente toda la hoja.

### Sugerencia de captura de pantalla

Se recomienda tomar una captura de pantalla del formulario `INVENTARIOCRUD` después de seleccionar un reactivo en el `ListBox`, cuando los campos del formulario aparecen llenos y el indicador de estado muestra si el reactivo está disponible, próximo a vencer, vencido o con stock bajo. Esta captura permitiría evidenciar la interfaz principal, la consulta del inventario y la alerta visual para control de calidad.

Como segunda evidencia visual, se sugiere capturar el momento de generación del reporte, específicamente cuando el sistema pregunta si se desea guardar como PDF o Excel. Esta imagen demostraría la capacidad del sistema para convertir la información operativa en un documento útil para seguimiento administrativo.
