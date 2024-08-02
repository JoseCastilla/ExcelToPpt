
# Automatización de Presentaciones de PowerPoint con Datos de Excel

## Descripción

Este proyecto demuestra cómo automatizar la generación y actualización de presentaciones de PowerPoint utilizando datos de una hoja de Excel a través de VBA (Visual Basic for Applications). El proyecto incluye los siguientes pasos:

1. **Preparación del Archivo de Excel:**
   - Se crea un archivo de Excel que contiene una tabla con los datos necesarios para generar y actualizar las diapositivas de PowerPoint. La tabla incluye columnas para el índice de la diapositiva, el identificador del título, el identificador de la forma en la diapositiva, el texto del título y el contenido de la diapositiva.

2. **Ejecución del Código VBA:**
   - Un script VBA se encarga de abrir o crear una presentación de PowerPoint en una carpeta específica. Luego, itera a través de los datos de la tabla en Excel para actualizar el contenido de las diapositivas en PowerPoint.
   - El script verifica si las diapositivas y las formas existen, las crea si es necesario y actualiza su contenido con los datos proporcionados.

3. **Resultados:**
   - Al ejecutar el script VBA, se genera una presentación de PowerPoint con títulos y contenido basado en los datos del archivo de Excel. La presentación se guarda en una carpeta llamada "presentaciones" en el mismo directorio que el archivo de Excel.

## Estructura del Archivo de Excel

La estructura de la tabla en el archivo de Excel es la siguiente:

| SlideIndex | SlideIdentifier | Placeholder             | ReplacementText                       | ContentText                                      |
|------------|------------------|-------------------------|---------------------------------------|--------------------------------------------------|
| 1          | Title 1          | Placeholder 1           | Ingeniería Textil                     | La ingeniería textil es una disciplina ...       |
| 2          | Title 2          | Placeholder 2           | Historia de la Ingeniería Textil      | La historia de la ingeniería textil ...          |
| 3          | Title 3          | Placeholder 3           | Definición de la Ingeniería Textil    | La definición de la ingeniería textil ...        |

## Código VBA

El código VBA se encuentra en un módulo dentro del archivo de Excel y realiza las siguientes funciones:

- Abre o crea una presentación de PowerPoint.
- Itera a través de las filas de la tabla en Excel.
- Actualiza el título y el contenido de las diapositivas en PowerPoint.
- Guarda la presentación en la carpeta "presentaciones".

```vba
Sub Main()
    Dim mainSht As Worksheet
    Dim mainTable As ListObject
    Dim pptFilePath As String
    Dim Row As Range
    
    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    
    Dim slideIndex As Long
    Dim strSlideIdentifier As String
    Dim strReplacement As String
    Dim strContent As String
    
    Dim excelFilePath As String
    Dim presentationDir As String
    
    On Error GoTo ErrorHandler
    
    ' Obtener el directorio del archivo de Excel
    excelFilePath = ThisWorkbook.Path
    presentationDir = excelFilePath & "\presentaciones"
    
    ' Crear la carpeta "presentaciones" si no existe
    If Dir(presentationDir, vbDirectory) = "" Then
        MkDir presentationDir
    End If
    
    Set mainSht = ThisWorkbook.Sheets("Principal")
    
    ' Verificar si la tabla "T_Principal" existe
    On Error Resume Next
    Set mainTable = mainSht.ListObjects("T_Principal")
    On Error GoTo ErrorHandler
    
    If mainTable Is Nothing Then
        MsgBox "No se pudo encontrar la tabla 'T_Principal' en la hoja 'Principal'.", vbCritical
        Exit Sub
    End If
    
    pptFilePath = presentationDir & "\presentacion.pptx"
    
    ' Crear una nueva aplicación de PowerPoint
    Set pptApp = CreateObject("PowerPoint.Application")
    
    ' Verificar si el archivo de PowerPoint existe
    If Dir(pptFilePath) = "" Then
        ' El archivo no existe, crear una nueva presentación
        Set pptPres = pptApp.Presentations.Add
        pptPres.SaveAs pptFilePath
    Else
        ' El archivo existe, abrir la presentación
        Set pptPres = pptApp.Presentations.Open(Filename:=pptFilePath)
    End If
    
    For Each Row In mainTable.DataBodyRange.Rows
        slideIndex = Row.Cells(1, 1)
        strSlideIdentifier = Row.Cells(1, 2).Text
        strReplacement = Row.Cells(1, 4).Text
        strContent = Row.Cells(1, 5).Text
        
        ' Verificar si la diapositiva existe, si no, agregarla
        If slideIndex > pptPres.Slides.Count Then
            Set pptSlide = pptPres.Slides.Add(slideIndex, 1) ' 1 corresponde a ppLayoutText
        Else
            Set pptSlide = pptPres.Slides(slideIndex)
        End If
        
        ' Actualizar el título de la diapositiva
        Call TextToPPT(pptSlide, strSlideIdentifier, strReplacement)
        
        ' Actualizar el contenido de la diapositiva
        Call TextToPPT(pptSlide, strSlideIdentifier & "_content", strContent)
    Next
    
    pptPres.Save
    pptPres.Close
    pptApp.Quit
    Set pptPres = Nothing
    Set pptApp = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical
    If Not pptPres Is Nothing Then
        pptPres.Close
        Set pptPres = Nothing
    End If
    If Not pptApp Is Nothing Then
        pptApp.Quit
        Set pptApp = Nothing
    End If
End Sub

Sub TextToPPT(slide As Object, strIdentifier As String, strReplacement As String)
    Dim shape As Object
    On Error Resume Next ' En caso de que la forma no exista
    Set shape = slide.Shapes(strIdentifier)
    If shape Is Nothing Then
        ' Si la forma no existe, crear una nueva caja de texto con el identificador
        Set shape = slide.Shapes.AddTextbox(1, 10, 10, 500, 50) ' 1 corresponde a msoTextOrientationHorizontal
        shape.Name = strIdentifier
    End If
    shape.TextFrame.TextRange.Text = strReplacement
    On Error GoTo 0
End Sub
```

## Instrucciones para la Ejecución

1. **Preparación del Entorno:**
   - Asegúrate de tener Microsoft Excel y PowerPoint instalados.
   - Descarga y abre el archivo de Excel `ingenieria_textil_con_contenido.xlsx`.

2. **Agregar el Código VBA:**
   - Abre el Editor de VBA presionando `Alt + F11` en Excel.
   - Inserta un nuevo módulo y copia el código VBA proporcionado.

3. **Ejecutar el Código VBA:**
   - Coloca el cursor dentro de la subrutina `Main`.
   - Presiona `F5` para ejecutar el script.

4. **Revisar la Presentación Generada:**
   - La presentación de PowerPoint se guardará en una carpeta llamada "presentaciones" ubicada en el mismo directorio que el archivo de Excel.
   - Abre la presentación para revisar los títulos y contenidos actualizados.

## Licencia

Este proyecto está licenciado bajo la Licencia MIT. Consulta el archivo [LICENSE](LICENSE) para obtener más detalles.
