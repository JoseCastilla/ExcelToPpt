Attribute VB_Name = "Módulo1"
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

