Attribute VB_Name = "NewMacros"
Sub AutoConvertUrlaTexto()
    Dim objDoc As Document
    Dim hl As Hyperlink
    Dim rng As Range
    Dim startPos As Long
    Dim endPos As Long
    Dim textRange As Range
    
    ' Asigna el documento actual a objDoc
    Set objDoc = ActiveDocument
    
    ' Habilita la opci—n de AutoFormato para reemplazar enlaces con hiperv’nculos
    Word.Options.AutoFormatReplaceHyperlinks = True
    
    ' Aplica AutoFormato al rango completo del documento
    objDoc.Range.AutoFormat
    
    ' Establece el rango en todo el documento para buscar las URLs convertidas
    Set rng = objDoc.Range
    
    ' Buscamos unicamente en los hiperv’nculos del documento
    For Each hl In objDoc.Hyperlinks
        ' Verificamos si el hiperv’nculo fue a–adido recientemente usando la direccion base
        Set textRange = hl.Range
        startPos = InStr(textRange.Text, "https://estudusfqedu-my.sharepoint.com/")
        
        If startPos > 0 Then
            ' Cambia el texto a la palabra que gustes
            hl.TextToDisplay = "AutoFormat"
        End If
    Next hl
End Sub

Sub ConvertirURLsEnHiperVvistaprevia()
    Dim fld As Field
    Dim rng As Range
    Dim linkText As String
    
    ' Recorrer todos los campos en el documento
    For Each fld In ActiveDocument.Fields
        ' Verificar si el campo es un campo de combinaci—n
        If fld.Type = wdFieldMergeField Then
            Set rng = fld.Result
            linkText = rng.Text
            ' Verifica si el texto del campo contiene la URL espec’fica
            If InStr(1, linkText, "https://estudusfqedu-my.sharepoint.com/", vbTextCompare) > 0 Then
                ' Eliminar el texto anterior y agregar el hiperv’nculo
                rng.Text = ""
                rng.Hyperlinks.Add Anchor:=rng, Address:=linkText, TextToDisplay:="click"
            End If
        End If
    Next fld
End Sub

Sub SendMailT()
    Dim doc As Document
    Dim pagina As Integer
    Dim correo As String
    Dim emailApp As Object
    Dim email As Object
    Dim tempDoc As Document
    
    Set doc = ActiveDocument
    Set emailApp = CreateObject("Outlook.Application")
    
    For pagina = 1 To doc.Sections.Count
        ' Extraer el correo del pie de p‡gina
        correo = doc.Sections(pagina).Footers(wdHeaderFooterPrimary).Range.Text
        
        ' Crear un nuevo documento temporal para copiar el contenido de la p‡gina
        Set tempDoc = Documents.Add
        doc.Sections(pagina).Range.Copy
        tempDoc.Range.Paste
        
        ' Crear el email
        Set email = emailApp.CreateItem(0)
        With email
            .To = correo
            .Subject = "ONE mensaje"
            .Body = tempDoc.Content.Text
            .Display
        End With
        
        ' Cerrar y eliminar el documento temporal sin guardar cambios
        tempDoc.Close False
        Set tempDoc = Nothing
    Next pagina
    
    MsgBox "Correos enviados exitosamente.", vbInformation
End Sub

Sub SendMailH()
    Dim doc As Document
    Dim pagina As Integer
    Dim correo As String
    Dim emailApp As Object
    Dim email As Object
    Dim tempDoc As Document
    Dim tempFilePath As String
    Dim tempFileNum As Integer
    Dim tempFileContent As String
    
    Set doc = ActiveDocument
    Set emailApp = CreateObject("Outlook.Application")
    
    For pagina = 1 To doc.Sections.Count
        ' Extraer el correo del pie de p‡gina
        correo = doc.Sections(pagina).Footers(wdHeaderFooterPrimary).Range.Text
        
        ' Crear un nuevo documento temporal para copiar el contenido de la p‡gina
        Set tempDoc = Documents.Add
        doc.Sections(pagina).Range.Copy
        tempDoc.Range.Paste
        
        ' Guardar el documento temporal como HTML
        tempFilePath = Environ("TEMP") & "\tempfile.html"
        tempDoc.SaveAs2 FileName:=tempFilePath, FileFormat:=wdFormatHTML
        
        ' Leer el contenido HTML del archivo guardado
        tempFileNum = FreeFile
        Open tempFilePath For Input As tempFileNum
        tempFileContent = Input$(LOF(tempFileNum), tempFileNum)
        Close tempFileNum
        
        ' Crear el email con cuerpo en HTML
        Set email = emailApp.CreateItem(0)
        With email
            .To = correo
            .Subject = "ONE mensaje"
            .HTMLBody = tempFileContent & .HTMLBody
            ' Abre una vista previa del correo
            .Display
        End With
        
        ' Cerrar y eliminar el documento temporal sin guardar cambios
        tempDoc.Close False
        Set tempDoc = Nothing
        
        ' Borrar el archivo temporal
        Kill tempFilePath
    Next pagina
    
    MsgBox "Correos enviados exitosamente.", vbInformation
End Sub

