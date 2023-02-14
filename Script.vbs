Sub Renombrar_fotos()
    
    Dim Ainicial As String
    Dim Afinal As String
    Dim j As Double
    j = 1
    Ainicial = "DSC0"
    Afinal = "DSC0"
    
    Dim Ruta As String
    Ruta = ActiveDocument.Path
    
    For i = Right(Ainicial, 4) To Right(Afinal, 4)
        
        completo = Ruta & "\FOTOS\DSC0" & i & ".jpg"
        If Dir(completo) = "" Then
            j = j - 1
        Else
            Name completo As Ruta & "\FOTOS\" & i - Right(Ainicial, 4) + j & ".jpg"
        End If
    Next i

End Sub

Sub Generar_documento()
    
    'Definición de variables
    Dim registro() As Variant
    Dim fecha As String
    Dim apto As String
    Dim torre As Double
    Dim Ruta As String
    Dim conjunto As String
    Dim direccion As String
    Ruta = ActiveDocument.Path
    Dim municipio As String
    Dim obra As String
    Dim constructora As String
    
    'Elementos básicos del documento
    registro() = Array( _
)
    
    fecha = ""
    
    apto = ""
    
    'Datos de la edificación
    conjunto = "Edificio "
    direccion = "Calle "
    municipio = "Bogotá"
    
    'Datos del proyecto
    obra = ""
    constructora = ""
    
    'Generando titulo y membrete de tabla
    Selection.HomeKey Unit:=wdStory
    Selection.EndKey Unit:=wdLine
    Selection.TypeText Text:=UCase(obra)
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.EndKey Unit:=wdLine
    Selection.TypeText Text:=UCase(constructora)
    
    ActiveDocument.Tables(1).Cell(1, 2).Range.Text = "Apartamento " & apto & " ubicado en el predio identificado con nomenclatura urbana " & conjunto & " " & direccion & " de la ciudad de " & municipio
    
    ActiveDocument.Tables(1).Cell(2, 3).Range.Text = "FECHA DE TOMA: " & fecha & " del " & Year(Date)
      
    'Llenando tabla
    ActiveDocument.Tables(1).Cell(3, 2).Range.Text = registro(0) & " - " & registro(1)
    ActiveDocument.Tables(1).Cell(3, 3).Range.Text = registro(2)
    Insertar_Imagenes CInt(registro(0)), CInt(registro(1))
    For i = 1 To UBound(registro) / 3
        ThisDocument.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:=3, _
                          DefaultTableBehavior:=wdWord9TableBehavior, _
                          AutoFitBehavior:=wdAutoFitFixed
                          
        With ThisDocument.Tables(i + 1)

            .Rows.Height = 0.56
            .Rows(1).Cells(1).Width = CentimetersToPoints(2.26)
            .Rows(1).Cells(2).Width = CentimetersToPoints(2)
            .Rows(1).Cells(3).Width = CentimetersToPoints(13.25)
            Select Case i
                Case 1 To 6
                    m = 0
                Case 7 To 14
                    m = 1
                Case 14 To 20
                    m = 2
                Case 20 To 27
                    m = 3
                Case Else
                    m = 4
            End Select
            .Rows(1).Cells(1).Range.Text = "00:" & m & vbCrLf & "00:" & m
            .Rows(1).Cells(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Rows(1).Cells(1).VerticalAlignment = wdCellAlignVerticalBottom
            If (registro(i * 3) = registro(i * 3 + 1)) Then
                .Rows(1).Cells(2).Range.Text = registro(i * 3)
            Else
                .Rows(1).Cells(2).Range.Text = registro(i * 3) & " - " & registro(i * 3 + 1)
            End If
            .Rows(1).Cells(2).Range.Bold = wdToggle
            .Rows(1).Cells(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Rows(1).Cells(2).VerticalAlignment = wdCellAlignVerticalBottom
            .Rows(1).Cells(2).Range.Font.Bold = True
            .Rows(1).Cells(3).Range.Text = registro(i * 3 + 2)
            .Rows(1).Cells(3).VerticalAlignment = wdCellAlignVerticalBottom
            
        End With
        If Len(registro(i * 3 + 2)) > 72 Then
            CircleInsert
            ArrowInsert
        End If
        Insertar_Imagenes CInt(registro(i * 3)), CInt(registro(i * 3 + 1))
    Next i
    
End Sub

Sub Insertar_Imagenes(Inicio As Double, Fin As Double)
    Dim Ruta As String
    Dim RutaTotal As String
        
    Ruta = ActiveDocument.Path
    Dim SHP As InlineShape
   
    For i = Inicio To Fin
        Selection.EndKey Unit:=wdStory
        RutaTotal = Ruta & "\FOTOS\" & i & ".jpg"
        
        Set SHP = Selection.InlineShapes.AddPicture(FileName:=RutaTotal, _
            LinkToFile:=False, SaveWithDocument:=True)
        With SHP
            .LockAspectRatio = msoFalse
            .Width = CentimetersToPoints(8.5)
            .Height = CentimetersToPoints(5.18)
        End With
        ActiveDocument.Content.InsertAfter Text:=" "
        
    Next
    Selection.EndKey Unit:=wdStory
    Selection.TypeBackspace
End Sub

Sub CircleInsert()

    Dim y As Double
    y = Selection.Characters.Last.Information(wdVerticalPositionRelativeToPage)

    Dim BBB As Shape

    Set BBB = ActiveDocument.Shapes.AddShape(Type:=msoShapeOval, _
        Left:=70.8, Top:=y + 28, Width:=50, Height:=100, _
        Anchor:=Selection.Range)

    With BBB
        .PictureFormat.TransparentBackground = True
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Line.Weight = 1
    End With

End Sub

Sub ArrowInsert()

    Dim y As Double
    y = Selection.Characters.Last.Information(wdVerticalPositionRelativeToPage)
    
    ActiveDocument.Shapes.AddConnector(msoConnectorStraight, 70.8, y + 28, 70.8 + 35, y + 35 + 28).Select
    
    With Selection
        With .ShapeRange.Line
            .EndArrowheadStyle = msoArrowheadOpen
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
            .Weight = 1.5
        End With
    End With

End Sub