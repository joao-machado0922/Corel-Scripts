Attribute VB_Name = "Textos_to_Curvas"
Sub TextosEmCurvas()
    Dim p As page
    Dim textos As ShapeRange
    Dim t As Shape
    Dim totalConvertido As Long
    
    ActiveDocument.BeginCommandGroup "Converter Textos em Curvas"
    Optimization = True
    
    
    For Each p In ActiveDocument.Pages
        
        ' Filtra e seleciona apenas objetos do tipo Texto (Type = 3) nesta página
        ' O True no final faz ele buscar também dentro de grupos
        Set textos = p.Shapes.FindShapes(Type:=cdrTextShape, Recursive:=True)
        
        If textos.Count > 0 Then
            
            For Each t In textos
                If t.Locked = False Then
                    t.ConvertToCurves
                End If
            Next t
            
            DoEvents
            
        End If
    Next p
    
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
End Sub
