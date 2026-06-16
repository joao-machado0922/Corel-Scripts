Sub Duplicar()
    Dim i As Integer
    Dim pOriginal As page
    Dim pDuplicada As page
    Dim sr As ShapeRange, srCopia As ShapeRange
    
    Optimization = True
    ActiveDocument.BeginCommandGroup "Duplicar P�gina"
    
    Set pOriginal = ActiveDocument.Pages(ActivePage.Index)
        
    Set pDuplicada = ActiveDocument.AddPages(1)
    
    pDuplicada.MoveTo pOriginal.Index + 1
        
    pDuplicada.SetSize pOriginal.SizeWidth, pOriginal.SizeHeight
        
    If pOriginal.Shapes.All.Count > 0 Then
            
        Set srCopia = pOriginal.Shapes.All.Duplicate
        srCopia.MoveToLayer pDuplicada.ActiveLayer
            
    End If
    
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh

End Sub

