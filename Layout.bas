Attribute VB_Name = "Layout"
Sub Layout()
    Dim obj As Shape
    Dim objGrupo As Shape
    Dim sr As ShapeRange
    Dim s As Shape
    Dim p As page
    Dim paginaDestino As page
    Dim primeira As page
    
    ActiveDocument.BeginCommandGroup "Layout"
    Optimization = True
    
    Set primeira = ActiveDocument.Pages(1)
    Set paginaDestino = ActiveDocument.AddPages(1)
    paginaDestino.SetSize primeira.Shapes.All.SizeWidth, primeira.Shapes.All.SizeHeight
    
    paginaDestino.MoveTo 1
    
    For i = ActiveDocument.Pages.Count To 2 Step -1
        ActiveDocument.Pages(i).Shapes.All.MoveToLayer paginaDestino.ActiveLayer
    Next i
        
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh

End Sub
