Attribute VB_Name = "Desagrupar_tudo"
Sub Desagrupar()
    Dim p As page
    Dim srFiltro As ShapeRange
    Dim s As Shape
    
    Optimization = False
    ActiveDocument.BeginCommandGroup "Desagrupar"
    
    For Each p In ActiveDocument.Pages
    
        p.Activate
        
        If p.Shapes.All.Count > 0 Then
        
            p.Shapes.All.Ungroup
            
            DoEvents
        
        End If
        
    Next p
    
    ActiveDocument.Pages(1).Activate
    
    ActiveDocument.EndCommandGroup
    Optimization = False
    ActiveWindow.Refresh
    
End Sub
