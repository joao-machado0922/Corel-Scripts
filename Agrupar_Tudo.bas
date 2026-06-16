Attribute VB_Name = "Agrupar_Tudo"
Sub Agrupar()
    Dim p As page
    Dim srFiltro As ShapeRange
    Dim s As Shape
    
    Optimization = False
    ActiveDocument.BeginCommandGroup "Agrupar"
    
    For Each p In ActiveDocument.Pages
    
        p.Activate
        
        Set srFiltro = CreateShapeRange
        
        For Each s In p.Shapes.All
        
            If s.Type <> cdrGuidelineShape And s.Locked = False Then
            
                srFiltro.Add s
            
            End If
        
        Next s
        
        If srFiltro.Count > 1 Then
        
            srFiltro.Group
            
            DoEvents
        
        End If
        
    Next p
    
    ActiveDocument.Pages(1).Activate
    
    ActiveDocument.EndCommandGroup
    Optimization = False
    ActiveWindow.Refresh
    
End Sub

