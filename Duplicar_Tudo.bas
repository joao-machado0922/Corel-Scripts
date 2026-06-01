Attribute VB_Name = "Duplicar_Tudo"
Sub DuplicarTodasAsPaginas()
    Dim i As Integer
    Dim pOriginal As page
    Dim pDuplicada As page
    Dim sr As ShapeRange, srCopia As ShapeRange
    
    Optimization = True
    ActiveDocument.BeginCommandGroup "Duplicar Páginas"
    
    For i = ActiveDocument.Pages.Count To 1 Step -1
        Set pOriginal = ActiveDocument.Pages(i)
        
        Set pDuplicada = ActiveDocument.AddPages(1)
        pDuplicada.MoveTo i + 1
        
        pDuplicada.SetSize pOriginal.SizeWidth, pOriginal.SizeHeight
        
        If pOriginal.Shapes.All.Count > 0 Then
            
            Set srCopia = pOriginal.Shapes.All.Duplicate
            srCopia.MoveToLayer pDuplicada.ActiveLayer
            
        End If
    Next i
    
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh

End Sub
