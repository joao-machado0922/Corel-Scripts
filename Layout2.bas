Attribute VB_Name = "Layout2"
Sub Layout2()

    Dim primeira As page
    Dim pagOrigem As page
    Dim copia As ShapeRange
    Dim proximoX As Double
    Dim totalH As Double, totalW As Double

    ActiveDocument.BeginCommandGroup "Layout2"
    Optimization = True

    ActiveDocument.Unit = cdrMillimeter
    
    Set pagOrigem = ActiveDocument.Pages(1)
    Set primeira = ActiveDocument.AddPages(1)
    primeira.SetSize pagOrigem.Shapes.All.SizeWidth, pagOrigem.Shapes.All.SizeHeight
    
    primeira.MoveTo 1
    
    proximoX = primeira.LeftX
    
    For i = 2 To ActiveDocument.Pages.Count
    
        Set copia = ActiveDocument.Pages(i).Shapes.All.Duplicate
        copia.MoveToLayer primeira.ActiveLayer
        
        copia.LeftX = proximoX
        
        copia.CenterY = primeira.CenterY
        
        proximoX = copia.RightX + 10
    
    Next i
    
    If primeira.Shapes.All.Count > 0 Then
        primeira.SetSize primeira.Shapes.All.SizeWidth + 10, primeira.Shapes.All.SizeHeight + 10
        primeira.Shapes.All.CenterX = primeira.CenterX
        primeira.Shapes.All.CenterY = primeira.CenterY
    End If
    
    primeira.Activate
        
    MsgBox "Finalizado"

    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh

End Sub
