Attribute VB_Name = "Ajustar_Margem"
Sub AjustarMargem()
    Dim sr As ShapeRange
    Dim x As Double, y As Double, w As Double, h As Double
    Dim margem As String
    
    ActiveDocument.BeginCommandGroup "Ajustar Margem 0mm"
    Optimization = True
    
    margem = InputBox("Digite a margem a ser definida: ")
    
    ActiveDocument.Unit = cdrMillimeter
    
    If margem <> "" And IsNumeric(margem) Then
    
        Set sr = ActivePage.Shapes.All
            
        For Each p In ActiveDocument.Pages
                
            Set sr = p.Shapes.All
                
            sr.GetBoundingBox x, y, w, h
        
            p.SetSize w + margem, h + margem
                    
            sr.CenterX = p.CenterX
            sr.CenterY = p.CenterY
                
        Next p
    Else
        GoTo Cancelar
    End If
    
Finalizar:
    ActiveDocument.EndCommandGroup
    Optimization = False
    ActiveWindow.Refresh
    Exit Sub
    
Cancelar:
    MsgBox "Ação Cancelada"
    ActiveDocument.EndCommandGroup
    Optimization = False
    ActiveWindow.Refresh
    Exit Sub

End Sub
