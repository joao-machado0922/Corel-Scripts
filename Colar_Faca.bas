Attribute VB_Name = "Colar_Faca"
Sub Colar_Faca()
    Dim p As page
    Dim s As Shape
    
    Optimization = True
    ActiveDocument.BeginCommandGroup "Colar Faca"
    
    Set s = ActiveShape
    
    If s Is Nothing Then
        MsgBox "Por favor, selecione o objeto que deseja copiar."
        Exit Sub
    End If
    
    s.Copy
    
    For Each p In ActiveDocument.Pages
        If p.Index <> ActivePage.Index Then
            p.Activate
            ActiveLayer.Paste
            ActiveSelection.OrderToFront
        End If
    Next p
    
    ActiveDocument.Pages(1).Activate
    
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    
End Sub

