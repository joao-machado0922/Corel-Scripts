Attribute VB_Name = "Colar_Sangria"
Sub Colar_Sangria()
    Dim p As page
    
    Optimization = True
    ActiveDocument.BeginCommandGroup "Colar Sangria"
    
    Set s = ActiveShape
    
    s.Copy
    
    If s Is Nothing Then
        MsgBox "Por favor, selecione o objeto que deseja copiar."
        Exit Sub
    End If
    
    For Each p In ActiveDocument.Pages
        If p.Index <> ActivePage.Index Then
            p.Activate
            ActiveLayer.Paste
            ActiveSelection.OrderToBack
        End If
    Next p
    
    ActiveDocument.Pages(1).Activate
    
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    
End Sub
