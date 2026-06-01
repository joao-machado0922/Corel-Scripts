Attribute VB_Name = "Deletar_Paginas"
Sub DeletarOutros()

    Dim page As page

    ActiveDocument.BeginCommandGroup "Deletar"
    Optimization = True
    
    Set page = ActivePage
    
    For Each p In ActiveDocument.Pages
    
        If p.Index <> page.Index Then
            p.Delete
        End If
    
    Next p
    
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
End Sub
