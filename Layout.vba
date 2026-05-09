Sub Layout()
    Dim obj As Shape
    Dim sr As ShapeRange
    Dim x As Double, y As Double, w As Double, h As Double
    Dim margem As Double
    Dim paginaDestino As Page

    Optimization = True
    ActiveDocument.BeginCommandGroup "Layout"

    Set paginaDestino = ActiveDocument.Pages(1)

    paginaDestino.Shapes.All.Group

    For i = ActiveDocument.Pages.Count To 2 Step -1
        Dim p As Page
        Set p = ActiveDocument.Pages(i)

        If p.Shapes.All.Count > 0 Then

            Set obj = p.Shapes.All.Group

            obj.MoveToLayer paginaDestino.ActiveLayer

            p.Delete

        End If
    Next i

    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh

End Sub