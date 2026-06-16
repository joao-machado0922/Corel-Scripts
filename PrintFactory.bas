Attribute VB_Name = "PrintFactory"
Sub PrintFactory()
    Dim pOriginal As page
    Dim pNova As page
    Dim sr As ShapeRange
    Dim margem As Double
    Dim corSelecionada As New Color

    Optimization = True
    ActiveDocument.BeginCommandGroup "Distribuir Objetos"

    On Error GoTo ErroTratado

    Set pOriginal = ActivePage

    margem = 1

    ActiveDocument.Unit = cdrMillimeter

    Set sr = pOriginal.Shapes.All

    If sr.Count = 0 Then
        MsgBox "Nao ha objetos nesta pagina para distribuir."
        GoTo Finalizar
        Exit Sub
    End If

    If corSelecionada.UserAssignEx Then

        For Each s In sr

            Set pNova = ActiveDocument.AddPages(1)

            s.MoveToLayer pNova.ActiveLayer

            Set newSangra = ActiveLayer.CreateRectangle(0, s.SizeHeight + 6, s.SizeWidth + 6, 0)
            newSangra.Fill.ApplyUniformFill corSelecionada

            newSangra.Outline.SetNoOutline

            newSangra.OrderToBack

            pNova.SetSize newSangra.SizeWidth, newSangra.SizeHeight
            s.CenterX = pNova.CenterX
            s.CenterY = pNova.CenterY
            newSangra.CenterX = pNova.CenterX
            newSangra.CenterY = pNova.CenterY

        Next s

        pOriginal.Delete

    Else

        For Each s In sr

            Set pNova = ActiveDocument.AddPages(1)

            s.MoveToLayer pNova.ActiveLayer

            pNova.SetSize s.SizeWidth + margem, s.SizeHeight + margem
            s.CenterX = pNova.CenterX
            s.CenterY = pNova.CenterY

        Next s

        pOriginal.Delete

    End If

Finalizar:
    ActiveDocument.EndCommandGroup
    Optimization = False
    ActiveWindow.Refresh
    Exit Sub

ErroTratado:
    MsgBox "Ocorreu um erro: " & Err.Description
    Resume Finalizar

End Sub

