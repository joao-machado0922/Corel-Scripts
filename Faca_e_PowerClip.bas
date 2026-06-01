Attribute VB_Name = "Faca_e_PowerClip"
Sub FacaEPowerClip()
    Dim s As Shape
    Dim sFaca As Shape, sContorno As Shape, sArte As Shape
    Dim sr As ShapeRange
    Dim eff As Effect
    Dim nomeCorSpot As String
    
    nomeCorSpot = "CorteTotal"

    Optimization = True
    ActiveDocument.BeginCommandGroup "Faca e PowerClip"
    
    ' 1. Procura a forma que tem o contorno com a cor SPOT
    ' Filtramos apenas shapes que possuem contorno (Outline)
    For Each s In ActivePage.Shapes.FindShapes()
        If s.Outline.Type <> cdrNoOutline Then
            If s.Outline.Color.IsSpot Then
                If s.Outline.Color.SpotName = nomeCorSpot Then
                    Set sFaca = s
                    Exit For
                End If
            End If
        End If
    Next s

    If sFaca Is Nothing Then
        MsgBox "Nenhum objeto com o contorno SPOT '" & nomeCorSpot & "' foi encontrado."
        GoTo Finalize
    End If

    ' 2. Identificar a "Arte" (o que não é a faca e será colocado dentro do PowerClip)
    ' Aqui pegamos o que estiver selecionado além da faca, ou o maior objeto da página
    Set sr = ActivePage.Shapes.All
    For Each s In sr
        If s.StaticID <> sFaca.StaticID Then
            Set sArte = s ' Define o outro objeto como a arte
            Exit For
        End If
    Next s
    
    If sArte Is Nothing Then
        MsgBox "Não foi encontrado um objeto de 'Arte' para o PowerClip."
        GoTo Finalize
    End If

    ' 3. Aplicar a ferramenta Contorno (Contour)
    ' Exemplo: 2mm para fora, 1 etapa
    Set eff = sFaca.CreateContour(cdrContourOutside, 3, 1) ' 2mm de deslocamento
    
    ' 4. Separar a Curva (Break Apart)
    Set sr = eff.Separate
    Set sContorno = sr(1) ' O novo objeto criado pelo contorno
    
    ' 5. Limpar o novo objeto (remover contorno e dar um preenchimento temporário)
    sContorno.Fill.ApplyUniformFill CreateRGBColor(255, 255, 255)
    sContorno.Outline.SetNoOutline
    
    ' 6. Aplicar PowerClip: Coloca a Arte dentro da forma criada pelo Contorno
    sArte.AddToPowerClip sContorno
    
    ' Opcional: Mover a faca original para frente para visualização do corte
    sFaca.OrderToFront

Finalize:
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
End Sub
