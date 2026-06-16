function printFactory() {
    let doc = host.activeDocument;

    if (!doc) {
        alert("Nenhum documento aberto!");
        return;
    }

    doc.beginCommandGroup("Distribuir Objetos");
    host.optimization = true;

    try {
        let pOriginal = doc.activePage;
        let margem = 4;

        // Define a unidade do documento para Milímetros
        doc.unit = 3;

        // Captura todos os shapes da página original
        let sr = pOriginal.shapes.all();
        let totalShapes = sr.count;

        if (totalShapes === 0) {
            alert("Nao ha objetos nesta pagina para distribuir.");
            host.optimization = false;
            doc.endCommandGroup();
            return;
        }

        // Cria o objeto de cor nativo do Corel
        let corSelecionada = host.createColor();

        // Abre a janela nativa do Corel para o usuário escolher uma cor
        // Equivalente ao .UserAssignEx do VBA
        if (corSelecionada.userAssignEx()) {

            // IMPORTANTE: Loop de trás para frente. 
            // Conforme o shape muda de página, a coleção 'sr' diminui.
            for (let i = totalShapes; i >= 1; i--) {
                let s = sr.item(i);

                // Cria 1 nova página no final do documento
                let pNova = doc.addPages(1);

                // Move o objeto para a nova página
                s.moveToLayer(pNova.activeLayer);

                // No Corel JS, CreateRectangle pede: (Left, Top, Right, Bottom)
                // Criando baseado no tamanho do objeto + 6mm
                let larguraSangra = s.sizeWidth + 6;
                let alturaSangra = s.sizeHeight + 6;
                
                let newSangra = pNova.activeLayer.createRectangle(0, alturaSangra, larguraSangra, 0);
                
                // Aplica a cor preenchida uniformemente
                newSangra.fill.applyUniformFill(corSelecionada);
                newSangra.outline.setNoOutline();
                newSangra.orderToBack();

                // Ajusta o tamanho da página para o tamanho do retângulo de sangra
                pNova.setSize(larguraSangra, alturaSangra);

                // Centraliza tudo perfeitamente na nova página
                s.centerX = pNova.centerX;
                s.centerY = pNova.centerY;
                newSangra.centerX = pNova.centerX;
                newSangra.centerY = pNova.centerY;
            }

            // Deleta a página original que agora ficou vazia
            pOriginal.delete();

        } else {
            // Se o usuário cancelar a escolha da cor, roda o bloco sem a sangra colorida
            for (let i = totalShapes; i >= 1; i--) {
                let s = sr.item(i);

                let pNova = doc.addPages(1);
                s.moveToLayer(pNova.activeLayer);

                // Define o tamanho da página apenas com a margem de 1mm
                pNova.setSize(s.sizeWidth + margem, s.sizeHeight + margem);
                s.centerX = pNova.centerX;
                s.centerY = pNova.centerY;
            }

            pOriginal.delete();
        }

    } catch (error) {
        // Bloco de tratamento de erro equivalente ao On Error GoTo
        alert("Ocorreu um erro: " + error.message);
    }

    // Finalização obrigatória do ambiente
    host.optimization = false;
    doc.endCommandGroup();

    if (host.activeWindow) {
        host.activeWindow.refresh();
    }
}

// Executa a macro
printFactory();