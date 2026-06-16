function duplicarTodasAsPaginas() {
    let doc = host.activeDocument;

    if (!doc) {
        alert("Nenhum documento aberto!");
        return;
    }

    doc.beginCommandGroup("Duplicar Todas as Páginas");
    host.optimization = true;

    try {
        let paginas = doc.pages;
        let totalOriginal = paginas.count;

        for (let i = totalOriginal; i >= 1; i--) {
            let pOriginal = paginas.item(i);

            let pDuplicada = doc.addPages(1);

            pDuplicada.moveTo(i + 1);

            pDuplicada.setSize(pOriginal.sizeWidth, pOriginal.sizeHeight);

            let shapesOriginal = pOriginal.shapes.all();
            if (shapesOriginal.count > 0) {
                let srCopia = shapesOriginal.duplicate();
                srCopia.moveToLayer(pDuplicada.activeLayer);
            }
        }

    } catch (error) {
        alert("Erro ao duplicar todas as páginas: " + error.message);
    }

    host.optimization = false;
    doc.endCommandGroup();

    if (host.activeWindow) {
        host.activeWindow.refresh();
    }
}

duplicarTodasAsPaginas();