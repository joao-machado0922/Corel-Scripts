function layout() {
    let doc = host.activeDocument;
    if (!doc) {
        alert("Nenhum documento aberto!");
        return;
    }

    doc.beginCommandGroup("Layout");
    host.optimization = true;

    let paginas = doc.pages;
    let paginaDestino = paginas.item(1);

    for (let i = 2; i <= paginas.count; i++) {
        let p = paginas.item(i);
        if (p.shapes.all().count > 0) {
            p.shapes.all().moveToLayer(paginaDestino.activeLayer);
        }
    }

    host.optimization = false;
    doc.endCommandGroup();
    
    if (host.activeWindow) {
        host.activeWindow.refresh();
    }

}

layout();