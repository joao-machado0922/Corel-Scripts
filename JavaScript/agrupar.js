function agruparTudo() {
    let doc = host.activeDocument;

    if (!doc) {
        alert("Nenhum documento aberto!");
        return;
    }

    doc.beginCommandGroup("Agrupar");
    host.optimization = false;

    let paginas = doc.pages;

    for (let i = 1; i <= paginas.count; i++) {
        let p = paginas.item(i);
        p.activate();

        let srFiltro = host.createShapeRange();
        
        let todosShapes = p.shapes.all();

        for (let j = 1; j <= todosShapes.count; j++) {
            let s = todosShapes.item(j);

            if (s.type !== 11 && s.locked === false) {
                srFiltro.add(s);
            }
        }

        if (srFiltro.count > 1) {
            srFiltro.group();
        }
    }

    doc.pages.item(1).activate();

    doc.endCommandGroup();

    if (host.activeWindow) {
        host.activeWindow.refresh();
    }
    
}

agruparTudo();