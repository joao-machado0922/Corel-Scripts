function ajustarMargem() {
    let doc = host.activeDocument;
    if (!doc) {
        alert("Nenhum documento aberto!");
        return;
    }

    doc.beginCommandGroup("Ajustar Margem JavaScript");
    host.optimization = true;

    doc.unit = 3;

    let paginas = doc.pages;

    for (let i = 1; i <= paginas.count; i++) {
        let p = paginas.item(i);
        let sr = p.shapes.all();

        if (sr.count > 0) {

            p.setSize(p.shapes.all().sizeWidth - 6, p.shapes.all().sizeHeight - 6);

            sr.centerX = p.centerX;
            sr.centerY = p.centerY;
        }
    }

    doc.endCommandGroup();
    
    host.optimization = false;

    if (host.activeWindow) {
        host.activeWindow.refresh();
    }
}

ajustarMargem();