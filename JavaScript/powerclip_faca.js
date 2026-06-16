function powerClipFaca() {

    let doc = host.activeDocument;

    if (!doc) {
        alert("Nenhum documento aberto!");
        return;
    }

    doc.beginCommandGroup("Colar Faca");
    host.optimization = true;

    let paginas = doc.pages;

    for (let i = 1; i <= paginas.count; i++) {
        let p = paginas.item(i);

        if (p.shapes.count === 2) {
            p.shapes.item(2).addToPowerClip(p.shapes.item(1));
        }
    }

    host.optimization = false;
    doc.endCommandGroup();

    if (host.activeWindow) {
        host.activeWindow.refresh();
    }

}

powerClipFaca()