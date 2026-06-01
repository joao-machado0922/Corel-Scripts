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
        paginas[i].activate;

        let shapes = [];

        for (s in paginas[i].shapes.all()) {
            shapes.push(s);
        }

        if (shapes.length > 1) {
            shapes.group();
        }

    }

    doc.pages.item(1).activate();

    doc.endCommandGroup();

    if (host.activeWindow) {
        host.activeWindow.refresh();
    }
}

agruparTudo()