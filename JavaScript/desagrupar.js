function desagruparTodasAsPaginas() {
    var doc = host.activeDocument;

    if (!doc) {
        alert("Nenhum documento aberto!");
        return;
    }

    doc.beginCommandGroup("Desagrupar");

    host.optimization = false;

    var paginas = doc.pages;

    for (var i = 1; i <= paginas.count; i++) {
        var p = paginas.item(i);

        p.activate();

        var todosShapes = p.shapes.all();

        if (todosShapes.count > 0) {
            todosShapes.ungroup();
        }
    }

    doc.pages.item(1).activate();

    doc.endCommandGroup();

    if (host.activeWindow) {
        host.activeWindow.refresh();
    }
}

desagruparTodasAsPaginas();