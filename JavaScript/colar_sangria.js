function colarSangria() {
    let doc = host.activeDocument;

    if (!doc) {
        alert("Nenhum documento aberto!");
        return;
    }

    doc.beginCommandGroup("Colar Sangria");
    host.optimization = false;

    let objeto = doc.activeShape;

    if (!objeto) {
        alert("Por favor, selecione o objeto que deseja copiar.");
    } else {
        objeto.copy();
        
        let paginas = doc.pages;
        let paginaAtivaIndex = doc.activePage.index

        for (let i = 1; i <= paginas.count ; i++) {
            if (paginas.item(i).index != paginaAtivaIndex) {
                paginas.item(i).activate();
                let sangria = doc.activeLayer.paste();
                sangria.orderToBack();
            }
        }
        paginas.item(paginaAtivaIndex).activate();
    }

    doc.endCommandGroup();

    if (host.activeWindow) {
        host.activeWindow.refresh();
    }
    
}

colarSangria();