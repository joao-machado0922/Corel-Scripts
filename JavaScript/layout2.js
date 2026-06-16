function layout2() {
    let doc = host.activeDocument;
    if (!doc) {
        alert("Nenhum documento aberto!");
        return;
    }

    doc.beginCommandGroup("Layout");
    host.optimization = false;

    doc.unit = 3;

    let paginas = doc.pages;
    let pagOrigem = paginas.item(1);
    
    let primeira = doc.addPages(1);
    primeira.moveTo(1);
    
    if (pagOrigem.shapes.all().count > 0) {
        primeira.setSize(pagOrigem.shapes.all().sizeWidth, pagOrigem.shapes.all().sizeHeight);
    }
    
    let proximoX = primeira.leftX;

    for (let i = 2; i <= paginas.count; i++) {
        let p = paginas.item(i);
        p.activate();

        if (p.shapes.all().count > 0) {
            let copia = p.shapes.all().duplicate();
            
            copia.moveToLayer(primeira.activeLayer);
            
            copia.leftX = proximoX;
            copia.centerY = primeira.centerY;
            
            proximoX = copia.rightX + 10;
        }
    }

    let todosShapesPrimeira = primeira.shapes.all();
    if (todosShapesPrimeira.count > 0) {
        
        let x = 0, y = 0, totalW = 0, totalH = 0;
        
        todosShapesPrimeira.getBoundingBox(x, y, totalW, totalH, true);
        
        primeira.setSize(proximoX - primeira.leftX, todosShapesPrimeira.sizeHeight);
        
        todosShapesPrimeira.centerX = primeira.centerX;
        todosShapesPrimeira.centerY = primeira.centerY;
    }

    primeira.activate();

    host.optimization = false;
    doc.endCommandGroup();
    
    if (host.activeWindow) {
        host.activeWindow.refresh();
    }

}

layout2();