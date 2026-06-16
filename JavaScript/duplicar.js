function duplicarPaginaAtivaNativo() {
    let doc = host.activeDocument;
    if (!doc) return;

    doc.beginCommandGroup("Duplicar Página Nativo");
    host.optimization = true;

    try {
        let pOriginal = doc.activePage;

        // O comando mágico. O número 2 significa "cdrCopyEverything"
        // Ele cria a cópia exata logo após a página atual automaticamente
        pOriginal.duplicate(2); 

    } catch (error) {
        alert("Erro: " + error.message);
    }

    host.optimization = false;
    doc.endCommandGroup();
    if (host.activeWindow) host.activeWindow.refresh();
}

duplicarPaginaAtivaNativo();