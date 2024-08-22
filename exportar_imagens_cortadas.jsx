// Variável global para armazenar a pasta de destino
var outputFolder = null;

// Função principal que processa todos os documentos abertos
function processAllOpenDocuments() {
    var initialDocument = app.activeDocument; // Documento inicialmente ativo
    var documents = app.documents;

    for (var i = 0; i < documents.length; i++) {
        app.activeDocument = documents[i];
        processDocument(documents[i]);
    }

    app.activeDocument = initialDocument; // Retorna ao documento inicial
    alert("Exportação concluída!");
}

// Processa cada documento individualmente
function processDocument(doc) {
    var projectName = doc.name.replace(/\.[^\.]+$/, ''); // Remove a extensão do nome do projeto

    // Seleciona a pasta de destino apenas na primeira vez
    if (!outputFolder) {
        outputFolder = Folder.selectDialog("Selecione a pasta para salvar as imagens exportadas:");
    }

    if (outputFolder) {
        var layers = doc.layers;
        for (var i = 0; i < layers.length - 1; i++) { // Exclui a última camada
            var layer = layers[i];
            if (layer.typename == "ArtLayer" && layer.visible) {
                exportLayerAsPNG(layer, projectName, outputFolder);
            }
        }
    }
}

// Exporta cada camada visível como PNG
function exportLayerAsPNG(layer, projectName, outputFolder) {
    var layerName = layer.name.replace(/[\/\\\:\*\?\"\<\>\|]/g, "_"); // Limpa o nome da camada

    // Duplica o documento atual para preservar o original
    var tempDoc = app.activeDocument.duplicate();

    // Oculta todas as camadas na cópia temporária, exceto a camada atual
    hideAllLayers(tempDoc);
    tempDoc.artLayers.getByName(layer.name).visible = true;

    tempDoc.trim(TrimType.TRANSPARENT); // Aplica o trim para remover áreas transparentes

    // Define o caminho e nome do arquivo de saída
    var fileName = projectName + "_" + layerName + ".png";
    var filePath = new File(outputFolder + "/" + fileName);

    // Define as opções de exportação para PNG
    var pngOptions = new PNGSaveOptions();
    pngOptions.compression = 9;
    tempDoc.saveAs(filePath, pngOptions, true, Extension.LOWERCASE);

    tempDoc.close(SaveOptions.DONOTSAVECHANGES); // Fecha o documento temporário sem salvar
}

// Oculta todas as camadas no documento ou grupo
function hideAllLayers(parent) {
    for (var i = 0; i < parent.layers.length; i++) {
        parent.layers[i].visible = false;
    }
}

// Executa o script para todos os documentos abertos
processAllOpenDocuments();
