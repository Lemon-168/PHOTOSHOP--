#target photoshop

function main() {
    var doc = app.activeDocument;

    // 选择 CSV 文件
    var csvFile = File.openDialog("请选择包含人名的 CSV 文件", "*.csv");
    if (!csvFile) {
        alert("未选择 CSV 文件，无法运行。");
        return;
    }

    // 读取 CSV 文件中的人名（A列，从上往下）
    var names = [];
    if (csvFile.open("r")) {
        while (!csvFile.eof) {
            var line = csvFile.readln();
line = line.replace(/^\s+|\s+$/g, "");

            if (line !== "") {
                names.push(line);
            }
        }
        csvFile.close();
    } else {
        alert("无法打开 CSV 文件。");
        return;
    }

    // PS所有可用的字体，目前应该让开始就指向中文适用字体
    var fonts = app.fonts;
    var fontList = [];
    for (var i = 0; i < fonts.length; i++) {
        fontList.push(fonts[i].name);
    }


    var win = new Window("dialog", "邀请函批量-文本参数-暮光精灵");
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];

    var fontGroup = win.add("group");
    fontGroup.add("statictext", undefined, "字体名称:");
    var fontDropdown = fontGroup.add("dropdownlist", undefined, fontList);
    fontDropdown.selection = 0;

    var sizeGroup = win.add("group");
    sizeGroup.add("statictext", undefined, "字体大小:");
    var sizeInput = sizeGroup.add("edittext", undefined, "108");

    var colorGroup = win.add("group");
    colorGroup.add("statictext", undefined, "文本颜色:");
    var colorBtn = colorGroup.add("button", undefined, "选择颜色");

    // 文本位置（临时测试）取左下角为目的点（非左上角）
    var posGroup = win.add("group");
    posGroup.add("statictext", undefined, "文本位置 (X,Y):");
    var xInput = posGroup.add("edittext", undefined, "200");
    posGroup.add("statictext", undefined, ",");
    var yInput = posGroup.add("edittext", undefined, "200");

    var styleGroup = win.add("group");
    styleGroup.add("statictext", undefined, "文本样式:");
    var boldCheckbox = styleGroup.add("checkbox", undefined, "粗体");
    var italicCheckbox = styleGroup.add("checkbox", undefined, "斜体");

    var exportGroup = win.add("group");
    exportGroup.add("statictext", undefined, "导出路径:");
    var exportInput = exportGroup.add("edittext", undefined, Folder.desktop.absoluteURI + "/Images");
    var exportBtn = exportGroup.add("button", undefined, "浏览...");

    var btnGroup = win.add("group");
    var cancelBtn = btnGroup.add("button", undefined, "取消");
    var okBtn = btnGroup.add("button", undefined, "确定");

    exportBtn.onClick = function () {
        var folder = Folder.selectDialog("选择导出目录");
        if (folder) exportInput.text = folder.absoluteURI;
    };

    var selectedColor = new SolidColor();
    selectedColor.rgb.red = 255;
    selectedColor.rgb.green = 0;
    selectedColor.rgb.blue = 0;

    colorBtn.onClick = function () {
        app.showColorPicker();
        var newColor = app.foregroundColor;
        selectedColor.rgb.red = newColor.rgb.red;
        selectedColor.rgb.green = newColor.rgb.green;
        selectedColor.rgb.blue = newColor.rgb.blue;
        colorBtn.graphics.backgroundColor = colorBtn.graphics.newBrush(
            colorBtn.graphics.BrushType.SOLID_COLOR,
            [newColor.rgb.red / 255, newColor.rgb.green / 255, newColor.rgb.blue / 255, 1]
        );
    };

    okBtn.onClick = function () {
        var fontName = fontDropdown.selection.text;
        var fontSize = parseInt(sizeInput.text);
        var xPos = parseInt(xInput.text);
        var yPos = parseInt(yInput.text);
        var exportPath = exportInput.text;
        var isBold = boldCheckbox.value;
        var isItalic = italicCheckbox.value;

        var exportFolder = new Folder(exportPath);
        if (!exportFolder.exists) {
            exportFolder.create();
        }

        var exportOptions = new ExportOptionsSaveForWeb();
        exportOptions.format = SaveDocumentType.PNG;
        exportOptions.PNG8 = false;
        exportOptions.transparency = true;

        for (var i = 0; i < names.length; i++) {
            var textLayer = doc.artLayers.add();
            textLayer.kind = LayerKind.TEXT;

            var textItem = textLayer.textItem;
            textItem.contents = names[i];
            textItem.font = fontName;
            textItem.size = fontSize;
            textItem.color = selectedColor;
            textItem.position = [xPos, yPos];
            textItem.fauxBold = isBold;
            textItem.fauxItalic = isItalic;

            var exportFile = new File(exportPath + "/客人-" + names[i] + ".png");
            doc.exportDocument(exportFile, ExportType.SAVEFORWEB, exportOptions);

            textLayer.visible = false;
        }

        alert("导出完成！");
        win.close();
    };

    cancelBtn.onClick = function () {
        win.close();
    };

    win.center();
    win.show();
}

main();
